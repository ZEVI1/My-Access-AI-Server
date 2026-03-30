using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using MS.Access.MCP.Interop;

class Program
{
    private static void RegisterProcessEvents(AccessWorker accessWorker)
    {
        AppDomain.CurrentDomain.ProcessExit += (sender, args) =>
        {
            FileLogger.Log("ProcessExit: cleaning up Access worker and file logger.");
            try { accessWorker.Dispose(); } catch { }
            FileLogger.Shutdown();
        };

        Console.CancelKeyPress += (sender, args) =>
        {
            FileLogger.Log("CancelKeyPress: cleaning up Access worker and file logger.");
            try { accessWorker.Dispose(); } catch { }
            FileLogger.Shutdown();
        };
    }

    private sealed class RpcRequest
    {
        public string Message { get; }
        public DateTime ReceivedAt { get; }
        public DateTime StartedAt { get; set; }
        public bool HasResponseSent { get; set; }
        public bool TimedOut { get; set; }

        public RpcRequest(string message, DateTime receivedAt)
        {
            Message = message;
            ReceivedAt = receivedAt;
            StartedAt = DateTime.MinValue;
            HasResponseSent = false;
            TimedOut = false;
        }
    }

    private sealed class AccessWorker : IDisposable
    {
        private readonly BlockingCollection<RpcRequest> _requestQueue = new(new ConcurrentQueue<RpcRequest>());
        private readonly BlockingCollection<string> _responseQueue;
        private readonly CancellationTokenSource _internalCancellation = new();
        private readonly TimeSpan _requestTimeout = TimeSpan.FromSeconds(30);
        private readonly object _currentLock = new();
        private RpcRequest? _currentRequest;
        private readonly Thread _workerThread;
        private readonly Task _monitorTask;
        private bool _disposed;

        public AccessWorker(BlockingCollection<string> responseQueue, CancellationToken externalCancellationToken)
        {
            _responseQueue = responseQueue;
            _workerThread = new Thread(Run)
            {
                IsBackground = true,
                Name = "AccessWorker"
            };
            _workerThread.SetApartmentState(ApartmentState.STA);

            _workerThread.Start();
            _monitorTask = Task.Run(() => MonitorTimeoutsAsync(externalCancellationToken));
        }

        public void EnqueueRequest(string message)
        {
            if (_disposed)
                return;

            try
            {
                _requestQueue.Add(new RpcRequest(message, DateTime.UtcNow));
            }
            catch (InvalidOperationException)
            {
                // The queue has been completed.
            }
        }

        private void Run()
        {
            using var accessService = new AccessInteropService();
            try
            {
                foreach (var request in _requestQueue.GetConsumingEnumerable(_internalCancellation.Token))
                {
                    ProcessRequest(accessService, request);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                FileLogger.Log("Error", "AccessWorkerFailure", new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

        private void ProcessRequest(AccessInteropService accessService, RpcRequest request)
        {
            lock (_currentLock)
            {
                _currentRequest = request;
                request.StartedAt = DateTime.UtcNow;
                request.HasResponseSent = false;
                request.TimedOut = false;
            }

            try
            {
                var response = ProcessRpcMessage(accessService, request.Message);
                if (response == null)
                    return;

                lock (_currentLock)
                {
                    if (request.HasResponseSent)
                        return;

                    request.HasResponseSent = true;
                }

                var jsonResponse = JsonSerializer.Serialize(response, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
                _responseQueue.Add(jsonResponse);
                FileLogger.Log("Information", "RpcResponseSent", new { requestId = response.Id, durationMs = (DateTime.UtcNow - request.StartedAt).TotalMilliseconds, method = ExtractMethod(request.Message) });
            }
            catch (Exception ex)
            {
                FileLogger.Log("Error", "RpcProcessingError", new { error = ex.Message, stackTrace = ex.StackTrace });

                lock (_currentLock)
                {
                    if (request.HasResponseSent)
                        return;

                    request.HasResponseSent = true;
                }

                var errorResponse = CreateErrorResponse(ExtractId(request.Message), -32603, "Internal error", ex.Message);
                var jsonError = JsonSerializer.Serialize(errorResponse, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
                _responseQueue.Add(jsonError);
            }
            finally
            {
                lock (_currentLock)
                {
                    _currentRequest = null;
                }
            }
        }

        private async Task MonitorTimeoutsAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && !_requestQueue.IsCompleted)
            {
                await Task.Delay(500, cancellationToken).ConfigureAwait(false);

                RpcRequest? currentRequest;
                lock (_currentLock)
                {
                    currentRequest = _currentRequest;
                }

                if (currentRequest != null &&
                    !currentRequest.TimedOut &&
                    !currentRequest.HasResponseSent &&
                    DateTime.UtcNow - currentRequest.StartedAt > _requestTimeout)
                {
                    lock (_currentLock)
                    {
                        currentRequest.TimedOut = true;
                        currentRequest.HasResponseSent = true;
                    }

                    var id = ExtractId(currentRequest.Message);
                    var timeoutResponse = CreateErrorResponse(id, -32000, $"Request timed out after {_requestTimeout.TotalSeconds:N0} seconds");
                    var jsonResponse = JsonSerializer.Serialize(timeoutResponse, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
                    _responseQueue.Add(jsonResponse);
                    FileLogger.Log("Warning", "RpcRequestTimeout", new { requestId = id, method = ExtractMethod(currentRequest.Message), durationMs = (DateTime.UtcNow - currentRequest.StartedAt).TotalMilliseconds });
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            _disposed = true;
            _requestQueue.CompleteAdding();
            _internalCancellation.Cancel();

            try { _workerThread.Join(2000); } catch { }
            try { _monitorTask.Wait(1000); } catch { }
        }
    }

    private static string? ExtractMethod(string message)
    {
        try
        {
            using var document = JsonDocument.Parse(message);
            if (document.RootElement.TryGetProperty("method", out var methodElement) && methodElement.ValueKind == JsonValueKind.String)
                return methodElement.GetString();
        }
        catch { }

        return null;
    }

    static async Task Main(string[] args)
    {
        Console.OutputEncoding = new System.Text.UTF8Encoding(false);
        Console.InputEncoding = System.Text.Encoding.UTF8;

        AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
        {
            var message = $"AssemblyResolve requested for: {args.Name}";
            FileLogger.Log(message);
            return null;
        };
        
        var cancellationSource = new CancellationTokenSource();
        using var stdin = Console.OpenStandardInput();
        using var responseQueue = new BlockingCollection<string>(new ConcurrentQueue<string>());
        using var accessWorker = new AccessWorker(responseQueue, cancellationSource.Token);
        RegisterProcessEvents(accessWorker);

        var writerTask = Task.Run(() => ResponseWriterLoop(responseQueue, cancellationSource.Token));

        try
        {
            while (true)
            {
                var message = await ReadRpcMessageAsync(stdin, cancellationSource.Token);
                if (message == null)
                    break;

                if (string.IsNullOrWhiteSpace(message))
                    continue;

                FileLogger.Log("Information", "RpcRequestReceived", new { raw = message });
                accessWorker.EnqueueRequest(message);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            FileLogger.Log("Error", "FatalServerError", new { error = ex.Message, stackTrace = ex.StackTrace });
            var response = CreateErrorResponse(null, -32603, "Internal error", ex.Message);
            var jsonResponse = JsonSerializer.Serialize(response, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
            responseQueue.Add(jsonResponse);
            responseQueue.CompleteAdding();
            await writerTask.ConfigureAwait(false);
            Environment.Exit(1);
        }
        finally
        {
            accessWorker.Dispose();
            responseQueue.CompleteAdding();
            await writerTask.ConfigureAwait(false);
            FileLogger.Shutdown();
        }
    }

    static void ResponseWriterLoop(BlockingCollection<string> responseQueue, CancellationToken cancellationToken)
    {
        try
        {
            foreach (var response in responseQueue.GetConsumingEnumerable(cancellationToken))
            {
                Console.WriteLine(response);
                Console.Out.Flush();
            }
        }
        catch (OperationCanceledException) { }
    }

    static JsonRpcResponse? ProcessRpcMessage(AccessInteropService accessService, string message)
    {
        try
        {
            using var document = JsonDocument.Parse(message);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return CreateErrorResponse(null, -32600, "Invalid request");

            var id = ExtractId(root);
            var method = root.TryGetProperty("method", out var methodElement) && methodElement.ValueKind == JsonValueKind.String
                ? methodElement.GetString()
                : null;

            if (string.IsNullOrEmpty(method))
                return CreateErrorResponse(id, -32600, "Method is required");

            if (method.StartsWith("notifications/") || !root.TryGetProperty("id", out _))
            {
                FileLogger.Log($"Ignored notification: {method}");
                return null;
            }

            var paramsElement = root.TryGetProperty("params", out var p) ? p : default;
            object result;

            try
            {
                result = method switch
                {
                    "initialize" => HandleInitialize(),
                    "ping" => new { },
                    "tools/list" => HandleToolsList(),
                    "tools/call" => HandleToolsCall(accessService, paramsElement),
                    _ => throw new InvalidOperationException($"Unknown method: {method}")
                };
            }
            catch (ArgumentException ex)
            {
                return CreateErrorResponse(id, -32602, ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                return CreateErrorResponse(id, -32603, ex.Message);
            }
            catch (Exception ex)
            {
                return CreateErrorResponse(id, -32603, ex.Message);
            }

            return new JsonRpcResponse
            {
                Id = id,
                Result = result,
                Error = null
            };
        }
        catch (JsonException ex)
        {
            FileLogger.Log($"JSON parse error: {ex}");
            return CreateErrorResponse(null, -32700, "Parse error", ex.Message);
        }
    }

    static object? ExtractId(string message)
    {
        try
        {
            using var document = JsonDocument.Parse(message);
            return ExtractId(document.RootElement);
        }
        catch
        {
            return null;
        }
    }

    static object? ExtractId(JsonElement root)
    {
        if (!root.TryGetProperty("id", out var idElement))
            return null;

        return idElement.ValueKind switch
        {
            JsonValueKind.Number => idElement.TryGetInt32(out var intId) ? intId : idElement.GetInt64(),
            JsonValueKind.String => idElement.GetString(),
            JsonValueKind.Null => null,
            _ => idElement.ToString()
        };
    }

    static object? JsonElementToObject(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetInt64(out var longValue) ? longValue : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            JsonValueKind.Array => element.EnumerateArray().Select(JsonElementToObject).ToArray(),
            JsonValueKind.Object => element.EnumerateObject().ToDictionary(p => p.Name, p => JsonElementToObject(p.Value)),
            _ => element.ToString()
        };
    }

    static JsonRpcResponse CreateErrorResponse(object? id, int code, string message, object? data = null)
    {
        return new JsonRpcResponse
        {
            Id = id,
            Result = null,
            Error = new JsonRpcError
            {
                Code = code,
                Message = message,
                Data = data
            }
        };
    }

    static async Task<string?> ReadRpcMessageAsync(Stream stream, CancellationToken cancellationToken)
    {
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var firstLine = await ReadLineAsync(stream, cancellationToken);
        if (firstLine == null)
            return null;

        if (firstLine.TrimStart().StartsWith("{"))
            return firstLine;

        if (firstLine.Length > 0)
        {
            var headerParts = firstLine.Split(':', 2);
            if (headerParts.Length == 2)
                headers[headerParts[0].Trim()] = headerParts[1].Trim();
        }

        string? line;
        while (!string.IsNullOrWhiteSpace(line = await ReadLineAsync(stream, cancellationToken)))
        {
            var headerParts = line.Split(':', 2);
            if (headerParts.Length == 2)
                headers[headerParts[0].Trim()] = headerParts[1].Trim();
        }

        if (!headers.TryGetValue("Content-Length", out var contentLengthString) || !int.TryParse(contentLengthString, out var contentLength))
            return null;

        var buffer = new byte[contentLength];
        var offset = 0;
        while (offset < contentLength)
        {
            var read = await stream.ReadAsync(buffer, offset, contentLength - offset, cancellationToken);
            if (read == 0)
                break;
            offset += read;
        }

        return Encoding.UTF8.GetString(buffer, 0, offset);
    }

    static async Task<string?> ReadLineAsync(Stream stream, CancellationToken cancellationToken)
    {
        var bytes = new List<byte>();
        var buffer = new byte[1];

        while (true)
        {
            var read = await stream.ReadAsync(buffer, 0, 1, cancellationToken);
            if (read == 0)
            {
                if (bytes.Count == 0)
                    return null;
                break;
            }

            if (buffer[0] == '\n')
                break;

            if (buffer[0] != '\r')
                bytes.Add(buffer[0]);
        }

        return Encoding.UTF8.GetString(bytes.ToArray());
    }

    static object HandleInitialize()
    {
        return new
        {
            protocolVersion = "2024-11-05",
            capabilities = new { tools = new { } },
            serverInfo = new
            {
                name = "Access MCP Server",
                version = "1.0.0"
            }
        };
    }

    static object HandleToolsList()
    {
        return new
        {
            tools = new object[]
            {
                new { name = "connect_access", description = "Connect to an Access database", inputSchema = new { type = "object", properties = new { database_path = new { type = "string", description = "Path to the Access database file (.accdb or .mdb)" } }, required = new string[] { "database_path" } } },
                new { name = "disconnect_access", description = "Disconnect from the current Access database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "is_connected", description = "Check if connected to an Access database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_tables", description = "Get list of all tables in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_queries", description = "Get list of all queries in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_relationships", description = "Get list of all relationships in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "create_table", description = "Create a new table in the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" }, fields = new { type = "array", items = new { type = "object", properties = new { name = new { type = "string" }, type = new { type = "string" }, size = new { type = "integer" }, required = new { type = "boolean" }, allow_zero_length = new { type = "boolean" } } } } }, required = new string[] { "table_name", "fields" } } },
                new { name = "delete_table", description = "Delete a table from the database", inputSchema = new { type = "object", properties = new { table_name = new { type = "string" } }, required = new string[] { "table_name" } } },
                new { name = "launch_access", description = "Launch Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "close_access", description = "Close Microsoft Access application", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_forms", description = "Get list of all forms in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_reports", description = "Get list of all reports in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_macros", description = "Get list of all macros in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_modules", description = "Get list of all modules in the database", inputSchema = new { type = "object", properties = new { } } },
                new { name = "open_form", description = "Open a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "close_form", description = "Close a form in Access", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_vba_projects", description = "Get list of VBA projects", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_vba_code", description = "Get VBA code from a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" } }, required = new string[] { "project_name", "module_name" } } },
                new { name = "set_vba_code", description = "Set VBA code in a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, code = new { type = "string" } }, required = new string[] { "project_name", "module_name", "code" } } },
                new { name = "add_vba_procedure", description = "Add a VBA procedure to a module", inputSchema = new { type = "object", properties = new { project_name = new { type = "string" }, module_name = new { type = "string" }, procedure_name = new { type = "string" }, code = new { type = "string" } }, required = new string[] { "project_name", "module_name", "procedure_name", "code" } } },
                new { name = "compile_vba", description = "Compile VBA code", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_system_tables", description = "Get list of system tables", inputSchema = new { type = "object", properties = new { } } },
                new { name = "get_object_metadata", description = "Get metadata for database objects", inputSchema = new { type = "object", properties = new { } } },
                new { name = "execute_sql", description = "Execute raw SQL against the connected database", inputSchema = new { type = "object", properties = new { sql = new { type = "string" }, mode = new { type = "string", @enum = new string[] { "select", "nonquery", "scalar" } }, parameters = new { type = "array", items = new { type = "object" } } }, required = new string[] { "sql" } } },
                new { name = "refresh_schema_cache", description = "Refresh cached schema metadata", inputSchema = new { type = "object", properties = new { } } },
                new { name = "health_check", description = "Verify server health and status", inputSchema = new { type = "object", properties = new { } } },
                new { name = "server_info", description = "Get server information", inputSchema = new { type = "object", properties = new { } } },
                new { name = "form_exists", description = "Check if a form exists", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_form_controls", description = "Get list of controls in a form", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "get_control_properties", description = "Get properties of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" } }, required = new string[] { "form_name", "control_name" } } },
                new { name = "set_control_property", description = "Set a property of a control", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" }, control_name = new { type = "string" }, property_name = new { type = "string" }, value = new { type = "string" } }, required = new string[] { "form_name", "control_name", "property_name", "value" } } },
                new { name = "export_form_to_text", description = "Export a form to text format", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "import_form_from_text", description = "Import a form from text format", inputSchema = new { type = "object", properties = new { form_data = new { type = "string" } }, required = new string[] { "form_data" } } },
                new { name = "delete_form", description = "Delete a form from the database", inputSchema = new { type = "object", properties = new { form_name = new { type = "string" } }, required = new string[] { "form_name" } } },
                new { name = "export_report_to_text", description = "Export a report to text format", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } },
                new { name = "import_report_from_text", description = "Import a report from text format", inputSchema = new { type = "object", properties = new { report_data = new { type = "string" } }, required = new string[] { "report_data" } } },
                new { name = "delete_report", description = "Delete a report from the database", inputSchema = new { type = "object", properties = new { report_name = new { type = "string" } }, required = new string[] { "report_name" } } }
            }
        };
    }

    static object HandleToolsCall(AccessInteropService accessService, JsonElement arguments)
    {
        var toolName = arguments.GetProperty("name").GetString();
        
        return toolName switch
        {
            "connect_access" => HandleConnectAccess(accessService, arguments.GetProperty("arguments")),
            "disconnect_access" => HandleDisconnectAccess(accessService, arguments.GetProperty("arguments")),
            "is_connected" => HandleIsConnected(accessService, arguments.GetProperty("arguments")),
            "get_tables" => HandleGetTables(accessService, arguments.GetProperty("arguments")),
            "get_queries" => HandleGetQueries(accessService, arguments.GetProperty("arguments")),
            "get_relationships" => HandleGetRelationships(accessService, arguments.GetProperty("arguments")),
            "create_table" => HandleCreateTable(accessService, arguments.GetProperty("arguments")),
            "delete_table" => HandleDeleteTable(accessService, arguments.GetProperty("arguments")),
            "launch_access" => HandleLaunchAccess(accessService, arguments.GetProperty("arguments")),
            "close_access" => HandleCloseAccess(accessService, arguments.GetProperty("arguments")),
            "get_forms" => HandleGetForms(accessService, arguments.GetProperty("arguments")),
            "get_reports" => HandleGetReports(accessService, arguments.GetProperty("arguments")),
            "get_macros" => HandleGetMacros(accessService, arguments.GetProperty("arguments")),
            "get_modules" => HandleGetModules(accessService, arguments.GetProperty("arguments")),
            "open_form" => HandleOpenForm(accessService, arguments.GetProperty("arguments")),
            "close_form" => HandleCloseForm(accessService, arguments.GetProperty("arguments")),
            "get_vba_projects" => HandleGetVBAProjects(accessService, arguments.GetProperty("arguments")),
            "get_vba_code" => HandleGetVBACode(accessService, arguments.GetProperty("arguments")),
            "set_vba_code" => HandleSetVBACode(accessService, arguments.GetProperty("arguments")),
            "add_vba_procedure" => HandleAddVBAProcedure(accessService, arguments.GetProperty("arguments")),
            "compile_vba" => HandleCompileVBA(accessService, arguments.GetProperty("arguments")),
            "get_system_tables" => HandleGetSystemTables(accessService, arguments.GetProperty("arguments")),
            "get_object_metadata" => HandleGetObjectMetadata(accessService, arguments.GetProperty("arguments")),
            "execute_sql" => HandleExecuteSql(accessService, arguments.GetProperty("arguments")),
            "refresh_schema_cache" => HandleRefreshSchemaCache(accessService, arguments.GetProperty("arguments")),
            "health_check" => HandleHealthCheck(accessService, arguments.GetProperty("arguments")),
            "server_info" => HandleServerInfo(accessService, arguments.GetProperty("arguments")),
            "form_exists" => HandleFormExists(accessService, arguments.GetProperty("arguments")),
            "get_form_controls" => HandleGetFormControls(accessService, arguments.GetProperty("arguments")),
            "get_control_properties" => HandleGetControlProperties(accessService, arguments.GetProperty("arguments")),
            "set_control_property" => HandleSetControlProperty(accessService, arguments.GetProperty("arguments")),
            "export_form_to_text" => HandleExportFormToText(accessService, arguments.GetProperty("arguments")),
            "import_form_from_text" => HandleImportFormFromText(accessService, arguments.GetProperty("arguments")),
            "delete_form" => HandleDeleteForm(accessService, arguments.GetProperty("arguments")),
            "export_report_to_text" => HandleExportReportToText(accessService, arguments.GetProperty("arguments")),
            "import_report_from_text" => HandleImportReportFromText(accessService, arguments.GetProperty("arguments")),
            "delete_report" => HandleDeleteReport(accessService, arguments.GetProperty("arguments")),
            _ => throw new InvalidOperationException($"Unknown tool: {toolName}")
        };
    }

    static object HandleConnectAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            // Get database path from arguments
            if (!arguments.TryGetProperty("database_path", out var dbPathElement))
                return new { success = false, error = "database_path parameter is required" };
                
            var databasePath = dbPathElement.GetString();
            if (string.IsNullOrEmpty(databasePath))
                return new { success = false, error = "database_path cannot be empty" };
            
            // Check if database file exists
            if (!File.Exists(databasePath))
                return new { success = false, error = $"Database file not found: {databasePath}" };
                
            accessService.Connect(databasePath);
            
            // Verify connection was successful
            if (!accessService.IsConnected)
                return new { success = false, error = "Failed to establish database connection" };
                
            return new { success = true, message = $"Connected to {databasePath}", connected = true };
        }
        catch (Exception ex)
        {
            var full = ex.ToString();
            FileLogger.Log($"HandleConnectAccess exception: {full}");
            return new { success = false, error = ex.Message, details = full };
        }
    }

    static object HandleDisconnectAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.Disconnect();
            return new { success = true, message = "Disconnected from database" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleIsConnected(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var isConnected = accessService.IsConnected;
            return new { success = true, connected = isConnected };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tables = accessService.GetTables();
            return new { success = true, tables = tables.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetQueries(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var queries = accessService.GetQueries();
            return new { success = true, queries = queries.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetRelationships(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var relationships = accessService.GetRelationships();
            return new { success = true, relationships = relationships.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCreateTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tableName = arguments.GetProperty("table_name").GetString();
            if (string.IsNullOrEmpty(tableName))
                return new { success = false, error = "Table name is required" };
                
            var fieldsArray = arguments.GetProperty("fields");
            var fields = new List<FieldInfo>();

            foreach (var fieldElement in fieldsArray.EnumerateArray())
            {
                fields.Add(new FieldInfo
                {
                    Name = fieldElement.GetProperty("name").GetString() ?? "",
                    Type = fieldElement.GetProperty("type").GetString() ?? "",
                    Size = fieldElement.GetProperty("size").GetInt32(),
                    Required = fieldElement.GetProperty("required").GetBoolean(),
                    AllowZeroLength = fieldElement.GetProperty("allow_zero_length").GetBoolean()
                });
            }

            accessService.CreateTable(tableName, fields);
            return new { success = true, message = $"Created table {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteTable(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var tableName = arguments.GetProperty("table_name").GetString();
            if (string.IsNullOrEmpty(tableName))
                return new { success = false, error = "Table name is required" };
                
            accessService.DeleteTable(tableName);
            return new { success = true, message = $"Deleted table {tableName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleLaunchAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.LaunchAccess();
            return new { success = true, message = "Access launched successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCloseAccess(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.CloseAccess();
            return new { success = true, message = "Access closed successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetForms(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var forms = accessService.GetForms();
            return new { success = true, forms = forms.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetReports(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reports = accessService.GetReports();
            return new { success = true, reports = reports.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetMacros(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var macros = accessService.GetMacros();
            return new { success = true, macros = macros.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetModules(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var modules = accessService.GetModules();
            return new { success = true, modules = modules.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleOpenForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.OpenForm(formName);
            return new { success = true, message = $"Opened form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCloseForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.CloseForm(formName);
            return new { success = true, message = $"Closed form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetVBAProjects(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projects = accessService.GetVBAProjects();
            return new { success = true, projects = projects.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetVBACode(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                return new { success = false, error = "Project name and module name are required" };
                
            var code = accessService.GetVBACode(projectName, moduleName);
            return new { success = true, code = code };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleSetVBACode(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            var code = arguments.GetProperty("code").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(code))
                return new { success = false, error = "Project name, module name, and code are required" };
                
            accessService.SetVBACode(projectName, moduleName, code);
            return new { success = true, message = $"Updated VBA code in {moduleName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleAddVBAProcedure(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var projectName = arguments.GetProperty("project_name").GetString();
            var moduleName = arguments.GetProperty("module_name").GetString();
            var procedureName = arguments.GetProperty("procedure_name").GetString();
            var code = arguments.GetProperty("code").GetString();
            
            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName) || 
                string.IsNullOrEmpty(procedureName) || string.IsNullOrEmpty(code))
                return new { success = false, error = "All parameters are required" };
                
            accessService.AddVBAProcedure(projectName, moduleName, procedureName, code);
            return new { success = true, message = $"Added VBA procedure {procedureName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleCompileVBA(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.CompileVBA();
            return new { success = true, message = "VBA compiled successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetSystemTables(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var systemTables = accessService.GetSystemTables();
            return new { success = true, system_tables = systemTables.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetObjectMetadata(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var metadata = accessService.GetObjectMetadata();
            return new { success = true, metadata = metadata };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExecuteSql(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            if (!arguments.TryGetProperty("sql", out var sqlElement) || sqlElement.ValueKind != JsonValueKind.String)
                return new { success = false, error = "sql parameter is required" };

            var sql = sqlElement.GetString();
            if (string.IsNullOrWhiteSpace(sql))
                return new { success = false, error = "sql parameter cannot be empty" };

            var mode = "select";
            if (arguments.TryGetProperty("mode", out var modeElement) && modeElement.ValueKind == JsonValueKind.String)
            {
                mode = modeElement.GetString()?.ToLowerInvariant() ?? "select";
            }

            List<object?>? parameters = null;
            if (arguments.TryGetProperty("parameters", out var paramsElement) && paramsElement.ValueKind == JsonValueKind.Array)
            {
                parameters = new List<object?>();
                foreach (var item in paramsElement.EnumerateArray())
                {
                    parameters.Add(JsonElementToObject(item));
                }
            }

            var result = accessService.ExecuteSql(sql, parameters, mode);
            return new { success = true, result = result };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleRefreshSchemaCache(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            accessService.RefreshSchemaCache();
            return new { success = true, message = "Schema cache refreshed" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleHealthCheck(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            return new
            {
                success = true,
                connected = accessService.IsConnected,
                uptime = DateTime.UtcNow,
                server = "Access MCP Server"
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleServerInfo(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            return new
            {
                success = true,
                name = "Access MCP Server",
                version = "1.0.0",
                connected = accessService.IsConnected
            };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleFormExists(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var exists = accessService.FormExists(formName);
            return new { success = true, exists = exists };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetFormControls(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var controls = accessService.GetFormControls(formName);
            return new { success = true, controls = controls.ToArray() };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleGetControlProperties(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();
            
            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName))
                return new { success = false, error = "Form name and control name are required" };
                
            var properties = accessService.GetControlProperties(formName, controlName);
            return new { success = true, properties = properties };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleSetControlProperty(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            var controlName = arguments.GetProperty("control_name").GetString();
            var propertyName = arguments.GetProperty("property_name").GetString();
            var value = arguments.GetProperty("value").GetString();
            
            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName) || 
                string.IsNullOrEmpty(propertyName) || string.IsNullOrEmpty(value))
                return new { success = false, error = "All parameters are required" };
                
            accessService.SetControlProperty(formName, controlName, propertyName, value);
            return new { success = true, message = $"Updated property {propertyName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExportFormToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            var formData = accessService.ExportFormToText(formName);
            return new { success = true, form_data = formData };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleImportFormFromText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formData = arguments.GetProperty("form_data").GetString();
            if (string.IsNullOrEmpty(formData))
                return new { success = false, error = "Form data is required" };
                
            accessService.ImportFormFromText(formData);
            return new { success = true, message = "Form imported successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteForm(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var formName = arguments.GetProperty("form_name").GetString();
            if (string.IsNullOrEmpty(formName))
                return new { success = false, error = "Form name is required" };
                
            accessService.DeleteForm(formName);
            return new { success = true, message = $"Deleted form {formName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleExportReportToText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };
                
            var reportData = accessService.ExportReportToText(reportName);
            return new { success = true, report_data = reportData };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleImportReportFromText(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportData = arguments.GetProperty("report_data").GetString();
            if (string.IsNullOrEmpty(reportData))
                return new { success = false, error = "Report data is required" };
                
            accessService.ImportReportFromText(reportData);
            return new { success = true, message = "Report imported successfully" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }

    static object HandleDeleteReport(AccessInteropService accessService, JsonElement arguments)
    {
        try
        {
            var reportName = arguments.GetProperty("report_name").GetString();
            if (string.IsNullOrEmpty(reportName))
                return new { success = false, error = "Report name is required" };
                
            accessService.DeleteReport(reportName);
            return new { success = true, message = $"Deleted report {reportName}" };
        }
        catch (Exception ex)
        {
            return new { success = false, error = ex.Message };
        }
    }
}

public class JsonRpcRequest
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("method")]
    public string Method { get; set; } = string.Empty;

    [JsonPropertyName("params")]
    public JsonElement Params { get; set; }
}

public class JsonRpcResponse
{
    [JsonPropertyName("jsonrpc")]
    public string Jsonrpc { get; set; } = "2.0";

    [JsonPropertyName("id")]
    public object? Id { get; set; }

    [JsonPropertyName("result")]
    public object? Result { get; set; }

    [JsonPropertyName("error")]
    public JsonRpcError? Error { get; set; }
}

public class JsonRpcError
{
    [JsonPropertyName("code")]
    public int Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("data")]
    public object? Data { get; set; }
}
