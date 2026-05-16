using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Reflection;
using System.Collections.Concurrent;
using System.IO;
using System.Text;

namespace MS.Access.MCP.Interop
{
    public class AccessInteropService : IDisposable
    {
        private OleDbConnection? _oleDbConnection;
        private string? _currentDatabasePath;
        private bool _disposed = false;

        // COM Automation fields
        private object? _accessApplication;  // Late-bound Access.Application COM object
        private object? _currentDatabase;     // Late-bound Access.Database COM object

        // Schema cache
        private List<TableInfo>? _cachedTables;
        private List<QueryInfo>? _cachedQueries;
        private List<RelationshipInfo>? _cachedRelationships;
        private List<SystemTableInfo>? _cachedSystemTables;
        private readonly object _schemaCacheLock = new();

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException($"Database file not found: {databasePath}");

            _currentDatabasePath = databasePath;
            
            try
            {
                // Use OleDb for connection to avoid COM interop dependency during connect
                FileLogger.Log("AccessInteropService.Connect: opening OleDb connection");
                var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};";
                _oleDbConnection = new OleDbConnection(connectionString);
                _oleDbConnection.Open();
                FileLogger.Log($"AccessInteropService.Connect: connected to '{databasePath}'.");

                // Do not require Access.Application during initial connection
                _accessApplication = null;
                _currentDatabase = null;
                InvalidateSchemaCache();
            }
            catch (Exception ex)
            {
                // Clean up on failure
                if (_oleDbConnection != null)
                {
                    try {
                        _oleDbConnection.Close();
                        _oleDbConnection.Dispose();
                    }
                    catch (Exception dbEx)
                    {
                        FileLogger.Log($"Error closing OleDb connection on failure: {dbEx.Message}");
                    }
                    _oleDbConnection = null;
                }

                FileLogger.Log($"AccessInteropService.Connect: failed to connect to database: {ex.Message}");
                throw new InvalidOperationException($"Failed to connect to database: {ex.Message}", ex);
            }
        }

        public void Disconnect()
        {
            try
            {
                FileLogger.Log("AccessInteropService.Disconnect: starting cleanup.");
                // Close OleDb connection first
                if (_oleDbConnection != null)
                {
                    try
                    {
                        _oleDbConnection.Close();
                        _oleDbConnection.Dispose();
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"Error closing OleDb connection: {ex.Message}");
                    }
                    _oleDbConnection = null;
                }

                // Release COM objects in reverse order
                // 1. Close the database
                if (_currentDatabase != null)
                {
                    try
                    {
                        var databaseType = _currentDatabase.GetType();
                        databaseType.InvokeMember("Close", BindingFlags.InvokeMethod, null, _currentDatabase, null);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"Error closing database: {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObjectSafe(_currentDatabase);
                        _currentDatabase = null;
                    }
                }

                // 2. Quit Access Application
                if (_accessApplication != null)
                {
                    try
                    {
                        var accessType = _accessApplication.GetType();

                        try
                        {
                            var forms = accessType.InvokeMember("Forms", BindingFlags.GetProperty, null, _accessApplication, null);
                            ReleaseComObjectSafe(forms);
                        }
                        catch { }

                        try
                        {
                            var reports = accessType.InvokeMember("Reports", BindingFlags.GetProperty, null, _accessApplication, null);
                            ReleaseComObjectSafe(reports);
                        }
                        catch { }

                        accessType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, _accessApplication, null);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"Error quitting Access: {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObjectSafe(_accessApplication);
                        _accessApplication = null;
                    }
                }

                _currentDatabasePath = null;
                InvalidateSchemaCache();
                FileLogger.Log("AccessInteropService.Disconnect: cleanup complete.");
            }
            catch (Exception ex)
            {
                FileLogger.Log($"Error during disconnect: {ex.Message}");
            }
        }

        private void InvalidateSchemaCache()
        {
            lock (_schemaCacheLock)
            {
                _cachedTables = null;
                _cachedQueries = null;
                _cachedRelationships = null;
                _cachedSystemTables = null;
            }
        }

        public void RefreshSchemaCache()
        {
            lock (_schemaCacheLock)
            {
                _cachedTables = null;
                _cachedQueries = null;
                _cachedRelationships = null;
                _cachedSystemTables = null;
            }

            _ = GetTables();
            _ = GetQueries();
            _ = GetRelationships();
            _ = GetSystemTables();
        }

        private static void ReleaseComObjectSafe(object? comObject)
        {
            if (comObject == null)
                return;

            try
            {
                while (Marshal.ReleaseComObject(comObject) > 0) { }
            }
            catch { }
        }

        private dynamic? EnsureAccessApplication()
        {
            if (_accessApplication != null)
                return _accessApplication;

            if (string.IsNullOrEmpty(_currentDatabasePath))
                return null;

            try
            {
                LaunchAccess();
                return _accessApplication;
            }
            catch (Exception ex)
            {
                FileLogger.Log($"EnsureAccessApplication failed: {ex.Message}");
                return null;
            }
        }

        public bool IsConnected => _oleDbConnection?.State == System.Data.ConnectionState.Open;

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            lock (_schemaCacheLock)
            {
                if (_cachedTables != null)
                    return _cachedTables;
            }

            var tables = new List<TableInfo>();
            
            // Use OleDb to get table information
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && !tableName.StartsWith("~"))
                {
                    var fields = GetTableFields(tableName);
                    tables.Add(new TableInfo
                    {
                        Name = tableName,
                        Fields = fields,
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            lock (_schemaCacheLock)
            {
                _cachedTables = tables;
            }

            return tables;
        }

        public List<QueryInfo> GetQueries()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            lock (_schemaCacheLock)
            {
                if (_cachedQueries != null)
                    return _cachedQueries;
            }

            var queries = new List<QueryInfo>();
            var schema = _oleDbConnection!.GetSchema("Views");

            foreach (System.Data.DataRow row in schema.Rows)
            {
                var queryName = row["TABLE_NAME"]?.ToString();
                if (string.IsNullOrEmpty(queryName))
                    continue;

                string sql = string.Empty;
                dynamic? currentDb = null;
                dynamic? queryDef = null;
                try
                {
                    var accessApp = EnsureAccessApplication();
                    if (accessApp != null)
                    {
                        currentDb = accessApp.CurrentDb();
                        queryDef = currentDb.QueryDefs[queryName];
                        if (queryDef != null)
                        {
                            sql = queryDef.SQL ?? string.Empty;
                        }
                    }
                }
                catch { }
                finally
                {
                    ReleaseComObjectSafe(queryDef);
                    ReleaseComObjectSafe(currentDb);
                }

                queries.Add(new QueryInfo
                {
                    Name = queryName,
                    SQL = sql,
                    Type = "Query"
                });
            }

            lock (_schemaCacheLock)
            {
                _cachedQueries = queries;
            }

            return queries;
        }

        public object ExecuteSql(string sql, List<object?>? parameters = null, string mode = "select")
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(sql)) throw new ArgumentException("SQL statement is required.", nameof(sql));

            sql = sql.Trim();
            if (sql.EndsWith(";"))
                sql = sql.TrimEnd(';').TrimEnd();

            if (sql.IndexOf(';') >= 0)
                throw new InvalidOperationException("Multiple SQL statements are not allowed.");

            var normalizedMode = mode?.Trim().ToLowerInvariant() ?? "select";
            if (normalizedMode != "select" && normalizedMode != "nonquery" && normalizedMode != "scalar")
                throw new ArgumentException("Invalid SQL execution mode. Allowed values are select, nonquery, scalar.", nameof(mode));

            var placeholderCount = 0;
            foreach (var ch in sql)
            {
                if (ch == '?')
                    placeholderCount++;
            }

            if (parameters != null && placeholderCount != parameters.Count)
                throw new ArgumentException($"SQL parameter count mismatch. Expected {placeholderCount}, got {parameters.Count}.", nameof(parameters));

            using var command = new OleDbCommand(sql, _oleDbConnection)
            {
                CommandType = CommandType.Text
            };

            if (parameters != null)
            {
                foreach (var parameter in parameters)
                {
                    command.Parameters.Add(CreateOleDbParameter(parameter));
                }
            }

            if (normalizedMode == "scalar")
            {
                return command.ExecuteScalar();
            }

            if (normalizedMode == "nonquery")
            {
                return command.ExecuteNonQuery();
            }

            using var reader = command.ExecuteReader();
            var rows = new List<Dictionary<string, object?>>();
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var name = reader.GetName(i);
                    var value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    row[name] = value;
                }
                rows.Add(row);
            }
            return rows;
        }

        public List<Dictionary<string, object?>> ReadTableData(string objectName, int limit = 50, int offset = 0)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (string.IsNullOrWhiteSpace(objectName)) throw new ArgumentException("Object name is required.", nameof(objectName));
            if (!IsValidObjectName(objectName)) throw new ArgumentException("Invalid object name.", nameof(objectName));
            if (limit <= 0) limit = 50;
            if (offset < 0) offset = 0;

            var sql = $"SELECT * FROM [{objectName}]";
            using var command = new OleDbCommand(sql, _oleDbConnection)
            {
                CommandType = CommandType.Text
            };

            using var reader = command.ExecuteReader();
            var rows = new List<Dictionary<string, object?>>();
            var skipped = 0;
            while (reader.Read())
            {
                if (skipped < offset)
                {
                    skipped++;
                    continue;
                }

                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var name = reader.GetName(i);
                    var rawValue = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    row[name] = NormalizeValue(rawValue);
                }
                rows.Add(row);

                if (rows.Count >= limit)
                    break;
            }

            return rows;
        }

        public object RunMacroOrVBA(string name, List<object?>? arguments = null)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Macro or function name is required.", nameof(name));

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please launch Access first.");

            try
            {
                var invocationArgs = new List<object?> { name };
                if (arguments != null)
                    invocationArgs.AddRange(arguments);

                var result = accessApp.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, accessApp, invocationArgs.ToArray());
                return NormalizeValue(result);
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                throw new InvalidOperationException($"Failed to run macro or VBA function '{name}': {comEx.Message}", comEx);
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                throw new InvalidOperationException($"Failed to run macro or VBA function '{name}': {tie.InnerException.Message}", tie.InnerException);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to run macro or VBA function '{name}': {ex.Message}", ex);
            }
        }

        public string GetFullSchemaMarkdown()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var tables = GetTables();
            var relationships = GetRelationships();
            var builder = new StringBuilder();

            builder.AppendLine("# Database schema");
            builder.AppendLine();
            builder.AppendLine("## Tables");
            builder.AppendLine();

            foreach (var table in tables)
            {
                builder.AppendLine($"### {EscapeMarkdown(table.Name)} ({table.RecordCount} rows)");
                builder.AppendLine();
                builder.AppendLine("| Field | Type | Size | Required | AllowZeroLength |");
                builder.AppendLine("|---|---|---|---|---|");
                foreach (var field in table.Fields)
                {
                    builder.AppendLine($"| {EscapeMarkdown(field.Name)} | {EscapeMarkdown(field.Type)} | {field.Size} | {field.Required} | {field.AllowZeroLength} |");
                }
                builder.AppendLine();
            }

            builder.AppendLine("## Relationships");
            builder.AppendLine();
            builder.AppendLine("| Name | Table | Foreign Table | Attributes |");
            builder.AppendLine("|---|---|---|---|");
            foreach (var relationship in relationships)
            {
                builder.AppendLine($"| {EscapeMarkdown(relationship.Name)} | {EscapeMarkdown(relationship.Table)} | {EscapeMarkdown(relationship.ForeignTable)} | {EscapeMarkdown(relationship.Attributes)} |");
            }

            return builder.ToString().Trim();
        }

        public string GenerateEfCoreModels()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var tables = GetTables();
            var builder = new StringBuilder();
            builder.AppendLine("using System;");
            builder.AppendLine();
            builder.AppendLine("namespace AccessEfCoreModels");
            builder.AppendLine("{");

            foreach (var table in tables)
            {
                var className = SanitizeIdentifier(table.Name);
                builder.AppendLine($"    public class {className}");
                builder.AppendLine("    {");

                if (table.Fields.Count == 0)
                {
                    builder.AppendLine("        // No fields available for this table.");
                }
                else
                {
                    foreach (var field in table.Fields)
                    {
                        var propertyName = SanitizeIdentifier(field.Name);
                        var propertyType = MapAccessTypeToCSharpType(field.Type);
                        builder.AppendLine($"        public {propertyType} {propertyName} {{ get; set; }}");
                    }
                }

                builder.AppendLine("    }");
                builder.AppendLine();
            }

            builder.AppendLine("}");
            return builder.ToString().TrimEnd();
        }

        private static object? NormalizeValue(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return value switch
            {
                byte[] bytes => Convert.ToBase64String(bytes),
                DateTime dateTime => dateTime.ToString("o"),
                Guid guid => guid.ToString(),
                char character => character.ToString(),
                string _ => value,
                bool _ => value,
                byte _ => value,
                sbyte _ => value,
                short _ => value,
                ushort _ => value,
                int _ => value,
                uint _ => value,
                long _ => value,
                ulong _ => value,
                float _ => value,
                double _ => value,
                decimal _ => value,
                _ => value.ToString()
            };
        }

        private static bool IsValidObjectName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            if (name.IndexOfAny(new[] { '[', ']', ';', '\'', '"' }) >= 0)
                return false;

            foreach (var character in name)
            {
                if (!(char.IsLetterOrDigit(character) || character == '_' || character == ' ' || character == '$' || character == '#' || character == '.'))
                    return false;
            }

            return true;
        }

        private static string EscapeMarkdown(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value.Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }

        private static string SanitizeIdentifier(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return "_";

            var builder = new StringBuilder();
            foreach (var ch in input)
            {
                if (char.IsLetterOrDigit(ch) || ch == '_')
                    builder.Append(ch);
                else
                    builder.Append('_');
            }

            var result = builder.ToString();
            if (result.Length == 0)
                return "_";

            if (!char.IsLetter(result[0]) && result[0] != '_')
                result = "_" + result;

            return result;
        }

        private static string SanitizeFileName(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return Guid.NewGuid().ToString("N");

            var builder = new StringBuilder();
            foreach (var ch in input)
            {
                if (Path.GetInvalidFileNameChars().Contains(ch) || char.IsWhiteSpace(ch))
                    builder.Append('_');
                else
                    builder.Append(ch);
            }

            var name = builder.ToString();
            return string.IsNullOrEmpty(name) ? Guid.NewGuid().ToString("N") : name;
        }

        private static string MapAccessTypeToCSharpType(string accessType)
        {
            if (string.IsNullOrWhiteSpace(accessType))
                return "string";

            if (int.TryParse(accessType, out var typeCode))
            {
                return typeCode switch
                {
                    2 => "int",
                    3 => "int",
                    4 => "float",
                    5 => "double",
                    6 => "decimal",
                    7 => "DateTime",
                    10 => "DateTime",
                    11 => "bool",
                    17 => "byte",
                    72 => "Guid",
                    128 => "byte[]",
                    130 => "string",
                    201 => "string",
                    203 => "string",
                    204 => "byte[]",
                    205 => "byte[]",
                    _ => "string",
                };
            }

            var normalized = accessType.Trim().ToLowerInvariant();
            if (normalized.Contains("char") || normalized.Contains("text") || normalized.Contains("memo") || normalized.Contains("string") || normalized.Contains("varchar"))
                return "string";
            if (normalized.Contains("date") || normalized.Contains("time"))
                return "DateTime";
            if (normalized.Contains("bool"))
                return "bool";
            if (normalized.Contains("currency") || normalized.Contains("decimal") || normalized.Contains("numeric"))
                return "decimal";
            if (normalized.Contains("double"))
                return "double";
            if (normalized.Contains("single") || normalized.Contains("float"))
                return "float";
            if (normalized.Contains("byte") || normalized.Contains("binary") || normalized.Contains("oleobject"))
                return "byte[]";
            if (normalized.Contains("guid"))
                return "Guid";

            return "string";
        }

        private static OleDbParameter CreateOleDbParameter(object? value)
        {
            var parameter = new OleDbParameter
            {
                Value = value ?? DBNull.Value,
                OleDbType = GetOleDbTypeForValue(value)
            };

            if (parameter.OleDbType == OleDbType.VarChar)
            {
                parameter.Size = 4000;
            }

            return parameter;
        }

        private static OleDbType GetOleDbTypeForValue(object? value)
        {
            return value switch
            {
                null => OleDbType.VarChar,
                string => OleDbType.VarChar,
                int => OleDbType.Integer,
                long => OleDbType.BigInt,
                bool => OleDbType.Boolean,
                DateTime => OleDbType.Date,
                decimal => OleDbType.Decimal,
                double => OleDbType.Double,
                float => OleDbType.Single,
                byte[] => OleDbType.Binary,
                Guid => OleDbType.Guid,
                _ => OleDbType.VarChar,
            };
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            lock (_schemaCacheLock)
            {
                if (_cachedRelationships != null)
                    return _cachedRelationships;
            }

            var relationships = new List<RelationshipInfo>();
            
            // Use OleDb to get relationship information
            var schema = _oleDbConnection!.GetSchema("ForeignKeys");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                relationships.Add(new RelationshipInfo
                {
                    Name = row["FK_NAME"] != DBNull.Value ? row["FK_NAME"].ToString()! : "",
                    Table = row["TABLE_NAME"] != DBNull.Value ? row["TABLE_NAME"].ToString()! : "",
                    ForeignTable = row["REFERENCED_TABLE_NAME"] != DBNull.Value ? row["REFERENCED_TABLE_NAME"].ToString()! : "",
                    Attributes = ""
                });
            }

            lock (_schemaCacheLock)
            {
                _cachedRelationships = relationships;
            }

            return relationships;
        }

        public void CreateTable(string tableName, List<FieldInfo> fields)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var fieldDefinitions = new List<string>();
            foreach (var field in fields)
            {
                var fieldDef = $"[{field.Name}] {field.Type}";
                if (field.Size > 0 && field.Type.ToLower() == "text")
                    fieldDef += $"({field.Size})";
                if (field.Required)
                    fieldDef += " NOT NULL";
                fieldDefinitions.Add(fieldDef);
            }

            var createSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", fieldDefinitions)})";
            var command = new OleDbCommand(createSql, _oleDbConnection);
            command.ExecuteNonQuery();
        }

        public void DeleteTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var command = new OleDbCommand($"DROP TABLE [{tableName}]", _oleDbConnection);
            command.ExecuteNonQuery();
        }

        #endregion

        #region 3. COM Automation (Simplified)

        public void LaunchAccess()
        {
            if (_accessApplication != null)
            {
                FileLogger.Log("AccessInteropService.LaunchAccess: Access is already launched.");
                return;
            }

            var accessType = Type.GetTypeFromProgID("Access.Application");
            if (accessType == null)
                throw new InvalidOperationException("Microsoft Access is not installed on this system.");

            FileLogger.Log("AccessInteropService.LaunchAccess: creating Access application via ProgID");
            _accessApplication = Activator.CreateInstance(accessType);
            if (_accessApplication == null)
                throw new InvalidOperationException("Failed to instantiate Access.Application via ProgID.");

            try
            {
                accessType.InvokeMember("Visible", BindingFlags.SetProperty, null, _accessApplication, new object[] { false });
            }
            catch { }

            try
            {
                accessType.InvokeMember("AutomationSecurity", BindingFlags.SetProperty, null, _accessApplication, new object[] { 3 });
            }
            catch { }

            try
            {
                accessType.InvokeMember("UserControl", BindingFlags.SetProperty, null, _accessApplication, new object[] { false });
            }
            catch { }

            try
            {
                accessType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _accessApplication, new object[] { false });
            }
            catch { }

            if (!string.IsNullOrEmpty(_currentDatabasePath))
            {
                try
                {
                    FileLogger.Log($"AccessInteropService.LaunchAccess: opening current database '{_currentDatabasePath}' in Access application");
                    accessType.InvokeMember("OpenCurrentDatabase", BindingFlags.InvokeMethod, null, _accessApplication, new object[] { _currentDatabasePath });
                    FileLogger.Log("AccessInteropService.LaunchAccess: database opened successfully in Access application.");
                }
                catch (Exception ex)
                {
                    FileLogger.Log($"AccessInteropService.LaunchAccess: failed to open database in Access app ({ex.Message})");
                }
            }
        }

        public void CloseAccess()
        {
            if (_accessApplication == null)
            {
                FileLogger.Log("AccessInteropService.CloseAccess: Access is not launched.");
                return;
            }

            try
            {
                var accessType = _accessApplication.GetType();
                accessType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, _accessApplication, null);
            }
            catch (Exception ex)
            {
                FileLogger.Log($"Error quitting Access: {ex.Message}");
            }
            finally
            {
                ReleaseComObjectSafe(_accessApplication);
                _accessApplication = null;
                _currentDatabase = null;
            }
        }

        public List<FormInfo> GetForms()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var forms = new List<FormInfo>();
            
            // Try to get forms from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32768", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    forms.Add(new FormInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Form"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return forms;
        }

        public List<ReportInfo> GetReports()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reports = new List<ReportInfo>();
            
            // Try to get reports from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32764", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    reports.Add(new ReportInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Report"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return reports;
        }

        public List<MacroInfo> GetMacros()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var macros = new List<MacroInfo>();
            
            // Try to get macros from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32766", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    macros.Add(new MacroInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Macro"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return macros;
        }

        public List<ModuleInfo> GetModules()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var modules = new List<ModuleInfo>();
            
            // Try to get modules from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    modules.Add(new ModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Module"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return modules;
        }

        public void OpenForm(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            try
            {
                // acForm = 2, acNormal = 0
                var doCmd = _accessApplication.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, _accessApplication, null);
                doCmd.GetType().InvokeMember("OpenForm", BindingFlags.InvokeMethod, null, doCmd, new object[] { formName, 2, null, null, 0 });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to open form '{formName}': {ex.Message}", ex);
            }
        }

        public void CloseForm(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            try
            {
                // acForm = 2, acSaveYes = 1
                var doCmd = _accessApplication.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, _accessApplication, null);
                doCmd.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doCmd, new object[] { 2, formName, 1 });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to close form '{formName}': {ex.Message}", ex);
            }
        }

        public void CreateForm(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            if (FormExists(formName))
                throw new InvalidOperationException($"Form '{formName}' already exists.");

            object? form = null;

            try
            {
                form = _accessApplication.GetType().InvokeMember("CreateForm", BindingFlags.InvokeMethod, null, _accessApplication, null);
                if (form == null)
                    throw new InvalidOperationException("Failed to create form.");

                form.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, form, new object[] { formName });
                form.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, form, new object[] { false });

                var doCmd = _accessApplication.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, _accessApplication, null);
                doCmd.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doCmd, new object[] { 2, formName, 1 });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create form '{formName}': {ex.Message}", ex);
            }
            finally
            {
                if (form != null)
                {
                    try { Marshal.ReleaseComObject(form); } catch { }
                }
            }
        }

        #endregion

        #region 4. VBA Extensibility (Simplified)

        public List<VBAProjectInfo> GetVBAProjects()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var projects = new List<VBAProjectInfo>();
            
            // Simplified VBA project discovery
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                var modules = new List<VBAModuleInfo>();
                while (reader.Read())
                {
                    modules.Add(new VBAModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = "Module",
                        HasCode = true
                    });
                }

                projects.Add(new VBAProjectInfo
                {
                    Name = "CurrentProject",
                    Description = "Current Access Project",
                    Modules = modules
                });
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return projects;
        }

        public string GetVBACode(string projectName, string moduleName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please call launch_access or connect first.");

            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                throw new ArgumentException("Project name and module name are required.");

            try
            {
                var currentProject = accessApp.GetType().InvokeMember("CurrentProject", BindingFlags.GetProperty, null, accessApp, null);
                var vbeProject = currentProject.GetType().InvokeMember("VBProject", BindingFlags.GetProperty, null, currentProject, null);
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                var vbComponents = vbeProject.GetType().InvokeMember("VBComponents", BindingFlags.GetProperty, null, vbeProject, null);
                var component = vbComponents.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, vbComponents, new object[] { moduleName });
                if (component == null)
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");

                var codeModule = component.GetType().InvokeMember("CodeModule", BindingFlags.GetProperty, null, component, null);
                var lineCount = Convert.ToInt32(codeModule.GetType().InvokeMember("CountOfLines", BindingFlags.GetProperty, null, codeModule, null));
                return lineCount > 0
                    ? (string)codeModule.GetType().InvokeMember("Lines", BindingFlags.InvokeMethod, null, codeModule, new object[] { 1, lineCount })
                    : string.Empty;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to get VBA code for module '{moduleName}': {ex.Message}", ex);
            }
        }

        public List<VBAModuleDetail> ExtractRawVBAModules()
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please call launch_access or connect first.");

            var modules = new List<VBAModuleDetail>();
            try
            {
                var currentProject = accessApp.GetType().InvokeMember("CurrentProject", BindingFlags.GetProperty, null, accessApp, null);
                var vbeProject = currentProject.GetType().InvokeMember("VBProject", BindingFlags.GetProperty, null, currentProject, null);
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                var vbComponents = vbeProject.GetType().InvokeMember("VBComponents", BindingFlags.GetProperty, null, vbeProject, null);
                var count = Convert.ToInt32(vbComponents.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, vbComponents, null));

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var component = vbComponents.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, vbComponents, new object[] { i });
                        var name = Convert.ToString(component.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, component, null)) ?? string.Empty;
                        var typeValue = component.GetType().InvokeMember("Type", BindingFlags.GetProperty, null, component, null);
                        var moduleType = typeValue?.ToString() ?? "Unknown";
                        var codeModule = component.GetType().InvokeMember("CodeModule", BindingFlags.GetProperty, null, component, null);
                        var lineCount = Convert.ToInt32(codeModule.GetType().InvokeMember("CountOfLines", BindingFlags.GetProperty, null, codeModule, null));
                        var code = lineCount > 0
                            ? (string)codeModule.GetType().InvokeMember("Lines", BindingFlags.InvokeMethod, null, codeModule, new object[] { 1, lineCount })
                            : string.Empty;

                        modules.Add(new VBAModuleDetail
                        {
                            ProjectName = "CurrentProject",
                            ModuleName = name,
                            ModuleType = moduleType,
                            Code = code
                        });
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"ExtractRawVBAModules: failed reading component {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract VBA modules: {ex.Message}", ex);
            }

            return modules;
        }

        public List<FormMetadata> ExtractFormMetadata()
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please call launch_access or connect first.");

            var formMetadata = new List<FormMetadata>();
            try
            {
                var currentProject = accessApp.GetType().InvokeMember("CurrentProject", BindingFlags.GetProperty, null, accessApp, null);
                var allForms = currentProject.GetType().InvokeMember("AllForms", BindingFlags.GetProperty, null, currentProject, null);
                var count = Convert.ToInt32(allForms.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, allForms, null));

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var formDef = allForms.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, allForms, new object[] { i });
                        var formName = Convert.ToString(formDef.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, formDef, null)) ?? string.Empty;
                        if (string.IsNullOrEmpty(formName))
                            continue;

                        var metadata = ExtractObjectUiMetadata(accessApp, formName, "Form");
                        formMetadata.Add(metadata);
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"ExtractFormMetadata: failed reading form {i}: {ex.Message}");
                        continue;
                    }
                }

                try
                {
                    var allReports = currentProject.GetType().InvokeMember("AllReports", BindingFlags.GetProperty, null, currentProject, null);
                    var reportCount = Convert.ToInt32(allReports.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, allReports, null));

                    for (int i = 1; i <= reportCount; i++)
                    {
                        try
                        {
                            var reportDef = allReports.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, allReports, new object[] { i });
                            var reportName = Convert.ToString(reportDef.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, reportDef, null)) ?? string.Empty;
                            if (string.IsNullOrEmpty(reportName))
                                continue;

                            var metadata = ExtractObjectUiMetadata(accessApp, reportName, "Report");
                            formMetadata.Add(metadata);
                        }
                        catch (Exception ex)
                        {
                            FileLogger.Log($"ExtractFormMetadata: failed reading report {i}: {ex.Message}");
                            continue;
                        }
                    }
                }
                catch
                {
                    // Reports are optional for metadata extraction
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to extract form metadata: {ex.Message}", ex);
            }

            return formMetadata;
        }

        public List<ComplexTypeInfo> ExtractComplexTypes()
        {
            var types = new List<ComplexTypeInfo>();
            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns");
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    var typeName = row["TYPE_NAME"] != DBNull.Value ? Convert.ToString(row["TYPE_NAME"]) ?? string.Empty : string.Empty;
                    var dataType = row["DATA_TYPE"] != DBNull.Value ? Convert.ToString(row["DATA_TYPE"]) ?? string.Empty : string.Empty;
                    var tableName = row["TABLE_NAME"] != DBNull.Value ? Convert.ToString(row["TABLE_NAME"]) ?? string.Empty : string.Empty;
                    var columnName = row["COLUMN_NAME"] != DBNull.Value ? Convert.ToString(row["COLUMN_NAME"]) ?? string.Empty : string.Empty;

                    if (typeName.IndexOf("attachment", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        typeName.IndexOf("oleobject", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        columnName.IndexOf("attachment", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        types.Add(new ComplexTypeInfo
                        {
                            TableName = tableName,
                            ColumnName = columnName,
                            DataType = string.IsNullOrEmpty(typeName) ? dataType : typeName,
                            Notes = "Detected attachment or binary complex type"
                        });
                    }
                }
            }
            catch
            {
                // Fallback when schema extraction fails
            }

            if (!types.Any())
            {
                types.Add(new ComplexTypeInfo
                {
                    TableName = "global",
                    ColumnName = "*",
                    DataType = "ATTACHMENT/LOOKUP",
                    Notes = "No explicit complex field types were detected, but Access may contain hidden lookup/attachment fields."
                });
            }

            return types;
        }

        public List<HiddenLookupInfo> ExtractHiddenLookupFields()
        {
            var lookups = new List<HiddenLookupInfo>();
            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns");
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    var typeName = row["TYPE_NAME"] != DBNull.Value ? Convert.ToString(row["TYPE_NAME"]) ?? string.Empty : string.Empty;
                    var tableName = row["TABLE_NAME"] != DBNull.Value ? Convert.ToString(row["TABLE_NAME"]) ?? string.Empty : string.Empty;
                    var columnName = row["COLUMN_NAME"] != DBNull.Value ? Convert.ToString(row["COLUMN_NAME"]) ?? string.Empty : string.Empty;

                    if (typeName.IndexOf("lookup", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        columnName.IndexOf("lookup", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        lookups.Add(new HiddenLookupInfo
                        {
                            TableName = tableName,
                            ColumnName = columnName,
                            Notes = "Detected lookup field metadata"
                        });
                    }
                }
            }
            catch
            {
                // Fallback to empty list
            }

            if (!lookups.Any())
            {
                lookups.Add(new HiddenLookupInfo
                {
                    TableName = "global",
                    ColumnName = "*",
                    Notes = "No explicit lookup field columns were detected."
                });
            }

            return lookups;
        }

        private FormMetadata ExtractObjectUiMetadata(object accessApp, string objectName, string objectType)
        {
            var metadata = new FormMetadata
            {
                Name = objectName,
                ObjectType = objectType,
                Controls = new List<ControlMetadata>()
            };

            try
            {
                var doCmd = accessApp.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, accessApp, null);
                if (objectType == "Form")
                {
                    doCmd.GetType().InvokeMember("OpenForm", BindingFlags.InvokeMethod, null, doCmd, new object[] { objectName, 0, null, null, 0, 1 });
                }
                else
                {
                    doCmd.GetType().InvokeMember("OpenReport", BindingFlags.InvokeMethod, null, doCmd, new object[] { objectName, 0, null, null, 1 });
                }

                var containerName = objectType == "Form" ? "Forms" : "Reports";
                var collection = accessApp.GetType().InvokeMember(containerName, BindingFlags.GetProperty, null, accessApp, null);
                var obj = collection.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, collection, new object[] { objectName });
                var controls = obj.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, obj, null);
                var controlCount = Convert.ToInt32(controls.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, controls, null));

                for (int j = 1; j <= controlCount; j++)
                {
                    try
                    {
                        var control = controls.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, controls, new object[] { j });
                        var name = Convert.ToString(control.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, control, null)) ?? string.Empty;
                        var typeValue = control.GetType().InvokeMember("ControlType", BindingFlags.GetProperty, null, control, null)?.ToString() ?? string.Empty;
                        var left = SafeGetInt32(control, "Left", 0);
                        var top = SafeGetInt32(control, "Top", 0);
                        var width = SafeGetInt32(control, "Width", 0);
                        var height = SafeGetInt32(control, "Height", 0);
                        var visible = SafeGetBool(control, "Visible", true);
                        var enabled = SafeGetBool(control, "Enabled", true);
                        var controlSource = SafeGetString(control, "ControlSource", string.Empty);

                        var events = new List<ControlEventInfo>();
                        foreach (var eventName in new[] { "OnClick", "OnDblClick", "OnChange", "OnCurrent", "OnLoad", "OnOpen" })
                        {
                            var eventValue = SafeGetString(control, eventName, string.Empty);
                            if (!string.IsNullOrEmpty(eventValue))
                            {
                                events.Add(new ControlEventInfo { EventName = eventName, PropertyValue = eventValue });
                            }
                        }

                        metadata.Controls.Add(new ControlMetadata
                        {
                            Name = name,
                            Type = typeValue,
                            Left = left,
                            Top = top,
                            Width = width,
                            Height = height,
                            Visible = visible,
                            Enabled = enabled,
                            ControlSource = controlSource,
                            BoundField = string.IsNullOrEmpty(controlSource) ? null : controlSource,
                            Events = events
                        });
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                FileLogger.Log($"ExtractObjectUiMetadata failed for {objectType} '{objectName}': {ex.Message}");
            }
            finally
            {
                try
                {
                    var doCmd = accessApp.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, accessApp, null);
                    if (objectType == "Form")
                    {
                        doCmd.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doCmd, new object[] { 2, objectName, 1 });
                    }
                    else
                    {
                        doCmd.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doCmd, new object[] { 3, objectName, 1 });
                    }
                }
                catch { }
            }

            return metadata;
        }

        public void SetVBACode(string projectName, string moduleName, string code)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please call launch_access or connect first.");

            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                throw new ArgumentException("Project name and module name are required.");

            try
            {
                var currentProject = accessApp.GetType().InvokeMember("CurrentProject", BindingFlags.GetProperty, null, accessApp, null);
                var vbeProject = currentProject.GetType().InvokeMember("VBProject", BindingFlags.GetProperty, null, currentProject, null);
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                var vbComponents = vbeProject.GetType().InvokeMember("VBComponents", BindingFlags.GetProperty, null, vbeProject, null);
                var component = vbComponents.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, vbComponents, new object[] { moduleName });
                if (component == null)
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");

                var codeModule = component.GetType().InvokeMember("CodeModule", BindingFlags.GetProperty, null, component, null);
                var countOfLines = (int)codeModule.GetType().InvokeMember("CountOfLines", BindingFlags.GetProperty, null, codeModule, null);

                if (countOfLines > 0)
                    codeModule.GetType().InvokeMember("DeleteLines", BindingFlags.InvokeMethod, null, codeModule, new object[] { 1, countOfLines });

                if (!string.IsNullOrEmpty(code))
                    codeModule.GetType().InvokeMember("AddFromString", BindingFlags.InvokeMethod, null, codeModule, new object[] { code });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to set VBA code for module '{moduleName}': {ex.Message}", ex);
            }
        }

        public void AddVBAProcedure(string projectName, string moduleName, string procedureName, string code)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName) || string.IsNullOrEmpty(procedureName))
                throw new ArgumentException("Project name, module name, and procedure name are required.");

            var procedureCode = code;
            if (string.IsNullOrEmpty(procedureCode))
            {
                procedureCode = $"Public Sub {procedureName}()\n" +
                                "    On Error GoTo ErrorHandler\n\n" +
                                "    ' TODO: implement\n\n" +
                                "ExitProcedure:\n" +
                                "    Exit Sub\n\n" +
                                "ErrorHandler:\n" +
                               $"    MsgBox \"Error \" & Err.Number & \" (\" & Err.Description & \") in procedure {procedureName}\"\n" +
                                "    Resume ExitProcedure\n" +
                                "End Sub";
            }

            try
            {
                dynamic accessApp = (dynamic?)_accessApplication;
                if (accessApp == null)
                    throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

                dynamic vbeProject = accessApp.CurrentProject?.VBProject;
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                dynamic component = vbeProject.VBComponents[moduleName];
                if (component == null)
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");

                dynamic codeModule = component.CodeModule;
                var lineCount = (int)codeModule.CountOfLines;
                var insertAt = lineCount + 1;

                codeModule.InsertLines(insertAt, procedureCode);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to add VBA procedure '{procedureName}' in module '{moduleName}': {comEx.Message}", comEx);
            }
        }

        public void CompileVBA()
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            try
            {
                dynamic accessApp = (dynamic?)_accessApplication;
                if (accessApp == null)
                    throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

                dynamic vbeProject = accessApp.CurrentProject?.VBProject;
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                // Attempt to compile by executing the Access menu command for VBA compile
                // acCmdCompile = 602 or 211? use RunCommand constant 356
                accessApp.DoCmd.RunCommand(600); // acCmdCompile may vary; if invalid, catches
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to compile VBA code: {comEx.Message}", comEx);
            }
        }

        #endregion

        #region 5. System Table Metadata Access

        public List<SystemTableInfo> GetSystemTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            lock (_schemaCacheLock)
            {
                if (_cachedSystemTables != null)
                    return _cachedSystemTables;
            }

            var systemTables = new List<SystemTableInfo>();
            try
            {
                using var command = new OleDbCommand(
                    "SELECT Name, DateCreate, DateUpdate FROM MSysObjects WHERE Name LIKE 'MSys%' OR Name LIKE '~%';",
                    _oleDbConnection);
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var name = reader["Name"]?.ToString() ?? string.Empty;
                    if (string.IsNullOrEmpty(name))
                        continue;

                    DateTime created = DateTime.MinValue;
                    DateTime updated = DateTime.MinValue;

                    try { created = reader["DateCreate"] != DBNull.Value ? Convert.ToDateTime(reader["DateCreate"]) : DateTime.MinValue; } catch { }
                    try { updated = reader["DateUpdate"] != DBNull.Value ? Convert.ToDateTime(reader["DateUpdate"]) : DateTime.MinValue; } catch { }

                    systemTables.Add(new SystemTableInfo
                    {
                        Name = name,
                        DateCreated = created,
                        LastUpdated = updated,
                        RecordCount = GetTableRecordCount(name)
                    });
                }
            }
            catch
            {
                // MSysObjects may be restricted; fallback to schema names with default timestamps
                var schema = _oleDbConnection!.GetSchema("Tables");
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    var tableName = row["TABLE_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(tableName) && (tableName.StartsWith("~") || tableName.StartsWith("MSys")))
                    {
                        systemTables.Add(new SystemTableInfo
                        {
                            Name = tableName,
                            DateCreated = DateTime.MinValue,
                            LastUpdated = DateTime.MinValue,
                            RecordCount = GetTableRecordCount(tableName)
                        });
                    }
                }
            }

            lock (_schemaCacheLock)
            {
                _cachedSystemTables = systemTables;
            }

            return systemTables;
        }

        public List<MetadataInfo> GetObjectMetadata()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var metadata = new List<MetadataInfo>();
            
            try
            {
                // Query MSysObjects table for object metadata
                var command = new OleDbCommand("SELECT * FROM MSysObjects", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    metadata.Add(new MetadataInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = reader["Type"]?.ToString() ?? "",
                        Flags = reader["Flags"]?.ToString() ?? "",
                        DateCreated = reader["DateCreate"]?.ToString() ?? "",
                        DateModified = reader["DateUpdate"]?.ToString() ?? ""
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible, return empty list
            }

            return metadata;
        }

        #endregion

        #region 6. Form & Control Discovery & Editing APIs (Simplified)

        public bool FormExists(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            try
            {
                var command = new OleDbCommand("SELECT COUNT(*) FROM MSysObjects WHERE Name = ? AND Type = -32768", _oleDbConnection);
                command.Parameters.AddWithValue("@Name", formName);
                var count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
            }
            catch
            {
                return false;
            }
        }

        public List<ControlInfo> GetFormControls(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not initialized. Please launch Access first.");

            var controlsInfo = new List<ControlInfo>();

            try
            {
                var forms = _accessApplication.GetType().InvokeMember("Forms", BindingFlags.GetProperty, null, _accessApplication, null);
                var form = forms.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, forms, new object[] { formName });
                var controls = form.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, form, null);
                var count = Convert.ToInt32(controls.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, controls, null));

                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        var control = controls.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, controls, new object[] { i });
                        controlsInfo.Add(new ControlInfo
                        {
                            Name = Convert.ToString(control.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, control, null)) ?? "",
                            Type = Convert.ToString(control.GetType().InvokeMember("ControlType", BindingFlags.GetProperty, null, control, null)) ?? "",
                            Left = Convert.ToInt32(control.GetType().InvokeMember("Left", BindingFlags.GetProperty, null, control, null)),
                            Top = Convert.ToInt32(control.GetType().InvokeMember("Top", BindingFlags.GetProperty, null, control, null)),
                            Width = Convert.ToInt32(control.GetType().InvokeMember("Width", BindingFlags.GetProperty, null, control, null)),
                            Height = Convert.ToInt32(control.GetType().InvokeMember("Height", BindingFlags.GetProperty, null, control, null)),
                            Visible = Convert.ToBoolean(control.GetType().InvokeMember("Visible", BindingFlags.GetProperty, null, control, null)),
                            Enabled = Convert.ToBoolean(control.GetType().InvokeMember("Enabled", BindingFlags.GetProperty, null, control, null))
                        });
                    }
                    catch (Exception ex)
                    {
                        FileLogger.Log($"Error reading control properties: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to enumerate controls in form '{formName}': {ex.Message}", ex);
            }

            return controlsInfo;
        }

        public ControlProperties GetControlProperties(string formName, string controlName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not initialized. Please launch Access first.");

            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName))
                throw new ArgumentException("Form name and control name are required.");

            try
            {
                var forms = _accessApplication.GetType().InvokeMember("Forms", BindingFlags.GetProperty, null, _accessApplication, null);
                var form = forms.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, forms, new object[] { formName });
                var control = form.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, form, null).GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, form.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, form, null), new object[] { controlName });

                var properties = new ControlProperties
                {
                    Name = Convert.ToString(control.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, control, null)) ?? "",
                    Type = Convert.ToString(control.GetType().InvokeMember("ControlType", BindingFlags.GetProperty, null, control, null)) ?? "",
                    Left = SafeGetInt32(control, "Left", 0),
                    Top = SafeGetInt32(control, "Top", 0),
                    Width = SafeGetInt32(control, "Width", 100),
                    Height = SafeGetInt32(control, "Height", 20),
                    Visible = SafeGetBool(control, "Visible", true),
                    Enabled = SafeGetBool(control, "Enabled", true),
                    BackColor = SafeGetInt32(control, "BackColor", -1),
                    ForeColor = SafeGetInt32(control, "ForeColor", 0),
                    FontName = SafeGetString(control, "FontName", "Arial"),
                    FontSize = SafeGetInt32(control, "FontSize", 11),
                    FontBold = SafeGetBool(control, "FontBold", false),
                    FontItalic = SafeGetBool(control, "FontItalic", false)
                };

                return properties;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to get properties for control '{controlName}' in form '{formName}': {ex.Message}", 
                    ex);
            }
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, object value)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not initialized.");

            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName) || string.IsNullOrEmpty(propertyName))
                throw new ArgumentException("Form name, control name, and property name are required.");

            try
            {
                var forms = _accessApplication.GetType().InvokeMember("Forms", BindingFlags.GetProperty, null, _accessApplication, null);
                var form = forms.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, forms, new object[] { formName });
                var control = form.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, form, null).GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, form.GetType().InvokeMember("Controls", BindingFlags.GetProperty, null, form, null), new object[] { controlName });

                control.GetType().InvokeMember(propertyName,
                    BindingFlags.SetProperty,
                    null,
                    control,
                    new object[] { value });
            }
            catch (System.Reflection.TargetInvocationException ex)
            {
                throw new InvalidOperationException(
                    $"Failed to set property '{propertyName}' on control '{controlName}': {ex.InnerException?.Message}",
                    ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to set property '{propertyName}' on control '{controlName}': {ex.Message}",
                    ex);
            }
        }

        // Helper methods to safely extract properties with fallback values
        private int SafeGetInt32(object obj, string propertyName, int defaultValue)
        {
            try
            {
                var value = obj.GetType().InvokeMember(propertyName, 
                    System.Reflection.BindingFlags.GetProperty, null, obj, null);
                return value != null ? Convert.ToInt32(value) : defaultValue;
            }
            catch (COMException)
            {
                return defaultValue;
            }
            catch (System.Reflection.TargetInvocationException)
            {
                return defaultValue;
            }
        }

        private bool SafeGetBool(object obj, string propertyName, bool defaultValue)
        {
            try
            {
                var value = obj.GetType().InvokeMember(propertyName, 
                    System.Reflection.BindingFlags.GetProperty, null, obj, null);
                return value != null ? Convert.ToBoolean(value) : defaultValue;
            }
            catch (COMException)
            {
                return defaultValue;
            }
        }

        private string SafeGetString(object obj, string propertyName, string defaultValue)
        {
            try
            {
                var value = obj.GetType().InvokeMember(propertyName, 
                    System.Reflection.BindingFlags.GetProperty, null, obj, null);
                return value?.ToString() ?? defaultValue;
            }
            catch (COMException)
            {
                return defaultValue;
            }
        }

        #endregion

        #region 7. Persistence & Versioning

        public string ExportFormToText(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formData = new
            {
                Name = formName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(formName),
                VBA = GetVBACode("CurrentProject", formName)
            };

            return JsonSerializer.Serialize(formData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportFormFromText(string formData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formInfo = JsonSerializer.Deserialize<FormExportData>(formData);
            if (formInfo == null) throw new ArgumentException("Invalid form data");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            dynamic? form = null;
            try
            {
                dynamic accessApp = _accessApplication;
                form = accessApp.CreateForm();
                form.Name = formInfo.Name;
                form.Visible = false;
                accessApp.DoCmd.Save(2, formInfo.Name);
                accessApp.DoCmd.Close(2, formInfo.Name, 1);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to import form '{formInfo.Name}': {ex.Message}", ex);
            }
            finally
            {
                ReleaseComObjectSafe(form);
            }
        }

        public void DeleteForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            try
            {
                dynamic accessApp = _accessApplication;
                accessApp.DoCmd.DeleteObject(2, formName);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to delete form '{formName}': {ex.Message}", ex);
            }
        }

        public string ExportReportToText(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportData = new
            {
                Name = reportName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(reportName) // Reuse form controls for reports
            };

            return JsonSerializer.Serialize(reportData, new JsonSerializerOptions { WriteIndented = true });
        }

        public string ExportReportToPdf(string reportName)
        {
            if (string.IsNullOrWhiteSpace(reportName))
                throw new ArgumentException("Report name is required.", nameof(reportName));

            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reports = GetReports();
            if (!reports.Exists(r => string.Equals(r.Name, reportName, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException($"Report '{reportName}' does not exist.", nameof(reportName));

            var accessApp = EnsureAccessApplication();
            if (accessApp == null)
                throw new InvalidOperationException("Access application is not available. Please launch Access first.");

            var outputFileName = $"{SanitizeFileName(reportName)}_{Guid.NewGuid():N}.pdf";
            var outputFilePath = Path.Combine(Path.GetTempPath(), outputFileName);

            dynamic? doCmd = null;
            try
            {
                doCmd = accessApp.GetType().InvokeMember("DoCmd", BindingFlags.GetProperty, null, accessApp, null);
                doCmd.GetType().InvokeMember("OutputTo", BindingFlags.InvokeMethod, null, doCmd, new object[] { 3, reportName, "PDF", outputFilePath, false });
                return Path.GetFullPath(outputFilePath);
            }
            catch (TargetInvocationException tie) when (tie.InnerException != null)
            {
                throw new InvalidOperationException($"Failed to export report '{reportName}' to PDF: {tie.InnerException.Message}", tie.InnerException);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to export report '{reportName}' to PDF: {ex.Message}", ex);
            }
            finally
            {
                ReleaseComObjectSafe(doCmd);
            }
        }

        public void ImportReportFromText(string reportData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportInfo = JsonSerializer.Deserialize<ReportExportData>(reportData);
            if (reportInfo == null) throw new ArgumentException("Invalid report data");

            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            dynamic? report = null;
            try
            {
                dynamic accessApp = _accessApplication;
                report = accessApp.CreateReport();
                report.Name = reportInfo.Name;
                report.Visible = false;
                accessApp.DoCmd.Save(3, reportInfo.Name);
                accessApp.DoCmd.Close(3, reportInfo.Name, 1);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to import report '{reportInfo.Name}': {ex.Message}", ex);
            }
            finally
            {
                ReleaseComObjectSafe(report);
            }
        }

        public void DeleteReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            if (_accessApplication == null)
                throw new InvalidOperationException("Access application is not launched. Please call launch_access first.");

            try
            {
                dynamic accessApp = _accessApplication;
                accessApp.DoCmd.DeleteObject(3, reportName);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to delete report '{reportName}': {ex.Message}", ex);
            }
        }

        #endregion

        #region Helper Methods

        private List<FieldInfo> GetTableFields(string tableName)
        {
            var fields = new List<FieldInfo>();
            
            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName });
                
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    fields.Add(new FieldInfo
                    {
                        Name = row["COLUMN_NAME"] != DBNull.Value ? row["COLUMN_NAME"].ToString()! : "",
                        Type = row["DATA_TYPE"] != DBNull.Value ? row["DATA_TYPE"].ToString()! : "",
                        Size = row["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value ? Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"]) : 0,
                        Required = row["IS_NULLABLE"] != DBNull.Value && row["IS_NULLABLE"].ToString() == "NO",
                        AllowZeroLength = true // Default value
                    });
                }
            }
            catch
            {
                // Return empty list if table doesn't exist or can't be accessed
            }

            return fields;
        }

        private long GetTableRecordCount(string tableName)
        {
            try
            {
                var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", _oleDbConnection);
                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                Disconnect();
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }

        ~AccessInteropService()
        {
            Dispose();
        }
    }

    #region Data Models

    public class TableInfo
    {
        public string Name { get; set; } = "";
        public List<FieldInfo> Fields { get; set; } = new();
        public long RecordCount { get; set; }
    }

    public class FieldInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
    }

    public class QueryInfo
    {
        public string Name { get; set; } = "";
        public string SQL { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class RelationshipInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public string ForeignTable { get; set; } = "";
        public string Attributes { get; set; } = "";
    }

    public class FormInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ReportInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class MacroInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ModuleInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class VBAProjectInfo
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public List<VBAModuleInfo> Modules { get; set; } = new();
    }

    public class VBAModuleInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public bool HasCode { get; set; }
    }

    public class SystemTableInfo
    {
        public string Name { get; set; } = "";
        public DateTime DateCreated { get; set; }
        public DateTime LastUpdated { get; set; }
        public long RecordCount { get; set; }
    }

    public class MetadataInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Flags { get; set; } = "";
        public string DateCreated { get; set; } = "";
        public string DateModified { get; set; } = "";
    }

    public class ControlInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
    }

    public class ControlProperties
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public int BackColor { get; set; }
        public int ForeColor { get; set; }
        public string FontName { get; set; } = "";
        public int FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
    }

    public class ControlEventInfo
    {
        public string EventName { get; set; } = "";
        public string PropertyValue { get; set; } = "";
    }

    public class ControlMetadata
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public string ControlSource { get; set; } = "";
        public string? BoundField { get; set; }
        public List<ControlEventInfo> Events { get; set; } = new();
    }

    public class FormMetadata
    {
        public string Name { get; set; } = "";
        public string ObjectType { get; set; } = "Form";
        public List<ControlMetadata> Controls { get; set; } = new();
    }

    public class VBAModuleDetail
    {
        public string ProjectName { get; set; } = "";
        public string ModuleName { get; set; } = "";
        public string ModuleType { get; set; } = "";
        public string Code { get; set; } = "";
    }

    public class ComplexTypeInfo
    {
        public string TableName { get; set; } = "";
        public string ColumnName { get; set; } = "";
        public string DataType { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    public class HiddenLookupInfo
    {
        public string TableName { get; set; } = "";
        public string ColumnName { get; set; } = "";
        public string Notes { get; set; } = "";
    }

    public class FormExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
        public string VBA { get; set; } = "";
    }

    public class ReportExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
    }

    public static class FileLogger
    {
        private static readonly string LogPath = Path.Combine(AppContext.BaseDirectory, "mcp_server.log");
        private static readonly BlockingCollection<string> Queue = new(new ConcurrentQueue<string>());
        private static readonly Task BackgroundWriter;
        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            WriteIndented = false
        };

        static FileLogger()
        {
            try
            {
                var logDirectory = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrEmpty(logDirectory))
                    Directory.CreateDirectory(logDirectory);
            }
            catch { }

            BackgroundWriter = Task.Factory.StartNew(() =>
            {
                foreach (var entry in Queue.GetConsumingEnumerable())
                {
                    try
                    {
                        File.AppendAllText(LogPath, entry + Environment.NewLine, Encoding.UTF8);
                    }
                    catch { }
                }
            }, TaskCreationOptions.LongRunning);
        }

        public static void Log(string message)
        {
            Log(new { level = "Information", @event = "LogMessage", message });
        }

        public static void Log(string level, string eventName, object? data = null)
        {
            Log(new
            {
                timestamp = DateTime.UtcNow,
                level,
                @event = eventName,
                data
            });
        }

        public static void Log(object payload)
        {
            try
            {
                var line = JsonSerializer.Serialize(payload, SerializerOptions);
                Queue.Add(line);
            }
            catch { }
        }

        public static void Shutdown()
        {
            try
            {
                Queue.CompleteAdding();
                BackgroundWriter.Wait(1000);
            }
            catch { }
        }
    }

    #endregion
} 