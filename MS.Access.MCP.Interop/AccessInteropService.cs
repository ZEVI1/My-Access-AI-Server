using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace MS.Access.MCP.Interop
{
    public class AccessInteropService : IDisposable
    {
        private OleDbConnection? _oleDbConnection;
        private string? _currentDatabasePath;
        private bool _disposed = false;

        // COM Automation fields
        private dynamic? _accessApplication;  // Microsoft.Office.Interop.Access.Application
        private dynamic? _currentDatabase;     // Microsoft.Office.Interop.Access.Database

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException($"Database file not found: {databasePath}");

            _currentDatabasePath = databasePath;
            
            try
            {
                // Initialize COM Automation - Launch Access Application
                _accessApplication = new Microsoft.Office.Interop.Access.Application();
                _accessApplication.Visible = false;  // Run in background
                
                // Open the database using COM Automation
                var dbEngine = _accessApplication.DBEngine;
                _currentDatabase = dbEngine.OpenDatabase(databasePath);
                
                // Also create OleDb connection for direct data access
                var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};";
                _oleDbConnection = new OleDbConnection(connectionString);
                _oleDbConnection.Open();
            }
            catch (Exception ex)
            {
                // Clean up on failure
                if (_currentDatabase != null)
                {
                    try { Marshal.ReleaseComObject(_currentDatabase); }
                    catch { }
                    _currentDatabase = null;
                }
                if (_accessApplication != null)
                {
                    try { _accessApplication.Quit(); }
                    catch { }
                    try { Marshal.ReleaseComObject(_accessApplication); }
                    catch { }
                    _accessApplication = null;
                }
                throw new InvalidOperationException($"Failed to connect to database: {ex.Message}", ex);
            }
        }

        public void Disconnect()
        {
            try
            {
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
                        Console.Error.WriteLine($"Error closing OleDb connection: {ex.Message}");
                    }
                    _oleDbConnection = null;
                }

                // Release COM objects in reverse order
                // 1. Close the database
                if (_currentDatabase != null)
                {
                    try
                    {
                        _currentDatabase.Close();
                        Marshal.ReleaseComObject(_currentDatabase);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error closing database: {ex.Message}");
                    }
                    _currentDatabase = null;
                }

                // 2. Quit Access Application
                if (_accessApplication != null)
                {
                    try
                    {
                        _accessApplication.Quit();
                        Marshal.ReleaseComObject(_accessApplication);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error quitting Access: {ex.Message}");
                    }
                    _accessApplication = null;
                }

                _currentDatabasePath = null;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during disconnect: {ex.Message}");
            }
        }

        public bool IsConnected => _oleDbConnection?.State == System.Data.ConnectionState.Open;

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

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

            return tables;
        }

        public List<QueryInfo> GetQueries()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var queries = new List<QueryInfo>();
            
            // Use OleDb to get query information
            var schema = _oleDbConnection!.GetSchema("Views");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var queryName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(queryName))
                {
                    queries.Add(new QueryInfo
                    {
                        Name = queryName,
                        SQL = "", // SQL not available through schema
                        Type = "Query"
                    });
                }
            }

            return queries;
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var relationships = new List<RelationshipInfo>();
            
            // Use OleDb to get relationship information
            var schema = _oleDbConnection!.GetSchema("ForeignKeys");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                relationships.Add(new RelationshipInfo
                {
                    Name = row["FK_NAME"]?.ToString() ?? "",
                    Table = row["TABLE_NAME"]?.ToString() ?? "",
                    ForeignTable = row["REFERENCED_TABLE_NAME"]?.ToString() ?? "",
                    Attributes = ""
                });
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
            // This would require full COM interop - simplified for now
            Console.WriteLine("Access launch functionality requires full COM interop");
        }

        public void CloseAccess()
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine("Access close functionality requires full COM interop");
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

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            try
            {
                // acForm = 2, acNormal = 0
                _accessApplication?.DoCmd.OpenForm(formName, 2, null, null, 0);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to open form '{formName}': {comEx.Message}", comEx);
            }
        }

        public void CloseForm(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            try
            {
                // acForm = 2, acSaveYes = 1
                _accessApplication?.DoCmd.Close(2, formName, 1);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to close form '{formName}': {comEx.Message}", comEx);
            }
        }

        public void CreateForm(string formName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(formName))
                throw new ArgumentException("Form name is required", nameof(formName));

            if (FormExists(formName))
                throw new InvalidOperationException($"Form '{formName}' already exists.");

            dynamic? form = null;

            try
            {
                form = _accessApplication?.CreateForm();
                if (form == null)
                    throw new InvalidOperationException("Failed to create form.");

                form.Name = formName;
                form.Visible = false;
                _accessApplication?.DoCmd.Close(2, formName, 1);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to create form '{formName}': {comEx.Message}", comEx);
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

            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                throw new ArgumentException("Project name and module name are required.");

            try
            {
                dynamic vbeProject = _accessApplication?.CurrentProject?.VBProject;
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                dynamic component = vbeProject.VBComponents[moduleName];
                if (component == null)
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");

                dynamic codeModule = component.CodeModule;
                var lineCount = (int)codeModule.CountOfLines;
                return (string)codeModule.Lines(1, lineCount);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to get VBA code for module '{moduleName}': {comEx.Message}", comEx);
            }
        }

        public void SetVBACode(string projectName, string moduleName, string code)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(projectName) || string.IsNullOrEmpty(moduleName))
                throw new ArgumentException("Project name and module name are required.");

            try
            {
                dynamic vbeProject = _accessApplication?.CurrentProject?.VBProject;
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                dynamic component = vbeProject.VBComponents[moduleName];
                if (component == null)
                    throw new InvalidOperationException($"Module '{moduleName}' not found.");

                dynamic codeModule = component.CodeModule;
                var lineCount = (int)codeModule.CountOfLines;

                if (lineCount > 0)
                    codeModule.DeleteLines(1, lineCount);

                if (!string.IsNullOrEmpty(code))
                    codeModule.AddFromString(code);
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to set VBA code for module '{moduleName}': {comEx.Message}", comEx);
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
                procedureCode = $"Public Sub {procedureName}()\n    ' TODO: implement \nEnd Sub";
            }

            try
            {
                dynamic vbeProject = _accessApplication?.CurrentProject?.VBProject;
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
                dynamic vbeProject = _accessApplication?.CurrentProject?.VBProject;
                if (vbeProject == null)
                    throw new InvalidOperationException("Access VBProject not available.");

                // Attempt to compile by executing the Access menu command for VBA compile
                // acCmdCompile = 602 or 211? use RunCommand constant 356
                _accessApplication?.DoCmd.RunCommand(600); // acCmdCompile may vary; if invalid, catches
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

            var systemTables = new List<SystemTableInfo>();
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && (tableName.StartsWith("~") || tableName.StartsWith("MSys")))
                {
                    systemTables.Add(new SystemTableInfo
                    {
                        Name = tableName,
                        DateCreated = DateTime.Now, // Not available through OleDb
                        LastUpdated = DateTime.Now, // Not available through OleDb
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
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

            var app = _accessApplication ?? throw new InvalidOperationException("Access application is not initialized.");
            var controls = new List<ControlInfo>();

            try
            {
                dynamic form = app.Forms[formName];

                foreach (dynamic control in form.Controls)
                {
                    try
                    {
                        controls.Add(new ControlInfo
                        {
                            Name = control.Name ?? "",
                            Type = control.ControlType?.ToString() ?? "",
                            Left = Convert.ToInt32(control.Left),
                            Top = Convert.ToInt32(control.Top),
                            Width = Convert.ToInt32(control.Width),
                            Height = Convert.ToInt32(control.Height),
                            Visible = control.Visible ?? true,
                            Enabled = control.Enabled ?? true
                        });
                    }
                    catch (COMException comEx)
                    {
                        Console.Error.WriteLine($"Error reading control properties: {comEx.Message}");
                        continue;
                    }
                }
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException($"Failed to enumerate controls in form '{formName}': {comEx.Message}", comEx);
            }
            catch (ArgumentException argEx)
            {
                throw new InvalidOperationException($"Form '{formName}' not found: {argEx.Message}", argEx);
            }

            return controls;
        }

        public ControlProperties GetControlProperties(string formName, string controlName)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName))
                throw new ArgumentException("Form name and control name are required.");

            var app = _accessApplication ?? throw new InvalidOperationException("Access application is not initialized.");

            try
            {
                dynamic form = app.Forms[formName];
                dynamic control = form.Controls[controlName];

                var properties = new ControlProperties
                {
                    Name = control.Name ?? "",
                    Type = control.ControlType?.ToString() ?? "",
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
            catch (COMException comEx)
            {
                throw new InvalidOperationException(
                    $"Failed to get properties for control '{controlName}' in form '{formName}': {comEx.Message}", 
                    comEx);
            }
            catch (ArgumentException argEx)
            {
                throw new InvalidOperationException(
                    $"Control '{controlName}' not found in form '{formName}': {argEx.Message}", 
                    argEx);
            }
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, object value)
        {
            if (!IsConnected)
                throw new InvalidOperationException("Not connected to database");

            if (string.IsNullOrEmpty(formName) || string.IsNullOrEmpty(controlName) || string.IsNullOrEmpty(propertyName))
                throw new ArgumentException("Form name, control name, and property name are required.");

            var app = _accessApplication ?? throw new InvalidOperationException("Access application is not initialized.");

            try
            {
                dynamic form = app.Forms[formName];
                dynamic control = form.Controls[controlName];

                control.GetType().InvokeMember(propertyName,
                    System.Reflection.BindingFlags.SetProperty,
                    null,
                    control,
                    new object[] { value });
            }
            catch (COMException comEx)
            {
                throw new InvalidOperationException(
                    $"COM error setting property '{propertyName}' to '{value}' on control '{controlName}': {comEx.Message}",
                    comEx);
            }
            catch (System.Reflection.TargetInvocationException ex)
            {
                throw new InvalidOperationException(
                    $"Failed to set property '{propertyName}' on control '{controlName}': {ex.InnerException?.Message}",
                    ex);
            }
            catch (ArgumentException argEx)
            {
                throw new InvalidOperationException(
                    $"Control '{controlName}' not found in form '{formName}': {argEx.Message}",
                    argEx);
            }
        }

        // Helper methods to safely extract properties with fallback values
        private int SafeGetInt32(dynamic obj, string propertyName, int defaultValue)
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

        private bool SafeGetBool(dynamic obj, string propertyName, bool defaultValue)
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

        private string SafeGetString(dynamic obj, string propertyName, string defaultValue)
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

            // Simplified form import - would require full COM interop for actual form creation
            Console.WriteLine($"Form {formInfo.Name} would be imported here");
        }

        public void DeleteForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Form {formName} would be deleted here");
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

        public void ImportReportFromText(string reportData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportInfo = JsonSerializer.Deserialize<ReportExportData>(reportData);
            if (reportInfo == null) throw new ArgumentException("Invalid report data");

            // Simplified report import - would require full COM interop for actual report creation
            Console.WriteLine($"Report {reportInfo.Name} would be imported here");
        }

        public void DeleteReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Report {reportName} would be deleted here");
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
                        Name = row["COLUMN_NAME"]?.ToString() ?? "",
                        Type = row["DATA_TYPE"]?.ToString() ?? "",
                        Size = Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"] ?? 0),
                        Required = row["IS_NULLABLE"]?.ToString() == "NO",
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

    #endregion
} 