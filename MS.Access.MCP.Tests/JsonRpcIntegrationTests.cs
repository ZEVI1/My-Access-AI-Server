using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using MS.Access.MCP.Interop;
using Xunit;

namespace MS.Access.MCP.Tests
{
    public class JsonRpcIntegrationTests
    {
        [Fact]
        public async Task ReadRpcMessageAsync_ParsesContentLengthAndReturnsPayload()
        {
            var payload = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"ping\"}";
            var message = $"Content-Length: {Encoding.UTF8.GetByteCount(payload)}\r\n\r\n{payload}";

            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(message));
            var result = await Program.ReadRpcMessageAsync(stream, CancellationToken.None);

            Assert.Equal(payload, result);
        }

        [Fact]
        public void ProcessRpcMessage_Ping_ReturnsSuccessResult()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"ping\"}";
            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error);
            Assert.Equal(1L, response.Id);
            Assert.NotNull(response.Result);
        }

        [Fact]
        public void ProcessRpcMessage_ToolsList_ReturnsToolsArray()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":\"tools\",\"method\":\"tools/list\"}";
            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error);
            Assert.Equal("tools", response.Id);

            var resultJson = JsonSerializer.Serialize(response.Result);
            Assert.Contains("\"tools\":", resultJson);
            Assert.Contains("connect_access", resultJson);
            Assert.Contains("delete_report", resultJson);
            Assert.Contains("export_report_to_pdf", resultJson);
            Assert.Contains("generate_ef_core_models", resultJson);
            Assert.Contains("get_table_data", resultJson);
            Assert.Contains("run_macro_or_vba", resultJson);
            Assert.Contains("get_full_schema_markdown", resultJson);
        }

        [Fact]
        public void AccessInteropService_ReadTableDataWithoutConnection_ThrowsInvalidOperationException()
        {
            using var accessService = new AccessInteropService();
            var exception = Assert.Throws<InvalidOperationException>(() => accessService.ReadTableData("AnyTable"));
            Assert.Contains("Not connected", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AccessInteropService_RunMacroOrVBAWithoutAccess_ThrowsInvalidOperationException()
        {
            using var accessService = new AccessInteropService();
            var exception = Assert.Throws<InvalidOperationException>(() => accessService.RunMacroOrVBA("TestMacro"));
            Assert.Contains("Access application is not available", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AccessInteropService_GenerateEfCoreModelsWithoutConnection_ThrowsInvalidOperationException()
        {
            using var accessService = new AccessInteropService();
            var exception = Assert.Throws<InvalidOperationException>(() => accessService.GenerateEfCoreModels());
            Assert.Contains("Not connected", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AccessInteropService_ExportReportToPdfWithoutConnection_ThrowsInvalidOperationException()
        {
            using var accessService = new AccessInteropService();
            var exception = Assert.Throws<InvalidOperationException>(() => accessService.ExportReportToPdf("AnyReport"));
            Assert.Contains("Not connected", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ProcessRpcMessage_InvalidJson_ReturnsParseError()
        {
            using var accessService = new AccessInteropService();
            var response = Program.ProcessRpcMessage(accessService, "{ invalid json }");

            Assert.NotNull(response);
            Assert.NotNull(response.Error);
            Assert.Equal(-32700, response.Error.Code);
        }

        [Fact]
        public void AccessInteropService_ConnectInvalidPath_ThrowsFileNotFoundException()
        {
            using var accessService = new AccessInteropService();

            var exception = Assert.Throws<FileNotFoundException>(() => accessService.Connect("nonexistent-path.accdb"));

            Assert.Contains("Database file not found", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void AccessInteropService_Connect_NullOrEmptyPath_ThrowsArgumentNullException()
        {
            using var accessService = new AccessInteropService();
            Assert.Throws<ArgumentNullException>(() => accessService.Connect(string.Empty));
            Assert.Throws<ArgumentNullException>(() => accessService.Connect(null!));
        }

        [Fact]
        public void CreateTable_InvalidTableName_ThrowsArgumentException()
        {
            using var accessService = new AccessInteropService();
            // The validation happens before IsConnected, so this should throw ArgumentException
            Assert.Throws<ArgumentException>(() => accessService.ReadTableData("Invalid;Name"));
        }
    }
}
