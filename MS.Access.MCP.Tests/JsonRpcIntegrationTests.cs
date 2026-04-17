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
        public void ProcessRpcMessage_ReadTableDataWithoutConnection_ReturnsErrorInResult()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":\"test\",\"method\":\"tools/call\",\"params\":{\"name\":\"get_table_data\",\"arguments\":{\"object_name\":\"AnyTable\",\"limit\":10,\"offset\":0}}}";

            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error); // Errors from tools are in Result
            Assert.NotNull(response.Result);

            var resultJson = JsonSerializer.Serialize(response.Result);
            Assert.Contains("\"success\":false", resultJson);
            Assert.Contains("Not connected", resultJson, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ProcessRpcMessage_RunMacroOrVBAWithoutAccess_ReturnsErrorInResult()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":\"test\",\"method\":\"tools/call\",\"params\":{\"name\":\"run_macro_or_vba\",\"arguments\":{\"name\":\"TestMacro\"}}}";

            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error);
            Assert.NotNull(response.Result);

            var resultJson = JsonSerializer.Serialize(response.Result);
            Assert.Contains("\"success\":false", resultJson);
            Assert.Contains("Access application is not available", resultJson, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ProcessRpcMessage_GenerateEfCoreModelsWithoutConnection_ReturnsErrorInResult()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":\"test\",\"method\":\"tools/call\",\"params\":{\"name\":\"generate_ef_core_models\",\"arguments\":{}}}";

            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error);
            Assert.NotNull(response.Result);

            var resultJson = JsonSerializer.Serialize(response.Result);
            Assert.Contains("\"success\":false", resultJson);
            Assert.Contains("Not connected", resultJson, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ProcessRpcMessage_ExportReportToPdfWithoutConnection_ReturnsErrorInResult()
        {
            using var accessService = new AccessInteropService();
            var request = "{\"jsonrpc\":\"2.0\",\"id\":\"test\",\"method\":\"tools/call\",\"params\":{\"name\":\"export_report_to_pdf\",\"arguments\":{\"report_name\":\"AnyReport\"}}}";

            var response = Program.ProcessRpcMessage(accessService, request);

            Assert.NotNull(response);
            Assert.Null(response.Error);
            Assert.NotNull(response.Result);

            var resultJson = JsonSerializer.Serialize(response.Result);
            Assert.Contains("\"success\":false", resultJson);
            Assert.Contains("Not connected", resultJson, StringComparison.OrdinalIgnoreCase);
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
    }
}
