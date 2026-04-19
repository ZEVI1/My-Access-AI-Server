using System;
using System.Collections.Generic;
using System.Reflection;
using MS.Access.MCP.Interop;
using Xunit;

namespace MS.Access.MCP.Tests
{
    public class ExecuteSqlTests
    {
        // Helper to test validation logic which is done before DB interaction
        private class TestAccessInteropService : AccessInteropService
        {
            public void SetConnected(bool connected)
            {
                // We mock the connection state to test validation logic
                var field = typeof(AccessInteropService).GetField("_oleDbConnection",
                    BindingFlags.NonPublic | BindingFlags.Instance);

                if (connected)
                {
                    // This is a hack - setting to an object just so it's not null
                    // Be aware this will fail when it actually tries to use it
                    field!.SetValue(this, new System.Data.OleDb.OleDbConnection());
                }
            }
        }

        [Fact]
        public void ExecuteSql_EmptySql_ThrowsArgumentException()
        {
            using var service = new TestAccessInteropService();
            service.SetConnected(true);

            var exception = Assert.Throws<ArgumentException>(() =>
                service.ExecuteSql("  "));

            Assert.Contains("SQL statement is required", exception.Message);
        }

        [Fact]
        public void ExecuteSql_MultipleStatements_ThrowsInvalidOperationException()
        {
            using var service = new TestAccessInteropService();
            service.SetConnected(true);

            var exception = Assert.Throws<InvalidOperationException>(() =>
                service.ExecuteSql("SELECT * FROM A; SELECT * FROM B"));

            Assert.Contains("Multiple SQL statements are not allowed", exception.Message);
        }

        [Fact]
        public void ExecuteSql_InvalidMode_ThrowsArgumentException()
        {
            using var service = new TestAccessInteropService();
            service.SetConnected(true);

            var exception = Assert.Throws<ArgumentException>(() =>
                service.ExecuteSql("SELECT * FROM A", mode: "invalid_mode"));

            Assert.Contains("Invalid SQL execution mode", exception.Message);
        }

        [Fact]
        public void ExecuteSql_ParameterCountMismatch_ThrowsArgumentException()
        {
            using var service = new TestAccessInteropService();
            service.SetConnected(true);

            var exception = Assert.Throws<ArgumentException>(() =>
                service.ExecuteSql("SELECT * FROM A WHERE Id = ? AND Name = ?",
                new List<object?> { 1 })); // Only 1 parameter provided but 2 placeholders

            Assert.Contains("SQL parameter count mismatch", exception.Message);
        }

        [Fact]
        public void ExecuteSql_NotConnected_ThrowsInvalidOperationException()
        {
            using var service = new TestAccessInteropService();
            // Not setting connected to true

            var exception = Assert.Throws<InvalidOperationException>(() =>
                service.ExecuteSql("SELECT * FROM A WHERE Id = ?", new List<object?> { 1 }));

            Assert.Contains("Not connected to database", exception.Message);
        }
    }
}
