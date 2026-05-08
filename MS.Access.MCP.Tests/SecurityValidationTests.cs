using System;
using System.Collections.Generic;
using MS.Access.MCP.Interop;
using Xunit;

namespace MS.Access.MCP.Tests
{
    public class SecurityValidationTests
    {
        // Tests for CreateTable
        [Fact]
        public void CreateTable_InvalidTableName_ThrowsArgumentException()
        {
            var service = new AccessInteropService();
            var fields = new List<FieldInfo>
            {
                new FieldInfo { Name = "ValidField", Type = "TEXT" }
            };

            // Use an invalid table name (e.g., containing brackets)
            var ex = Assert.Throws<ArgumentException>(() => service.CreateTable("Invalid[Table]", fields));
            Assert.Contains("Invalid table name", ex.Message);
        }

        [Fact]
        public void CreateTable_InvalidFieldName_ThrowsArgumentException()
        {
            var service = new AccessInteropService();
            var fields = new List<FieldInfo>
            {
                new FieldInfo { Name = "Invalid;Field", Type = "TEXT" }
            };

            var ex = Assert.Throws<ArgumentException>(() => service.CreateTable("ValidTable", fields));
            Assert.Contains("Invalid field name", ex.Message);
        }

        [Fact]
        public void CreateTable_InvalidFieldType_ThrowsArgumentException()
        {
            var service = new AccessInteropService();
            var fields = new List<FieldInfo>
            {
                new FieldInfo { Name = "ValidField", Type = "VARCHAR(255); DROP TABLE Users" }
            };

            var ex = Assert.Throws<ArgumentException>(() => service.CreateTable("ValidTable", fields));
            Assert.Contains("Invalid field type", ex.Message);
        }

        [Fact]
        public void CreateTable_ValidInputsWithoutConnection_ThrowsInvalidOperationException()
        {
            var service = new AccessInteropService();
            var fields = new List<FieldInfo>
            {
                new FieldInfo { Name = "ValidField", Type = "TEXT" }
            };

            // Assuming IsConnected is false initially
            var ex = Assert.Throws<InvalidOperationException>(() => service.CreateTable("ValidTable", fields));
            Assert.Contains("Not connected to database", ex.Message);
        }

        // Tests for DeleteTable
        [Fact]
        public void DeleteTable_InvalidTableName_ThrowsArgumentException()
        {
            var service = new AccessInteropService();

            // Use an invalid table name (e.g., containing quotes)
            var ex = Assert.Throws<ArgumentException>(() => service.DeleteTable("Invalid'Table"));
            Assert.Contains("Invalid table name", ex.Message);
        }

        [Fact]
        public void DeleteTable_ValidInputWithoutConnection_ThrowsInvalidOperationException()
        {
            var service = new AccessInteropService();

            // Assuming IsConnected is false initially
            var ex = Assert.Throws<InvalidOperationException>(() => service.DeleteTable("ValidTable"));
            Assert.Contains("Not connected to database", ex.Message);
        }
    }
}
