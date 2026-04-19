using System;
using System.Collections.Generic;
using MS.Access.MCP.Interop;
using Xunit;

namespace MS.Access.MCP.Tests
{
    public class AccessInteropServiceTests
    {
        [Fact]
        public void GetFullSchemaMarkdown_WithoutConnection_ThrowsInvalidOperationException()
        {
            using var accessService = new AccessInteropService();
            var exception = Assert.Throws<InvalidOperationException>(() => accessService.GetFullSchemaMarkdown());
            Assert.Contains("Not connected", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void GenerateFullSchemaMarkdown_FormatsTablesAndRelationshipsCorrectly()
        {
            // Arrange
            using var accessService = new AccessInteropService();

            var tables = new List<TableInfo>
            {
                new TableInfo
                {
                    Name = "Employees",
                    RecordCount = 50,
                    Fields = new List<FieldInfo>
                    {
                        new FieldInfo { Name = "ID", Type = "AutoNumber", Size = 4, Required = true, AllowZeroLength = false },
                        new FieldInfo { Name = "FirstName", Type = "Short Text", Size = 255, Required = true, AllowZeroLength = true }
                    }
                },
                new TableInfo
                {
                    Name = "Departments",
                    RecordCount = 5,
                    Fields = new List<FieldInfo>
                    {
                        new FieldInfo { Name = "DeptID", Type = "AutoNumber", Size = 4, Required = true, AllowZeroLength = false },
                        new FieldInfo { Name = "DeptName", Type = "Short Text", Size = 50, Required = true, AllowZeroLength = false }
                    }
                }
            };

            var relationships = new List<RelationshipInfo>
            {
                new RelationshipInfo
                {
                    Name = "fk_employee_dept",
                    Table = "Employees",
                    ForeignTable = "Departments",
                    Attributes = "Enforced"
                }
            };

            // Act
            var markdown = accessService.GenerateFullSchemaMarkdown(tables, relationships);

            // Assert
            Assert.Contains("# Database schema", markdown);
            Assert.Contains("## Tables", markdown);
            Assert.Contains("### Employees (50 rows)", markdown);
            Assert.Contains("| ID | AutoNumber | 4 | True | False |", markdown);
            Assert.Contains("| FirstName | Short Text | 255 | True | True |", markdown);

            Assert.Contains("### Departments (5 rows)", markdown);
            Assert.Contains("| DeptID | AutoNumber | 4 | True | False |", markdown);
            Assert.Contains("| DeptName | Short Text | 50 | True | False |", markdown);

            Assert.Contains("## Relationships", markdown);
            Assert.Contains("| fk_employee_dept | Employees | Departments | Enforced |", markdown);
        }

        [Fact]
        public void GenerateFullSchemaMarkdown_EscapesMarkdownCharacters()
        {
            // Arrange
            using var accessService = new AccessInteropService();

            var tables = new List<TableInfo>
            {
                new TableInfo
                {
                    Name = "Data|Table", // Includes markdown pipe
                    RecordCount = 10,
                    Fields = new List<FieldInfo>
                    {
                        new FieldInfo { Name = "Field*Name", Type = "Type_With_Underscore", Size = 50, Required = false, AllowZeroLength = true }
                    }
                }
            };

            var relationships = new List<RelationshipInfo>
            {
                new RelationshipInfo
                {
                    Name = "rel_`name`",
                    Table = "Data|Table",
                    ForeignTable = "Other_Table",
                    Attributes = ""
                }
            };

            // Act
            var markdown = accessService.GenerateFullSchemaMarkdown(tables, relationships);

            // Assert
            // Checking the table header line for escaped pipe
            Assert.Contains("### Data\\|Table", markdown);

            // Checking the field line for escaped characters
            Assert.Contains("| Field*Name | Type_With_Underscore |", markdown);

            // Checking the relationship line for escaped characters
            Assert.Contains("| rel_`name` | Data\\|Table | Other_Table |  |", markdown);
        }
    }
}
