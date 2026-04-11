using System;
using System.Collections.Generic;
using System.Reflection;
using MS.Access.MCP.Interop;
using Xunit;

namespace MS.Access.MCP.Tests
{
    public class AccessInteropServiceTests
    {
        [Fact]
        public void RefreshSchemaCache_NotConnected_ThrowsInvalidOperationException()
        {
            // Arrange
            using var service = new AccessInteropService();

            // Act & Assert
            var exception = Assert.Throws<InvalidOperationException>(() => service.RefreshSchemaCache());
            Assert.Contains("Not connected to database", exception.Message);
        }

        [Fact]
        public void RefreshSchemaCache_ClearsCacheFieldsEvenWhenDisconnected()
        {
            // Arrange
            using var service = new AccessInteropService();

            // Use reflection to set some values in the private cache fields
            var bindingFlags = BindingFlags.Instance | BindingFlags.NonPublic;

            var cachedTablesField = typeof(AccessInteropService).GetField("_cachedTables", bindingFlags);
            var cachedQueriesField = typeof(AccessInteropService).GetField("_cachedQueries", bindingFlags);
            var cachedRelationshipsField = typeof(AccessInteropService).GetField("_cachedRelationships", bindingFlags);
            var cachedSystemTablesField = typeof(AccessInteropService).GetField("_cachedSystemTables", bindingFlags);

            cachedTablesField?.SetValue(service, new List<TableInfo>());
            cachedQueriesField?.SetValue(service, new List<QueryInfo>());
            cachedRelationshipsField?.SetValue(service, new List<RelationshipInfo>());
            cachedSystemTablesField?.SetValue(service, new List<SystemTableInfo>());

            // Verify they are not null before calling RefreshSchemaCache
            Assert.NotNull(cachedTablesField?.GetValue(service));
            Assert.NotNull(cachedQueriesField?.GetValue(service));
            Assert.NotNull(cachedRelationshipsField?.GetValue(service));
            Assert.NotNull(cachedSystemTablesField?.GetValue(service));

            // Act
            // RefreshSchemaCache will throw InvalidOperationException because it's not connected,
            // but it should have cleared the fields before that.
            Assert.Throws<InvalidOperationException>(() => service.RefreshSchemaCache());

            // Assert
            Assert.Null(cachedTablesField?.GetValue(service));
            Assert.Null(cachedQueriesField?.GetValue(service));
            Assert.Null(cachedRelationshipsField?.GetValue(service));
            Assert.Null(cachedSystemTablesField?.GetValue(service));
        }
    }
}
