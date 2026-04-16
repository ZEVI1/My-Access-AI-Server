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
        public void RefreshSchemaCache_ClearsCachedFields()
        {
            using var accessService = new AccessInteropService();

            // Set dummy data in cache fields using reflection
            var cachedTablesField = typeof(AccessInteropService).GetField("_cachedTables", BindingFlags.NonPublic | BindingFlags.Instance);
            var cachedQueriesField = typeof(AccessInteropService).GetField("_cachedQueries", BindingFlags.NonPublic | BindingFlags.Instance);
            var cachedRelationshipsField = typeof(AccessInteropService).GetField("_cachedRelationships", BindingFlags.NonPublic | BindingFlags.Instance);
            var cachedSystemTablesField = typeof(AccessInteropService).GetField("_cachedSystemTables", BindingFlags.NonPublic | BindingFlags.Instance);

            cachedTablesField?.SetValue(accessService, new List<TableInfo>());
            cachedQueriesField?.SetValue(accessService, new List<QueryInfo>());
            cachedRelationshipsField?.SetValue(accessService, new List<RelationshipInfo>());
            cachedSystemTablesField?.SetValue(accessService, new List<SystemTableInfo>());

            // Verify fields are not null initially
            Assert.NotNull(cachedTablesField?.GetValue(accessService));
            Assert.NotNull(cachedQueriesField?.GetValue(accessService));
            Assert.NotNull(cachedRelationshipsField?.GetValue(accessService));
            Assert.NotNull(cachedSystemTablesField?.GetValue(accessService));

            try
            {
                accessService.RefreshSchemaCache();
            }
            catch (InvalidOperationException)
            {
                // Expected because GetTables() etc will throw if not connected to a database
            }

            // Verify fields are cleared
            Assert.Null(cachedTablesField?.GetValue(accessService));
            Assert.Null(cachedQueriesField?.GetValue(accessService));
            Assert.Null(cachedRelationshipsField?.GetValue(accessService));
            Assert.Null(cachedSystemTablesField?.GetValue(accessService));
        }
    }
}
