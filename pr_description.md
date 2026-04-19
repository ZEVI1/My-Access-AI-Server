⚡ Prevent N+1 Query in GetSystemTables Record Counts

💡 **What:**
- Changed `RecordCount` in `SystemTableInfo` to be nullable (`long?`).
- Updated `GetSystemTables` to accept an optional `includeCounts = false` parameter.
- Skips fetching record counts by default during initial schema caching (`RefreshSchemaCache`), eliminating an N+1 query issue.
- Lazy loads `RecordCount` into cache if requested later via the tool parameter `include_counts`.

🎯 **Why:**
Inside the loop iterating over `MSysObjects`, `GetTableRecordCount` was unconditionally called, issuing a `SELECT COUNT(*)` for *every* system table. System tables can be numerous, and issuing individual count queries for each one adds significant overhead during database connection and schema refreshes. By making this optional and lazy-loaded, we prevent unnecessary N+1 queries by default, dramatically speeding up connection times and normal schema reading.

📊 **Measured Improvement:**
Since this project requires a Windows environment to run COM-based operations and test OleDb connections, I was unable to compile and profile benchmarks natively on this Linux-based agent. However, mathematically, eliminating the unconditional `SELECT COUNT(*)` for every single system table reduces the number of queries required for schema caching from `O(N)` (where N is the number of system tables) to `O(1)`, removing potentially dozens of individual database operations on connect.
