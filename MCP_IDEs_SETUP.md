# הגדרת MCP IDEs

הקובץ הזה מספק דוגמאות חיבור עבור IDEs של AI שמחוברים לשרת MCP.

## Windsurf

ב־Windsurf השתמש בקובץ ההגדרות `~/.codeium/windsurf/mcp_config.json` או בהגדרות המקומיות של הסביבה.

```json
{
  "name": "Access MCP Server",
  "command": "dotnet",
  "args": [
    "C:\\My-Access-AI-Server\\publish\\MS.Access.MCP.Official.dll"
  ],
  "cwd": "C:\\My-Access-AI-Server\\publish",
  "protocol": "stdio",
  "env": {
    "DOTNET_PRINT_TELEMETRY_MESSAGE": "false"
  }
}
```

> ודא שהנתיב מתייחס ל־`MS.Access.MCP.Official.dll` המפורסם בתיקיית `publish`.

## Cursor

ב־Cursor ניתן להגדיר חיבור MCP דרך ה־UI של התוסף או דרך הגדרות משתמש.

- `server command`: `dotnet`
- `server args`: `C:\My-Access-AI-Server\publish\MS.Access.MCP.Official.dll`
- `working directory`: `C:\My-Access-AI-Server\publish`
- `protocol`: `stdio`

אם יש אפשרות לכתוב קובץ JSON, השתמש במבנה דומה לזה:

```json
{
  "command": "dotnet",
  "args": [
    "C:\\My-Access-AI-Server\\publish\\MS.Access.MCP.Official.dll"
  ],
  "cwd": "C:\\My-Access-AI-Server\\publish",
  "protocol": "stdio"
}
```

## Roo Code / Cline

ב־Roo Code / Cline השתמש בקובץ `cline_mcp_settings.json` או בהגדרות חיבור MCP של הסביבה.

```json
{
  "name": "Access MCP Server",
  "type": "stdio",
  "command": "dotnet",
  "args": [
    "C:\\My-Access-AI-Server\\publish\\MS.Access.MCP.Official.dll"
  ],
  "cwd": "C:\\My-Access-AI-Server\\publish"
}
```

> הפעל את השרת בחיבור מסוג `stdio` כדי למנוע זיהום של פלט סטנדרטי שאינו JSON-RPC.
