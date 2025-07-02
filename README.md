# Google Docs MCP Server

This is a Model Context Protocol (MCP) server that allows you to connect to Google Docs through Claude. With this server, you can:

- List all Google Docs in your Drive
- Read the content of specific documents
- Create new documents
- Update existing documents
- Search for documents
- Delete documents
- Formatting tables
- **Get SEO keywords from Surfer SEO**

## Prerequisites

- Node.js v16.0.0 or later
- Google Cloud project with the Google Docs API and Google Drive API enabled
- OAuth 2.0 credentials for your Google Cloud project
- **Surfer SEO API key (optional, for SEO features)**


## Setup

1. Clone this repository and navigate to the project directory:

```bash
git clone https://github.com/SliwkaHubert/copy-tool-mcp.git
cd copy-tool-mcp
```

2. Install dependencies:

```bash
npm install
```

3. **Install additional dependencies for Surfer SEO integration:**

```bash
npm install node-fetch
npm install @types/node-fetch --save-dev
```

4. Create an OAuth 2.0 client ID in the Google Cloud Console:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Docs API and Google Drive API
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth client ID"
   - Select "Desktop app" for the application type
   - Download the JSON file and save it as `credentials.json` in your project directory

   > **Important**: The `credentials.json` and `token.json` files contain sensitive information and are excluded from version control via `.gitignore`. Never commit these files to your repository.

5. Surfer SEO API key:
    - Create surfer-config.json file in main catalog:
    {
      "apiKey": "yoursurferapi"
    }

6. Build the project:

```bash
npm run build
```

7. Run the server:

```bash
npm start
```

The first time you run the server, it will prompt you to authenticate with Google. Follow the on-screen instructions to authorize the application. This will generate a `token.json` file that stores your access tokens.

## Security Considerations

- **Credential Security**: Both `credentials.json` and `token.json` contain sensitive information and should never be shared or committed to version control. They are already added to the `.gitignore` file.
- **Surfer SEO API Key**: Store your Surfer SEO API key securely in the source code or environment variables.
- **Token Refresh**: The application automatically refreshes the access token when it expires.
- **Revoking Access**: If you need to revoke access, delete the `token.json` file and go to your [Google Account Security settings](https://myaccount.google.com/security) to remove the app from your authorized applications.

## Connecting to Claude for Desktop

To use this server with Claude for Desktop:

1. Edit your Claude Desktop configuration file:
   - On macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - On Windows: `%APPDATA%\Claude\claude_desktop_config.json`

2. Add the following to your configuration:

```json
{
  "mcpServers": {
    "googledocs": {
      "command": "node",
      "args": ["/absolute/path/to/build/server.js"]
    }
  }
}
```

Replace `/absolute/path/to/build/server.js` with the actual path to your built server.js file.

3. Restart Claude for Desktop.

## Development

### Project Structure

```
google-docs-integration/
├── build/                # Compiled JavaScript files
├── src/                  # TypeScript source code
│   └── server.ts         # Main server implementation
├── .gitignore            # Git ignore file
├── credentials.json      # OAuth 2.0 credentials (not in version control)
├── package.json          # Project dependencies and scripts
├── README.md             # Project documentation
├── token.json            # OAuth tokens (not in version control)
└── tsconfig.json         # TypeScript configuration
```

### Adding New Features

To add new features to the MCP server:

1. Modify the `src/server.ts` file to implement new functionality
2. Build the project with `npm run build`
3. Test your changes by running `npm start`

## Available Resources

- `googledocs://list` - Lists all Google Docs in your Drive
- `googledocs://{docId}` - Gets the content of a specific document by ID

## Available Tools

### Google Docs Tools
- `create-doc` - Creates a new Google Doc with the specified title and optional content
- `update-doc` - Updates an existing Google Doc with new content (append or replace)
- `search-docs` - Searches for Google Docs containing specific text
- `delete-doc` - Deletes a Google Doc by ID
- `insert-table` - Creates tables in Google Docs
- `update-table-cell` - Updates specific table cells
- `create-formatted-table` - Creates formatted tables with headers and data

### Surfer SEO Tools
- `get-surfer-keywords` - Gets SEO keywords from Surfer SEO Content Editor

## Available Prompts

- `create-doc-template` - Helps create a new document based on a specified topic and writing style
- `analyze-doc` - Analyzes the content of a document and provides a summary

## Usage Examples

Here are some example prompts you can use with Claude once the server is connected:

### Google Docs Examples
- "Show me a list of all my Google Docs"
- "Create a new Google Doc titled 'Meeting Notes' with the content 'Topics to discuss: ...'"
- "Update my document with ID '1abc123def456' to add this section at the end: ..."
- "Search my Google Docs for any documents containing 'project proposal'"
- "Delete the Google Doc with ID '1abc123def456'"
- "Create a formal document about climate change"
- "Analyze the content of document with ID '1abc123def456'"

### Surfer SEO Examples
- "Get SEO keywords from Surfer SEO Content Editor with ID 12345"
- "Show me the included terms and heading terms for Content Editor 67890"

## Dependencies

This project uses the following key dependencies:

- `@modelcontextprotocol/sdk` - MCP SDK for server implementation
- `googleapis` - Google APIs client library
- `@google-cloud/local-auth` - Google authentication
- `node-fetch` - HTTP client for API requests (used for Surfer SEO integration)
- `zod` - Schema validation

## Troubleshooting

If you encounter authentication issues:
1. Delete the `token.json` file in your project directory
2. Run the server again to trigger a new authentication flow

If you're having trouble with the Google Docs API:
1. Make sure the API is enabled in your Google Cloud Console
2. Check that your OAuth credentials have the correct scopes

If you encounter issues with Surfer SEO integration:
1. Verify your Surfer SEO API key is correct
2. Make sure `node-fetch` is properly installed
3. Check that the Content Editor ID exists and is in "completed" state

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin feature/your-feature-name`
5. Submit a pull request

## License

MIT