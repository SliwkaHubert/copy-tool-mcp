import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { google } from "googleapis";
import { authenticate } from "@google-cloud/local-auth";
import * as fs from "fs";
import * as path from "path";
import * as process from "process";
import { z } from "zod";
import { docs_v1, drive_v3 } from "googleapis";
import { OAuth2Client } from "google-auth-library";

// Set up OAuth2.0 scopes - we need full access to Docs and Drive
const SCOPES = [
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.readonly" // Add read-only scope as a fallback
];

// Resolve paths relative to the project root
const PROJECT_ROOT = path.resolve(path.join(path.dirname(new URL(import.meta.url).pathname), '..'));

// Dodaj po definicji TOKEN_PATH
const SURFER_API_KEY = path.join(PROJECT_ROOT, "surfer-config.json");

// The token path is where we'll store the OAuth credentials
const TOKEN_PATH = path.join(PROJECT_ROOT, "token.json");

// The credentials path is where your OAuth client credentials are stored
const CREDENTIALS_PATH = path.join(PROJECT_ROOT, "credentials.json");

// Create an MCP server instance
const server = new McpServer({
  name: "google-docs",
  version: "1.0.0",
});

/**
 * Load saved credentials if they exist, otherwise trigger the OAuth flow
 */
async function authorize() {
  try {
    // Load client secrets from a local file
    console.error("Reading credentials from:", CREDENTIALS_PATH);
    const content = fs.readFileSync(CREDENTIALS_PATH, "utf-8");
    const keys = JSON.parse(content);
    const clientId = keys.installed.client_id;
    const clientSecret = keys.installed.client_secret;
    const redirectUri = keys.installed.redirect_uris[0];
    
    console.error("Using client ID:", clientId);
    console.error("Using redirect URI:", redirectUri);
    
    // Create an OAuth2 client
    const oAuth2Client = new OAuth2Client(clientId, clientSecret, redirectUri);
    
    // Check if we have previously stored a token
    if (fs.existsSync(TOKEN_PATH)) {
      console.error("Found existing token, attempting to use it...");
      const token = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf-8"));
      oAuth2Client.setCredentials(token);
      return oAuth2Client;
    }
    
    // No token found, use the local-auth library to get one
    console.error("No token found, starting OAuth flow...");
    const client = await authenticate({
      scopes: SCOPES,
      keyfilePath: CREDENTIALS_PATH,
    });
    
    if (client.credentials) {
      console.error("Authentication successful, saving token...");
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(client.credentials));
      console.error("Token saved successfully to:", TOKEN_PATH);
    } else {
      console.error("Authentication succeeded but no credentials returned");
    }
    
    return client;
  } catch (err) {
    console.error("Error authorizing with Google:", err);
    if (err.message) console.error("Error message:", err.message);
    if (err.stack) console.error("Stack trace:", err.stack);
    throw err;
  }
}

// Create Docs and Drive API clients
let docsClient: docs_v1.Docs;
let driveClient: drive_v3.Drive;

// Initialize Google API clients
async function initClients() {
  try {
    console.error("Starting client initialization...");
    const auth = await authorize();
    console.error("Auth completed successfully:", !!auth);
    docsClient = google.docs({ version: "v1", auth: auth as any });
    console.error("Docs client created:", !!docsClient);
    driveClient = google.drive({ version: "v3", auth: auth as any });
    console.error("Drive client created:", !!driveClient);
    return true;
  } catch (error) {
    console.error("Failed to initialize Google API clients:", error);
    return false;
  }
}

// Initialize clients when the server starts
initClients().then((success) => {
  if (!success) {
    console.error("Failed to initialize Google API clients. Server will not work correctly.");
  } else {
    console.error("Google API clients initialized successfully.");
  }
});

// RESOURCES

// Resource for listing documents
server.resource(
  "list-docs",
  "googledocs://list",
  async (uri) => {
    try {
      const response = await driveClient.files.list({
        q: "mimeType='application/vnd.google-apps.document'",
        fields: "files(id, name, createdTime, modifiedTime)",
        pageSize: 50,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });

      const files = response.data.files || [];
      let content = "Google Docs in your Drive:\n\n";
      
      if (files.length === 0) {
        content += "No Google Docs found.";
      } else {
        files.forEach((file: any) => {
          content += `Title: ${file.name}\n`;
          content += `ID: ${file.id}\n`;
          content += `Created: ${file.createdTime}\n`;
          content += `Last Modified: ${file.modifiedTime}\n\n`;
        });
      }

      return {
        contents: [{
          uri: uri.href,
          text: content,
        }]
      };
    } catch (error) {
      console.error("Error listing documents:", error);
      return {
        contents: [{
          uri: uri.href,
          text: `Error listing documents: ${error}`,
        }]
      };
    }
  }
);

// Resource to get a specific document by ID
server.resource(
  "get-doc",
  new ResourceTemplate("googledocs://{docId}", { list: undefined }),
  async (uri, { docId }) => {
    try {
      const doc = await docsClient.documents.get({
        documentId: docId as string,
      });
      
      // Extract the document content
      let content = `Document: ${doc.data.title}\n\n`;
      
      // Process the document content from the complex data structure
      const document = doc.data;
      if (document && document.body && document.body.content) {
        let textContent = "";
        
        // Loop through the document's structural elements
        document.body.content.forEach((element: any) => {
          if (element.paragraph) {
            element.paragraph.elements.forEach((paragraphElement: any) => {
              if (paragraphElement.textRun && paragraphElement.textRun.content) {
                textContent += paragraphElement.textRun.content;
              }
            });
          }
        });
        
        content += textContent;
      }

      return {
        contents: [{
          uri: uri.href,
          text: content,
        }]
      };
    } catch (error) {
      console.error(`Error getting document ${docId}:`, error);
      return {
        contents: [{
          uri: uri.href,
          text: `Error getting document ${docId}: ${error}`,
        }]
      };
    }
  }
);

// TOOLS

// Tool do pobierania słów kluczowych z Surfer SEO
server.tool(
  "get-surfer-keywords",
  {
    contentEditorId: z.number().describe("ID of the Surfer SEO Content Editor"),
  },
  async ({ contentEditorId }) => {
    try {
      // Wywołanie API Surfer SEO
      const response = await fetch(`https://app.surferseo.com/api/v1/content_editors/${contentEditorId}/terms`, {
        headers: {
          "API-KEY": SURFER_API_KEY,
          "Content-Type": "application/json",
          "Accept": "application/json"
        }
      });

      // Sprawdź czy request się udał
      if (!response.ok) {
        throw new Error(`API Error: ${response.status} - ${response.statusText}`);
      }

      const data = await response.json() as any;

      // Wyciągnij słowa kluczowe zgodnie z logiką z przykładu
      const includedTerms = [];
      const headingTerms = [];

      for (const term of data.terms) {
        if (term.included === true) {
          includedTerms.push(term.term);
        }
        if (term.use_in_heading === true) {
          headingTerms.push(term.term);
        }
      }

      // Przygotuj czytelną odpowiedź
      let result = `Słowa kluczowe z Surfer SEO (Content Editor: ${contentEditorId})\n\n`;
      
      result += `📍 Słowa do włączenia (${includedTerms.length}):\n`;
      includedTerms.forEach(term => {
        result += `• ${term}\n`;
      });
      
      result += `\n📋 Słowa do nagłówków (${headingTerms.length}):\n`;
      headingTerms.forEach(term => {
        result += `• ${term}\n`;
      });

      return {
        content: [
          {
            type: "text",
            text: result,
          },
        ],
      };

    } catch (error) {
      console.error("Błąd podczas pobierania słów kluczowych:", error);
      return {
        content: [
          {
            type: "text",
            text: `Błąd: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to create a new document
server.tool(
  "create-doc",
  {
    title: z.string().describe("The title of the new document"),
    content: z.string().optional().describe("Optional initial content for the document"),
  },
  async ({ title, content = "" }) => {
    try {
      // Create a new document
      const doc = await docsClient.documents.create({
        requestBody: {
          title: title,
        },
      });

      const documentId = doc.data.documentId;

      // If content was provided, add it to the document
      if (content) {
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: {
                    index: 1,
                  },
                  text: content,
                },
              },
            ],
          },
        });
      }

      return {
        content: [
          {
            type: "text",
            text: `Document created successfully!\nTitle: ${title}\nDocument ID: ${documentId}\nYou can now reference this document using: googledocs://${documentId}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error creating document:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to update an existing document
server.tool(
  "update-doc",
  {
    docId: z.string().describe("The ID of the document to update"),
    content: z.string().describe("The content to add to the document"),
    replaceAll: z.boolean().optional().describe("Whether to replace all content (true) or append (false)"),
  },
  async ({ docId, content, replaceAll = false }) => {
    try {
      // Ensure docId is a string and not null/undefined
      if (!docId) {
        throw new Error("Document ID is required");
      }
      
      const documentId = docId.toString();
      
      if (replaceAll) {
        // First, get the document to find its length
        const doc = await docsClient.documents.get({
          documentId,
        });
        
        // Calculate the document length
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          doc.data.body.content.forEach((element: any) => {
            if (element.paragraph) {
              element.paragraph.elements.forEach((paragraphElement: any) => {
                if (paragraphElement.textRun && paragraphElement.textRun.content) {
                  documentLength += paragraphElement.textRun.content.length;
                }
              });
            }
          });
        }
        
        // Delete all content and then insert new content
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                deleteContentRange: {
                  range: {
                    startIndex: 1,
                    endIndex: documentLength,
                  },
                },
              },
              {
                insertText: {
                  location: {
                    index: 1,
                  },
                  text: content,
                },
              },
            ],
          },
        });
      } else {
        // Append content to the end of the document
        const doc = await docsClient.documents.get({
          documentId,
        });
        
        // Calculate the document length to append at the end
        let documentLength = 1; // Start at 1 (the first character position)
        if (doc.data.body && doc.data.body.content) {
          doc.data.body.content.forEach((element: any) => {
            if (element.paragraph) {
              element.paragraph.elements.forEach((paragraphElement: any) => {
                if (paragraphElement.textRun && paragraphElement.textRun.content) {
                  documentLength += paragraphElement.textRun.content.length;
                }
              });
            }
          });
        }
        
        // Append content at the end
        await docsClient.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: {
                    index: documentLength,
                  },
                  text: content,
                },
              },
            ],
          },
        });
      }
      
      return {
        content: [
          {
            type: "text",
            text: `Document updated successfully!\nDocument ID: ${docId}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error updating document:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to search for documents
server.tool(
  "search-docs",
  {
    query: z.string().describe("The search query to find documents"),
  },
  async ({ query }) => {
    try {
      const response = await driveClient.files.list({
        q: `mimeType='application/vnd.google-apps.document' and fullText contains '${query}'`,
        fields: "files(id, name, createdTime, modifiedTime)",
        pageSize: 10,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        corpora: 'allDrives',
      });
      
      // Add response logging for debugging
      console.error("Drive API Response:", JSON.stringify(response, null, 2));
      
      // Add better response validation
      if (!response || !response.data) {
        throw new Error("Invalid response from Google Drive API");
      }
      
      // Add null check and default to empty array
      const files = (response.data.files || []);
      
      let content = `Search results for "${query}":\n\n`;
      
      if (files.length === 0) {
        content += "No documents found matching your query.";
      } else {
        files.forEach((file: any) => {
          content += `Title: ${file.name}\n`;
          content += `ID: ${file.id}\n`;
          content += `Created: ${file.createdTime}\n`;
          content += `Last Modified: ${file.modifiedTime}\n\n`;
        });
      }
      
      return {
        content: [
          {
            type: "text",
            text: content,
          },
        ],
      };
    } catch (error) {
      console.error("Error searching documents:", error);
      // Include more detailed error information
      const errorMessage = error instanceof Error 
          ? `${error.message}\n${error.stack}` 
          : String(error);
          
      return {
        content: [
          {
            type: "text",
            text: `Error searching documents: ${errorMessage}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool to delete a document
server.tool(
  "delete-doc",
  {
    docId: z.string().describe("The ID of the document to delete"),
  },
  async ({ docId }) => {
    try {
      // Get the document title first for confirmation
      const doc = await docsClient.documents.get({ documentId: docId });
      const title = doc.data.title;
      
      // Delete the document
      await driveClient.files.delete({
        fileId: docId,
      });

      return {
        content: [
          {
            type: "text",
            text: `Document "${title}" (ID: ${docId}) has been successfully deleted.`,
          },
        ],
      };
    } catch (error) {
      console.error(`Error deleting document ${docId}:`, error);
      return {
        content: [
          {
            type: "text",
            text: `Error deleting document: ${error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// PROMPTS

// Prompt for document creation
server.prompt(
  "create-doc-template",
  { 
    title: z.string().describe("The title for the new document"),
    subject: z.string().describe("The subject/topic the document should be about"),
    style: z.string().describe("The writing style (e.g., formal, casual, academic)"),
  },
  ({ title, subject, style }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please create a Google Doc with the title "${title}" about ${subject} in a ${style} writing style. Make sure it's well-structured with an introduction, main sections, and a conclusion.`
      }
    }]
  })
);

// Prompt for document analysis
server.prompt(
  "analyze-doc",
  { 
    docId: z.string().describe("The ID of the document to analyze"),
  },
  ({ docId }) => ({
    messages: [{
      role: "user",
      content: {
        type: "text",
        text: `Please analyze the content of the document with ID ${docId}. Provide a summary of its content, structure, key points, and any suggestions for improvement.`
      }
    }]
  })
);

// Connect to the transport and start the server
async function main() {
  // Create a transport for communicating over stdin/stdout
  const transport = new StdioServerTransport();

  // Connect the server to the transport
  await server.connect(transport);
  
  console.error("Google Docs MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
// Poprawione narzędzia do obsługi tabel w Google Docs MCP Server

// Tool do tworzenia tabeli w dokumencie (poprawiona wersja)
server.tool(
  "insert-table",
  {
    docId: z.string().describe("The ID of the document to insert table into"),
    rows: z.number().describe("Number of rows in the table"),
    columns: z.number().describe("Number of columns in the table"),
    insertIndex: z.number().optional().describe("Index where to insert the table (default: end of document)"),
    tableData: z.array(z.array(z.string())).optional().describe("2D array of table data [row][column]"),
  },
  async ({ docId, rows, columns, insertIndex, tableData }) => {
    try {
      const documentId = docId.toString();
      
      // Pobierz dokument, aby poprawnie obliczyć indeks
      const doc = await docsClient.documents.get({ documentId });
      
      let targetIndex = insertIndex;
      if (!targetIndex) {
        // Oblicz poprawny indeks końca dokumentu
        targetIndex = 1;
        if (doc.data.body && doc.data.body.content) {
          for (const element of doc.data.body.content) {
            if (element.endIndex) {
              targetIndex = Math.max(targetIndex, element.endIndex);
            }
          }
        }
        // Wstaw przed ostatnim znakiem dokumentu (zwykle jest to znak końca paragrafu)
        targetIndex = Math.max(1, targetIndex - 1);
      }

      console.error(`Inserting table at index: ${targetIndex}`);

      // Wstaw tabelę
      const result = await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: [{
            insertTable: {
              rows: rows,
              columns: columns,
              location: {
                index: targetIndex,
              },
            },
          }],
        },
      });

      console.error("Table inserted successfully:", result.data);

      return {
        content: [
          {
            type: "text",
            text: `Table created successfully!\nRows: ${rows}\nColumns: ${columns}\nDocument ID: ${docId}\nInserted at index: ${targetIndex}`,
          },
        ],
      };
    } catch (error) {
      console.error("Error creating table:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating table: ${error.message || error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Tool do aktualizacji konkretnej komórki tabeli (poprawiona wersja)
server.tool(
  "update-table-cell",
  {
    docId: z.string().describe("The ID of the document containing the table"),
    tableIndex: z.number().describe("Index of the table in the document (0-based)"),
    row: z.number().describe("Row index (0-based)"),
    column: z.number().describe("Column index (0-based)"),
    text: z.string().describe("Text to insert in the cell"),
    replaceContent: z.boolean().optional().describe("Whether to replace existing content (true) or append (false)"),
  },
  async ({ docId, tableIndex, row, column, text, replaceContent = true }) => {
    try {
      const documentId = docId.toString();
      
      // Pobierz dokument, aby znaleźć tabelę
      const doc = await docsClient.documents.get({ documentId });
      
      // Znajdź tabelę w dokumencie
      let currentTableIndex = 0;
      let targetTable: any = null;
      
      if (doc.data.body && doc.data.body.content) {
        for (const element of doc.data.body.content) {
          if (element.table) {
            if (currentTableIndex === tableIndex) {
              targetTable = element.table;
              break;
            }
            currentTableIndex++;
          }
        }
      }
      
      if (!targetTable) {
        throw new Error(`Table with index ${tableIndex} not found in document`);
      }
      
      // Sprawdź, czy komórka istnieje
      if (!targetTable.tableRows || !targetTable.tableRows[row]) {
        throw new Error(`Row ${row} not found in table`);
      }
      
      if (!targetTable.tableRows[row].tableCells || !targetTable.tableRows[row].tableCells[column]) {
        throw new Error(`Column ${column} not found in row ${row}`);
      }
      
      const cell = targetTable.tableRows[row].tableCells[column];
      
      // Znajdź indeks początku i końca komórki
      let cellStartIndex = cell.startIndex;
      let cellEndIndex = cell.endIndex;
      
      console.error(`Cell indices: start=${cellStartIndex}, end=${cellEndIndex}`);
      
      const requests: any[] = [];
      
      if (replaceContent) {
        // Usuń istniejącą zawartość komórki, ale zostaw znak końca komórki
        if (cellEndIndex - cellStartIndex > 1) {
          requests.push({
            deleteContentRange: {
              range: {
                startIndex: cellStartIndex,
                endIndex: cellEndIndex - 1,
              },
            },
          });
        }
        
        // Wstaw nowy tekst na początku komórki
        requests.push({
          insertText: {
            location: {
              index: cellStartIndex,
            },
            text: text,
          },
        });
      } else {
        // Dodaj tekst przed znakiem końca komórki
        requests.push({
          insertText: {
            location: {
              index: cellEndIndex - 1,
            },
            text: text,
          },
        });
      }
      
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: { requests },
      });
      
      return {
        content: [
          {
            type: "text",
            text: `Table cell updated successfully!\nTable: ${tableIndex}\nRow: ${row}\nColumn: ${column}\nText: "${text}"`,
          },
        ],
      };
    } catch (error) {
      console.error("Error updating table cell:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error updating table cell: ${error.message || error}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Nowy tool - prostsze podejście do tworzenia tabeli z danymi
server.tool(
  "create-formatted-table",
  {
    docId: z.string().describe("The ID of the document to insert table into"),
    headers: z.array(z.string()).describe("Array of column headers"),
    data: z.array(z.array(z.string())).describe("2D array of table data [row][column]"),
    insertIndex: z.number().optional().describe("Index where to insert the table (default: end of document)"),
    headerStyle: z.boolean().optional().describe("Whether to make the first row bold (header style)"),
  },
  async ({ docId, headers, data, insertIndex, headerStyle = true }) => {
    try {
      const documentId = docId.toString();
      const columns = headers.length;
      const rows = data.length + 1; // +1 for headers
      
      // Najpierw utwórz pustą tabelę
      const doc = await docsClient.documents.get({ documentId });
      
      let targetIndex = insertIndex;
      if (!targetIndex) {
        targetIndex = 1;
        if (doc.data.body && doc.data.body.content) {
          for (const element of doc.data.body.content) {
            if (element.endIndex) {
              targetIndex = Math.max(targetIndex, element.endIndex);
            }
          }
        }
        targetIndex = Math.max(1, targetIndex - 1);
      }

      console.error(`Creating formatted table at index: ${targetIndex}`);

      // 1. Wstaw pustą tabelę
      await docsClient.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: [{
            insertTable: {
              rows: rows,
              columns: columns,
              location: {
                index: targetIndex,
              },
            },
          }],
        },
      });

      // 2. Poczekaj chwilę i pobierz zaktualizowany dokument
      await new Promise(resolve => setTimeout(resolve, 500));
      const updatedDoc = await docsClient.documents.get({ documentId });
      
      // 3. Znajdź nowo utworzoną tabelę
      let newTable: any = null;
      if (updatedDoc.data.body && updatedDoc.data.body.content) {
        for (const element of updatedDoc.data.body.content) {
          if (element.table && element.startIndex && element.startIndex >= targetIndex) {
            newTable = element.table;
            break;
          }
        }
      }
      
      if (!newTable) {
        throw new Error("Could not find the newly created table");
      }

      console.error("Found new table with", newTable.tableRows?.length, "rows");

      // 4. Wypełnij nagłówki
      const fillRequests: any[] = [];
      
      // Wypełnij nagłówki (pierwszy wiersz)
      if (newTable.tableRows && newTable.tableRows[0] && newTable.tableRows[0].tableCells) {
        for (let colIndex = 0; colIndex < headers.length && colIndex < newTable.tableRows[0].tableCells.length; colIndex++) {
          const headerText = headers[colIndex];
          const cell = newTable.tableRows[0].tableCells[colIndex];
          
          if (headerText && cell && cell.startIndex !== undefined) {
            fillRequests.push({
              insertText: {
                location: {
                  index: cell.startIndex,
                },
                text: headerText,
              },
            });
            
            // Pogrub nagłówek jeśli headerStyle = true
            if (headerStyle) {
              fillRequests.push({
                updateTextStyle: {
                  range: {
                    startIndex: cell.startIndex,
                    endIndex: cell.startIndex + headerText.length,
                  },
                  textStyle: {
                    bold: true,
                  },
                  fields: "bold",
                },
              });
            }
          }
        }
      }
      
      // Wypełnij dane (pozostałe wiersze)
      for (let rowIndex = 0; rowIndex < data.length && (rowIndex + 1) < newTable.tableRows.length; rowIndex++) {
        const rowData = data[rowIndex];
        const tableRow = newTable.tableRows[rowIndex + 1]; // +1 bo pierwszy wiersz to nagłówki
        
        if (tableRow && tableRow.tableCells) {
          for (let colIndex = 0; colIndex < rowData.length && colIndex < tableRow.tableCells.length; colIndex++) {
            const cellText = rowData[colIndex];
            const cell = tableRow.tableCells[colIndex];
            
            if (cellText && cell && cell.startIndex !== undefined) {
              fillRequests.push({
                insertText: {
                  location: {
                    index: cell.startIndex,
                  },
                  text: cellText,
                },
              });
            }
          }
        }
      }

      // 5. Wykonaj wszystkie operacje wypełniania
      if (fillRequests.length > 0) {
        console.error(`Executing ${fillRequests.length} fill requests`);
        
        // Wykonuj requests w mniejszych partiach, aby uniknąć przekroczenia limitów
        const batchSize = 10;
        for (let i = 0; i < fillRequests.length; i += batchSize) {
          const batch = fillRequests.slice(i, i + batchSize);
          await docsClient.documents.batchUpdate({
            documentId,
            requestBody: { requests: batch },
          });
          
          // Krótka pauza między partiami
          if (i + batchSize < fillRequests.length) {
            await new Promise(resolve => setTimeout(resolve, 200));
          }
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Formatted table created successfully!\nRows: ${rows} (including header)\nColumns: ${columns}\nDocument ID: ${docId}\nFilled ${fillRequests.length} cells`,
          },
        ],
      };
    } catch (error) {
      console.error("Error creating formatted table:", error);
      return {
        content: [
          {
            type: "text",
            text: `Error creating formatted table: ${error.message || error}`,
          },
        ],
        isError: true,
      };
    }
  }
);