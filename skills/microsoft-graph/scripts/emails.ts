#!/usr/bin/env bun
/**
 * Microsoft Graph Email Script
 *
 * Read, list, and search emails from Microsoft Graph API.
 *
 * Usage:
 *   bun run emails.ts <command> [options]
 *
 * Commands:
 *   list      List emails from a folder
 *   read      Read a specific email by ID
 *   search    Search emails
 *   folders   List mail folders
 *
 * Options:
 *   --profile    Credential profile (default: "default")
 *   --folder     Folder name or ID (default: "inbox")
 *   --top        Number of results (default: 10)
 *   --query      Search query for 'search' command
 *   --id         Email ID for 'read' command
 *   --format     Output format: json, text (default: text)
 *
 * Examples:
 *   bun run emails.ts list
 *   bun run emails.ts list --folder "Sent Items" --top 5
 *   bun run emails.ts read --id AAMkAG...
 *   bun run emails.ts search --query "from:boss@company.com"
 *   bun run emails.ts folders
 */

import { parseArgs } from "util";
import { GraphClient, GRAPH_SCOPES } from "./lib/graph-client";
import type { GraphEmail, MailFolder } from "./lib/types";

const { values, positionals } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    folder: { type: "string", default: "inbox" },
    top: { type: "string", default: "10" },
    query: { type: "string" },
    id: { type: "string" },
    format: { type: "string", default: "text" },
    help: { type: "boolean", short: "h", default: false },
  },
  allowPositionals: true,
});

const command = positionals[0];

if (values.help || !command) {
  console.log(`
Microsoft Graph Email Access

Usage:
  bun run emails.ts <command> [options]

Commands:
  list      List emails from a folder
  read      Read a specific email by ID
  search    Search emails
  folders   List mail folders

Options:
  --profile <name>    Credential profile (default: "default")
  --folder <name>     Folder name or ID (default: "inbox")
  --top <n>           Number of results (default: 10)
  --query <q>         Search query (for 'search' command)
  --id <id>           Email ID (for 'read' command)
  --format <fmt>      Output format: json, text (default: text)
  -h, --help          Show this help message

Search Query Examples:
  from:sender@example.com
  subject:meeting
  hasAttachments:true
  received>=2024-01-01
  "exact phrase"

Examples:
  bun run emails.ts list
  bun run emails.ts list --folder "Sent Items" --top 5
  bun run emails.ts read --id AAMkAG...
  bun run emails.ts search --query "from:boss@company.com subject:urgent"
  bun run emails.ts folders
`);
  process.exit(0);
}

function formatEmail(email: GraphEmail, full: boolean = false): string {
  const date = new Date(email.receivedDateTime).toLocaleString();
  const read = email.isRead ? "" : "[UNREAD] ";
  const attach = email.hasAttachments ? " 📎" : "";

  let output = `${read}${email.subject}${attach}
  From: ${email.from.emailAddress.name} <${email.from.emailAddress.address}>
  Date: ${date}
  ID: ${email.id}`;

  if (full && email.body) {
    output += `\n\n--- Body ---\n${email.body.content}`;
  } else {
    output += `\n  Preview: ${email.bodyPreview?.substring(0, 100)}...`;
  }

  return output;
}

function formatFolder(folder: MailFolder): string {
  return `${folder.displayName}
  ID: ${folder.id}
  Unread: ${folder.unreadItemCount} / Total: ${folder.totalItemCount}
  Child folders: ${folder.childFolderCount}`;
}

async function main() {
  const client = new GraphClient({ profile: values.profile });
  const scopes = [...GRAPH_SCOPES.user, ...GRAPH_SCOPES.mail];
  const top = parseInt(values.top!, 10);

  try {
    switch (command) {
      case "list": {
        const folder = encodeURIComponent(values.folder!);
        const response = await client.graphRequest<{ value: GraphEmail[] }>(
          `/me/mailFolders/${folder}/messages?$top=${top}&$orderby=receivedDateTime desc`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(response.value, null, 2));
        } else {
          console.log(`Emails in "${values.folder}" (${response.value.length} results):\n`);
          for (const email of response.value) {
            console.log(formatEmail(email));
            console.log();
          }
        }
        break;
      }

      case "read": {
        if (!values.id) {
          console.error("Error: --id is required for 'read' command");
          process.exit(1);
        }

        const email = await client.graphRequest<GraphEmail>(
          `/me/messages/${values.id}?$select=id,subject,from,receivedDateTime,body,bodyPreview,isRead,hasAttachments`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(email, null, 2));
        } else {
          console.log(formatEmail(email, true));
        }
        break;
      }

      case "search": {
        if (!values.query) {
          console.error("Error: --query is required for 'search' command");
          process.exit(1);
        }

        const query = encodeURIComponent(values.query);
        const response = await client.graphRequest<{ value: GraphEmail[] }>(
          `/me/messages?$search="${query}"&$top=${top}&$orderby=receivedDateTime desc`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(response.value, null, 2));
        } else {
          console.log(`Search results for "${values.query}" (${response.value.length} results):\n`);
          for (const email of response.value) {
            console.log(formatEmail(email));
            console.log();
          }
        }
        break;
      }

      case "folders": {
        const response = await client.graphRequest<{ value: MailFolder[] }>(
          `/me/mailFolders?$top=50`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(response.value, null, 2));
        } else {
          console.log("Mail Folders:\n");
          for (const folder of response.value) {
            console.log(formatFolder(folder));
            console.log();
          }
        }
        break;
      }

      default:
        console.error(`Unknown command: ${command}`);
        console.error("Run 'bun run emails.ts --help' for usage");
        process.exit(1);
    }
  } catch (error) {
    console.error("Error:", error);
    process.exit(1);
  }
}

main();
