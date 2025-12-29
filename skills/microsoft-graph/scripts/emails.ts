#!/usr/bin/env bun
/**
 * Microsoft Graph Email Script
 *
 * Read, list, and search emails from Microsoft Graph API.
 * Authentication is handled automatically - if needed, returns auth instructions.
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
 *
 * Output:
 *   Always JSON with structure:
 *   - Success: { "status": "success", "data": [...] }
 *   - Auth needed: { "status": "auth_required", "userCode": "...", "verificationUri": "..." }
 *   - Error: { "status": "error", "error": "..." }
 */

import { parseArgs } from "util";
import { GRAPH_SCOPES } from "./lib/graph-client";
import { ensureAuth, type ScriptResponse } from "./lib/auth-handler";
import type { GraphEmail, MailFolder } from "./lib/types";

const { values, positionals } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    folder: { type: "string", default: "inbox" },
    top: { type: "string", default: "10" },
    query: { type: "string" },
    id: { type: "string" },
    help: { type: "boolean", short: "h", default: false },
  },
  allowPositionals: true,
});

const command = positionals[0];

if (values.help) {
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
  -h, --help          Show this help message

Output:
  JSON with status field indicating success, auth_required, auth_pending, or error.

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

function output<T>(response: ScriptResponse<T>): void {
  console.log(JSON.stringify(response));
  process.exit(response.status === "success" ? 0 : 1);
}

async function graphRequest<T>(token: string, endpoint: string): Promise<T> {
  const url = `https://graph.microsoft.com/v1.0${endpoint}`;

  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Graph API error: ${response.status} - ${error}`);
  }

  return response.json();
}

async function main() {
  if (!command) {
    output({ status: "error", error: "No command specified. Use --help for usage." });
    return;
  }

  const requiredScopes = [...GRAPH_SCOPES.user, ...GRAPH_SCOPES.mail];

  // Ensure we're authenticated
  const auth = await ensureAuth({
    service: "microsoft-graph",
    profile: values.profile!,
    requiredScopes,
  });

  if (!auth.ok) {
    output(auth.response);
    return;
  }

  const token = auth.token;
  const top = parseInt(values.top!, 10);

  try {
    switch (command) {
      case "list": {
        const folder = encodeURIComponent(values.folder!);
        const response = await graphRequest<{ value: GraphEmail[] }>(
          token,
          `/me/mailFolders/${folder}/messages?$top=${top}&$orderby=receivedDateTime desc`
        );

        output({
          status: "success",
          data: response.value.map((email) => ({
            id: email.id,
            subject: email.subject,
            from: email.from.emailAddress,
            receivedDateTime: email.receivedDateTime,
            bodyPreview: email.bodyPreview,
            isRead: email.isRead,
            hasAttachments: email.hasAttachments,
          })),
        });
        break;
      }

      case "read": {
        if (!values.id) {
          output({ status: "error", error: "--id is required for 'read' command" });
          return;
        }

        const email = await graphRequest<GraphEmail>(
          token,
          `/me/messages/${values.id}?$select=id,subject,from,receivedDateTime,body,bodyPreview,isRead,hasAttachments`
        );

        output({
          status: "success",
          data: {
            id: email.id,
            subject: email.subject,
            from: email.from.emailAddress,
            receivedDateTime: email.receivedDateTime,
            body: email.body,
            bodyPreview: email.bodyPreview,
            isRead: email.isRead,
            hasAttachments: email.hasAttachments,
          },
        });
        break;
      }

      case "search": {
        if (!values.query) {
          output({ status: "error", error: "--query is required for 'search' command" });
          return;
        }

        const query = encodeURIComponent(values.query);
        const response = await graphRequest<{ value: GraphEmail[] }>(
          token,
          `/me/messages?$search="${query}"&$top=${top}&$orderby=receivedDateTime desc`
        );

        output({
          status: "success",
          data: response.value.map((email) => ({
            id: email.id,
            subject: email.subject,
            from: email.from.emailAddress,
            receivedDateTime: email.receivedDateTime,
            bodyPreview: email.bodyPreview,
            isRead: email.isRead,
            hasAttachments: email.hasAttachments,
          })),
        });
        break;
      }

      case "folders": {
        const response = await graphRequest<{ value: MailFolder[] }>(
          token,
          `/me/mailFolders?$top=50`
        );

        output({
          status: "success",
          data: response.value.map((folder) => ({
            id: folder.id,
            displayName: folder.displayName,
            unreadItemCount: folder.unreadItemCount,
            totalItemCount: folder.totalItemCount,
            childFolderCount: folder.childFolderCount,
          })),
        });
        break;
      }

      default:
        output({ status: "error", error: `Unknown command: ${command}. Use --help for usage.` });
    }
  } catch (error) {
    output({
      status: "error",
      error: error instanceof Error ? error.message : "Unknown error",
    });
  }
}

main();
