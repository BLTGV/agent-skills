#!/usr/bin/env bun
/**
 * Microsoft Graph Authentication Script
 *
 * Authenticates with Microsoft Graph API using device code flow.
 *
 * Usage:
 *   bun run auth.ts [--profile <name>] [--client-id <id>] [--tenant-id <id>]
 *
 * Options:
 *   --profile     Credential profile name (default: "default")
 *   --client-id   Azure AD application (client) ID
 *   --tenant-id   Azure AD tenant ID (default: "common" for multi-tenant)
 *   --scopes      Comma-separated list of scopes (default: all scopes)
 *   --list        List all stored credential profiles
 *   --delete      Delete a credential profile
 *
 * Examples:
 *   bun run auth.ts
 *   bun run auth.ts --profile work
 *   bun run auth.ts --client-id your-app-id --tenant-id your-tenant-id
 *   bun run auth.ts --profile work --client-id your-app-id
 *   bun run auth.ts --list
 *   bun run auth.ts --delete --profile old-account
 */

import { parseArgs } from "util";
import { GraphClient, GRAPH_SCOPES } from "./lib/graph-client";
import {
  listProfiles,
  getCredential,
  deleteCredential,
} from "./lib/credentials";

const { values } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    "client-id": { type: "string" },
    "tenant-id": { type: "string" },
    scopes: { type: "string" },
    list: { type: "boolean", default: false },
    delete: { type: "boolean", default: false },
    help: { type: "boolean", short: "h", default: false },
  },
});

if (values.help) {
  console.log(`
Microsoft Graph Authentication

Usage:
  bun run auth.ts [options]

Options:
  --profile <name>      Credential profile name (default: "default")
  --client-id <id>      Azure AD application (client) ID
  --tenant-id <id>      Azure AD tenant ID (default: "common" for multi-tenant)
  --scopes <scopes>     Comma-separated list of scopes
  --list                List all stored credential profiles
  --delete              Delete a credential profile
  -h, --help            Show this help message

Using Your Own App Registration:
  1. Go to https://portal.azure.com
  2. Navigate to Azure Active Directory > App registrations
  3. Create a new registration:
     - Name: Your app name
     - Supported account types: Choose based on your needs
     - Redirect URI: Select "Public client/native" and add:
       https://login.microsoftonline.com/common/oauth2/nativeclient
  4. Go to "Authentication" and enable "Allow public client flows"
  5. Go to "API permissions" and add:
     - Microsoft Graph > Delegated > Mail.Read
     - Microsoft Graph > Delegated > Calendars.Read
     - Microsoft Graph > Delegated > User.Read
  6. Copy the "Application (client) ID" from Overview
  7. Use it with: --client-id your-app-id

Examples:
  # Authenticate with default (Graph Explorer) client
  bun run auth.ts

  # Authenticate with your own app
  bun run auth.ts --client-id 12345678-1234-1234-1234-123456789abc

  # Authenticate with single-tenant app
  bun run auth.ts --client-id your-app-id --tenant-id your-tenant-id

  # Use 'work' profile with custom app
  bun run auth.ts --profile work --client-id your-app-id

  # List all profiles
  bun run auth.ts --list

  # Delete a profile
  bun run auth.ts --delete --profile old
`);
  process.exit(0);
}

async function main() {
  if (values.list) {
    const profiles = await listProfiles("microsoft-graph");
    if (profiles.length === 0) {
      console.log("No credential profiles found.");
      console.log("Run 'bun run auth.ts' to create one.");
    } else {
      console.log("Microsoft Graph credential profiles:\n");
      for (const profile of profiles) {
        const cred = await getCredential("microsoft-graph", profile);
        if (cred) {
          const expired = new Date(cred.expiresAt) < new Date();
          console.log(`  ${profile}:`);
          console.log(`    Account: ${cred.account}`);
          console.log(`    Scopes: ${cred.scopes.join(", ")}`);
          if (cred.clientId) {
            console.log(`    Client ID: ${cred.clientId}`);
          }
          if (cred.tenantId) {
            console.log(`    Tenant ID: ${cred.tenantId}`);
          }
          console.log(`    Status: ${expired ? "EXPIRED" : "Valid"}`);
          console.log();
        }
      }
    }
    return;
  }

  if (values.delete) {
    const deleted = await deleteCredential("microsoft-graph", values.profile!);
    if (deleted) {
      console.log(`Deleted credential profile: ${values.profile}`);
    } else {
      console.log(`Profile not found: ${values.profile}`);
    }
    return;
  }

  // Determine scopes
  let scopes: string[];
  if (values.scopes) {
    scopes = values.scopes.split(",").map((s) => s.trim());
  } else {
    // Request all available scopes
    scopes = [
      ...GRAPH_SCOPES.user,
      ...GRAPH_SCOPES.mail,
      ...GRAPH_SCOPES.calendar,
    ];
  }

  const clientId = values["client-id"];
  const tenantId = values["tenant-id"];

  console.log(`Authenticating with Microsoft Graph...`);
  console.log(`Profile: ${values.profile}`);
  if (clientId) {
    console.log(`Client ID: ${clientId}`);
  } else {
    console.log(`Client ID: (using default Graph Explorer client)`);
  }
  if (tenantId) {
    console.log(`Tenant ID: ${tenantId}`);
  }
  console.log(`Scopes: ${scopes.join(", ")}\n`);

  const client = new GraphClient({
    profile: values.profile,
    clientId,
    tenantId,
  });

  try {
    const result = await client.authenticate(scopes);
    console.log("\n✓ Authentication successful!");
    console.log(`  Account: ${result.account?.username}`);
    console.log(`  Expires: ${result.expiresOn?.toLocaleString()}`);
    if (clientId) {
      console.log(`  Client ID saved for future use`);
    }
    console.log(`\nCredentials saved to profile: ${values.profile}`);
  } catch (error) {
    console.error("Authentication failed:", error);
    process.exit(1);
  }
}

main();
