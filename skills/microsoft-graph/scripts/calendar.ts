#!/usr/bin/env bun
/**
 * Microsoft Graph Calendar Script
 *
 * View and search calendar events from Microsoft Graph API.
 * Authentication is handled automatically - if needed, returns auth instructions.
 *
 * Usage:
 *   bun run calendar.ts <command> [options]
 *
 * Commands:
 *   list      List upcoming calendar events
 *   view      View a specific event by ID
 *   search    Search calendar events
 *   today     Show today's events
 *   week      Show this week's events
 *
 * Options:
 *   --profile    Credential profile (default: "default")
 *   --top        Number of results (default: 10)
 *   --start      Start date (ISO format or relative: today, tomorrow)
 *   --end        End date (ISO format or relative: +7d, +1m)
 *   --query      Search query for 'search' command
 *   --id         Event ID for 'view' command
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
import type { GraphCalendarEvent } from "./lib/types";

const { values, positionals } = parseArgs({
  args: Bun.argv.slice(2),
  options: {
    profile: { type: "string", default: "default" },
    top: { type: "string", default: "10" },
    start: { type: "string" },
    end: { type: "string" },
    query: { type: "string" },
    id: { type: "string" },
    help: { type: "boolean", short: "h", default: false },
  },
  allowPositionals: true,
});

const command = positionals[0] || "list";

if (values.help) {
  console.log(`
Microsoft Graph Calendar Access

Usage:
  bun run calendar.ts <command> [options]

Commands:
  list      List upcoming calendar events (default)
  view      View a specific event by ID
  search    Search calendar events
  today     Show today's events
  week      Show this week's events

Options:
  --profile <name>    Credential profile (default: "default")
  --top <n>           Number of results (default: 10)
  --start <date>      Start date (ISO format or: today, tomorrow)
  --end <date>        End date (ISO format or: +7d, +1m, +1y)
  --query <q>         Search query (for 'search' command)
  --id <id>           Event ID (for 'view' command)
  -h, --help          Show this help message

Output:
  JSON with status field indicating success, auth_required, auth_pending, or error.

Date Examples:
  --start today --end +7d          Next 7 days
  --start 2024-01-01 --end 2024-01-31
  --start tomorrow

Examples:
  bun run calendar.ts                           # List upcoming events
  bun run calendar.ts today                     # Today's events
  bun run calendar.ts week                      # This week's events
  bun run calendar.ts list --start tomorrow --end +7d
  bun run calendar.ts search --query "1:1"
  bun run calendar.ts view --id AAMkAG...
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

function parseDate(input: string, baseDate: Date = new Date()): Date {
  const lower = input.toLowerCase();

  if (lower === "today") {
    const d = new Date(baseDate);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  if (lower === "tomorrow") {
    const d = new Date(baseDate);
    d.setDate(d.getDate() + 1);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  // Relative dates: +7d, +1m, +1y
  const relativeMatch = lower.match(/^\+(\d+)([dmy])$/);
  if (relativeMatch) {
    const [, amount, unit] = relativeMatch;
    const d = new Date(baseDate);
    switch (unit) {
      case "d":
        d.setDate(d.getDate() + parseInt(amount));
        break;
      case "m":
        d.setMonth(d.getMonth() + parseInt(amount));
        break;
      case "y":
        d.setFullYear(d.getFullYear() + parseInt(amount));
        break;
    }
    return d;
  }

  // ISO date
  return new Date(input);
}

function formatEventData(event: GraphCalendarEvent) {
  return {
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    isAllDay: event.isAllDay ?? false,
    location: event.location?.displayName ?? null,
    organizer: event.organizer?.emailAddress ?? null,
    attendees: event.attendees?.map((a) => ({
      email: a.emailAddress,
      status: a.status?.response ?? null,
    })) ?? [],
    bodyPreview: event.bodyPreview ?? null,
  };
}

async function main() {
  const requiredScopes = [...GRAPH_SCOPES.user, ...GRAPH_SCOPES.calendar];

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

  // Determine date range based on command
  let startDate: Date;
  let endDate: Date;

  const now = new Date();

  switch (command) {
    case "today":
      startDate = new Date(now);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(now);
      endDate.setHours(23, 59, 59, 999);
      break;

    case "week":
      startDate = new Date(now);
      startDate.setHours(0, 0, 0, 0);
      endDate = new Date(now);
      endDate.setDate(endDate.getDate() + 7);
      break;

    default:
      startDate = values.start ? parseDate(values.start) : now;
      endDate = values.end ? parseDate(values.end, startDate) : parseDate("+30d", startDate);
  }

  try {
    switch (command) {
      case "list":
      case "today":
      case "week": {
        const startISO = startDate.toISOString();
        const endISO = endDate.toISOString();

        const response = await graphRequest<{ value: GraphCalendarEvent[] }>(
          token,
          `/me/calendarView?startDateTime=${startISO}&endDateTime=${endISO}&$top=${top}&$orderby=start/dateTime`
        );

        output({
          status: "success",
          data: response.value.map(formatEventData),
        });
        break;
      }

      case "view": {
        if (!values.id) {
          output({ status: "error", error: "--id is required for 'view' command" });
          return;
        }

        const event = await graphRequest<GraphCalendarEvent>(
          token,
          `/me/events/${values.id}`
        );

        output({
          status: "success",
          data: formatEventData(event),
        });
        break;
      }

      case "search": {
        if (!values.query) {
          output({ status: "error", error: "--query is required for 'search' command" });
          return;
        }

        // Use filter instead of search for calendar (search not supported on events)
        const query = values.query.replace(/'/g, "''");
        const response = await graphRequest<{ value: GraphCalendarEvent[] }>(
          token,
          `/me/events?$filter=contains(subject,'${query}')&$top=${top}&$orderby=start/dateTime desc`
        );

        output({
          status: "success",
          data: response.value.map(formatEventData),
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
