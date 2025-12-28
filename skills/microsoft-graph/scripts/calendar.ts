#!/usr/bin/env bun
/**
 * Microsoft Graph Calendar Script
 *
 * View and search calendar events from Microsoft Graph API.
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
 *   --format     Output format: json, text (default: text)
 *
 * Examples:
 *   bun run calendar.ts list
 *   bun run calendar.ts today
 *   bun run calendar.ts week
 *   bun run calendar.ts list --start tomorrow --end +7d
 *   bun run calendar.ts search --query "team meeting"
 *   bun run calendar.ts view --id AAMkAG...
 */

import { parseArgs } from "util";
import { GraphClient, GRAPH_SCOPES } from "./lib/graph-client";
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
    format: { type: "string", default: "text" },
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
  --format <fmt>      Output format: json, text (default: text)
  -h, --help          Show this help message

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

function formatEvent(event: GraphCalendarEvent, full: boolean = false): string {
  const start = new Date(event.start.dateTime + "Z");
  const end = new Date(event.end.dateTime + "Z");

  const dateStr = event.isAllDay
    ? start.toLocaleDateString()
    : `${start.toLocaleString()} - ${end.toLocaleTimeString()}`;

  let output = `${event.subject}
  When: ${dateStr}${event.isAllDay ? " (All Day)" : ""}`;

  if (event.location?.displayName) {
    output += `\n  Where: ${event.location.displayName}`;
  }

  if (event.organizer) {
    output += `\n  Organizer: ${event.organizer.emailAddress.name} <${event.organizer.emailAddress.address}>`;
  }

  output += `\n  ID: ${event.id}`;

  if (full) {
    if (event.attendees && event.attendees.length > 0) {
      output += `\n  Attendees:`;
      for (const attendee of event.attendees) {
        const status = attendee.status?.response || "none";
        output += `\n    - ${attendee.emailAddress.name} <${attendee.emailAddress.address}> (${status})`;
      }
    }

    if (event.bodyPreview) {
      output += `\n\n--- Description ---\n${event.bodyPreview}`;
    }
  }

  return output;
}

async function main() {
  const client = new GraphClient({ profile: values.profile });
  const scopes = [...GRAPH_SCOPES.user, ...GRAPH_SCOPES.calendar];
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

        const response = await client.graphRequest<{ value: GraphCalendarEvent[] }>(
          `/me/calendarView?startDateTime=${startISO}&endDateTime=${endISO}&$top=${top}&$orderby=start/dateTime`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(response.value, null, 2));
        } else {
          const rangeStr =
            command === "today"
              ? "Today"
              : command === "week"
              ? "This Week"
              : `${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()}`;

          console.log(`Calendar Events (${rangeStr}) - ${response.value.length} results:\n`);
          for (const event of response.value) {
            console.log(formatEvent(event));
            console.log();
          }
        }
        break;
      }

      case "view": {
        if (!values.id) {
          console.error("Error: --id is required for 'view' command");
          process.exit(1);
        }

        const event = await client.graphRequest<GraphCalendarEvent>(
          `/me/events/${values.id}`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(event, null, 2));
        } else {
          console.log(formatEvent(event, true));
        }
        break;
      }

      case "search": {
        if (!values.query) {
          console.error("Error: --query is required for 'search' command");
          process.exit(1);
        }

        // Use filter instead of search for calendar (search not supported on events)
        const query = values.query.replace(/'/g, "''");
        const response = await client.graphRequest<{ value: GraphCalendarEvent[] }>(
          `/me/events?$filter=contains(subject,'${query}')&$top=${top}&$orderby=start/dateTime desc`,
          scopes
        );

        if (values.format === "json") {
          console.log(JSON.stringify(response.value, null, 2));
        } else {
          console.log(`Search results for "${values.query}" (${response.value.length} results):\n`);
          for (const event of response.value) {
            console.log(formatEvent(event));
            console.log();
          }
        }
        break;
      }

      default:
        console.error(`Unknown command: ${command}`);
        console.error("Run 'bun run calendar.ts --help' for usage");
        process.exit(1);
    }
  } catch (error) {
    console.error("Error:", error);
    process.exit(1);
  }
}

main();
