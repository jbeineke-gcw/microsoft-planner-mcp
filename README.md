# microsoft-planner-mcp

Lightweight MCP (Model Context Protocol) server for Microsoft Planner that uses `az rest` for authentication instead of complex OAuth flows.

## What This Is

A minimal MCP server that enables Claude Code to interact with Microsoft Planner tasks directly. It leverages Azure CLI's existing authentication (`az login`) to make Graph API calls, eliminating the need for app registrations or token management.

## Prerequisites

- **Node.js** (v18 or later)
- **Azure CLI** installed and authenticated:
  ```bash
  az login
  ```

## Installation

```bash
git clone https://github.com/jbeineke-gcw/microsoft-planner-mcp.git
cd microsoft-planner-mcp
npm install
npm run build
```

## Usage with Claude Code

Add the MCP server to Claude Code:

```bash
claude mcp add microsoft-planner-mcp node /path/to/microsoft-planner-mcp/dist/index.js
```

Or add directly to your MCP settings file:

```json
{
  "mcpServers": {
    "microsoft-planner-mcp": {
      "command": "node",
      "args": ["/path/to/microsoft-planner-mcp/dist/index.js"]
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `list-plans` | List all Planner plans accessible to the current user |
| `list-buckets` | List all buckets in a plan |
| `list-tasks` | List all tasks in a Planner plan |
| `get-task` | Get details of a specific task |
| `get-task-details` | Get extended task details (description, checklist, references) |
| `create-task` | Create a new task in a plan |
| `update-task` | Update task properties (title, progress, assignments, categories) |
| `update-task-details` | Update task description (supports GitHub links) |
| `delete-task` | Delete a Planner task |
| `add-checklist-item` | Add a single checklist item (subtask) to a task |
| `add-checklist-items` | Add multiple checklist items in one operation |
| `update-checklist-item` | Update a checklist item (toggle checked or rename) |
| `delete-checklist-item` | Remove a checklist item from a task |

## Claude Code Agent (Optional)

For enhanced automation, create a Claude Code agent at `~/.claude/agents/microsoft-planner.md` with:

- **Auto-assignment**: Automatically assign tasks to a default user
- **Auto-labeling**: Apply default categories/labels to tasks
- **Status Intelligence**: Infer task bucket and progress from conversation context
- **GitHub Integration**: Include repository links in task descriptions

See the [agent template](https://github.com/vyente-ruffin/microsoft-planner-mcp/wiki/Agent-Configuration) for configuration details.

## Finding Your IDs

To use the MCP tools, you'll need your Planner Plan ID and Bucket IDs:

1. Use `list-plans` to discover all accessible plans and their IDs
2. Use `list-buckets` with your Plan ID to get bucket IDs
3. Or find your Plan ID from the Planner web URL: `https://tasks.office.com/...planId=YOUR_PLAN_ID`

## Example Usage

Once configured, use natural language with Claude Code:

```
"List all my planner plans"
"List all tasks in the planner"
"Create a task called 'Review PR #123' in the backlog bucket"
"Mark task XYZ as complete"
"Add a description with the GitHub PR link to the task"
"Add a checklist with: design, implement, test, document"
"Check off the 'design' item on that task"
```

## How It Works

This server uses `az rest` to make Microsoft Graph API calls. The Azure CLI handles all authentication, so as long as you're logged in with `az login`, the server can access Planner data your account has permissions for.

All Planner operations that require ETags (update, delete) automatically fetch the current ETag before making changes.

## License

ISC
