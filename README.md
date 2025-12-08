# planner-mcp-lite

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
git clone https://github.com/405network/planner-mcp-lite.git
cd planner-mcp-lite
npm install
npm run build
```

## Usage with Claude Code

Add the MCP server to Claude Code:

```bash
claude mcp add planner-lite node /path/to/planner-mcp-lite/dist/index.js
```

Or add directly to your MCP settings file:

```json
{
  "mcpServers": {
    "planner-lite": {
      "command": "node",
      "args": ["/path/to/planner-mcp-lite/dist/index.js"]
    }
  }
}
```

## Available Tools

| Tool | Description |
|------|-------------|
| `list-tasks` | List all tasks in a Planner plan |
| `get-task` | Get details of a specific task |
| `get-task-details` | Get extended task details (description, checklist, references) |
| `create-task` | Create a new task in a plan |
| `update-task` | Update task properties (title, progress, assignments, categories) |
| `update-task-details` | Update task description (supports GitHub links) |
| `delete-task` | Delete a Planner task |
| `list-buckets` | List all buckets in a plan |

## Reference IDs for 405network Tenant

### Plan IDs
| Plan Name | Plan ID |
|-----------|---------|
| Your Plan Name | `your-plan-id-here` |

### Bucket IDs
| Bucket Name | Bucket ID |
|-------------|-----------|
| Your Bucket Name | `your-bucket-id-here` |

### User IDs
| User | User ID |
|------|---------|
| Your User | `your-user-id-here` |

## Example Usage

Once configured, use natural language with Claude Code:

```
"List all tasks in the planner"
"Create a task called 'Review PR #123' in the backlog bucket"
"Mark task XYZ as complete"
"Add a description with the GitHub PR link to the task"
```

## How It Works

This server uses `az rest` to make Microsoft Graph API calls. The Azure CLI handles all authentication, so as long as you're logged in with `az login`, the server can access Planner data your account has permissions for.

All Planner operations that require ETags (update, delete) automatically fetch the current ETag before making changes.

## License

ISC
