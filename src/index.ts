import { FastMCP } from "fastmcp";
import { z } from "zod";
import { execSync } from "child_process";
import { randomUUID } from "crypto";

const mcp = new FastMCP({
  name: "microsoft-planner-mcp",
  version: "1.0.0",
});

// Helper to execute az rest commands
function azRest(method: string, url: string, body?: object): string {
  const args = [`az rest --method ${method} --url "${url}"`];
  if (body) {
    const bodyJson = JSON.stringify(body).replace(/"/g, '\\"');
    args.push(`--headers "Content-Type=application/json" --body "${bodyJson}"`);
  }
  try {
    const result = execSync(args.join(" "), { encoding: "utf-8", maxBuffer: 10 * 1024 * 1024 });
    return result;
  } catch (error: any) {
    throw new Error(`az rest failed: ${error.message}`);
  }
}

// Helper to get ETag for update/delete operations
function getETag(taskId: string, isDetails: boolean = false): string {
  const url = isDetails
    ? `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`
    : `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
  const result = JSON.parse(azRest("GET", url));
  return result["@odata.etag"];
}

// Tool: List tasks for a plan
mcp.addTool({
  name: "list-tasks",
  description: "List all tasks in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/tasks`;
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: Get single task
mcp.addTool({
  name: "get-task",
  description: "Get details of a specific Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    const result = azRest("GET", url);
    return result;
  },
});

// Tool: Get task details (description, checklist, references)
mcp.addTool({
  name: "get-task-details",
  description: "Get extended task details including description and checklist",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
  }),
  execute: async ({ taskId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const result = azRest("GET", url);
    return result;
  },
});

// Tool: Create task
mcp.addTool({
  name: "create-task",
  description: "Create a new task in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The plan ID"),
    bucketId: z.string().describe("The bucket ID"),
    title: z.string().describe("Task title"),
  }),
  execute: async ({ planId, bucketId, title }) => {
    const url = "https://graph.microsoft.com/v1.0/planner/tasks";
    const body = { planId, bucketId, title };
    const result = azRest("POST", url, body);
    return result;
  },
});

// Tool: Update task (title, percentComplete, assignments, categories)
mcp.addTool({
  name: "update-task",
  description: "Update task properties (title, progress, assignments, categories). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    title: z.string().optional().describe("New title"),
    percentComplete: z.number().min(0).max(100).optional().describe("Progress 0-100"),
    assignUserId: z.string().optional().describe("User ID to assign"),
    category: z.string().optional().describe("Category to apply (category1-category25)"),
  }),
  execute: async ({ taskId, title, percentComplete, assignUserId, category }) => {
    const etag = getETag(taskId);
    const body: Record<string, any> = {};
    if (title !== undefined) body.title = title;
    if (percentComplete !== undefined) body.percentComplete = percentComplete;
    if (assignUserId) {
      body.assignments = {
        [assignUserId]: {
          "@odata.type": "#microsoft.graph.plannerAssignment",
          orderHint: " !",
        },
      };
    }
    if (category) {
      body.appliedCategories = { [category]: true };
    }

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Task updated successfully";
    } catch (error: any) {
      throw new Error(`Update failed: ${error.message}`);
    }
  },
});

// Tool: Update task details (description with GitHub links)
mcp.addTool({
  name: "update-task-details",
  description: "Update task description (use for GitHub links). Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    description: z.string().describe("Task description (supports markdown, include GitHub URLs)"),
  }),
  execute: async ({ taskId, description }) => {
    const etag = getETag(taskId, true);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify({ description })}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Task details updated successfully";
    } catch (error: any) {
      throw new Error(`Update details failed: ${error.message}`);
    }
  },
});

// Tool: Add checklist item
mcp.addTool({
  name: "add-checklist-item",
  description: "Add a checklist item (subtask) to a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    title: z.string().describe("Checklist item title"),
    isChecked: z.boolean().optional().default(false).describe("Whether the item is checked"),
  }),
  execute: async ({ taskId, title, isChecked }) => {
    const etag = getETag(taskId, true);
    const itemId = randomUUID();
    const body = {
      checklist: {
        [itemId]: {
          "@odata.type": "#microsoft.graph.plannerChecklistItem",
          title,
          isChecked,
        },
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || JSON.stringify({ success: true, itemId, title });
    } catch (error: any) {
      throw new Error(`Add checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Add multiple checklist items at once
mcp.addTool({
  name: "add-checklist-items",
  description: "Add multiple checklist items (subtasks) to a Planner task in one operation",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    items: z.array(z.string()).describe("Array of checklist item titles"),
  }),
  execute: async ({ taskId, items }) => {
    const etag = getETag(taskId, true);
    const checklist: Record<string, any> = {};

    for (const title of items) {
      const itemId = randomUUID();
      checklist[itemId] = {
        "@odata.type": "#microsoft.graph.plannerChecklistItem",
        title,
        isChecked: false,
      };
    }

    const body = { checklist };
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || JSON.stringify({ success: true, itemCount: items.length });
    } catch (error: any) {
      throw new Error(`Add checklist items failed: ${error.message}`);
    }
  },
});

// Tool: Update checklist item (toggle or rename)
mcp.addTool({
  name: "update-checklist-item",
  description: "Update a checklist item (toggle checked state or rename)",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    itemId: z.string().describe("The checklist item ID (from get-task-details)"),
    title: z.string().optional().describe("New title for the item"),
    isChecked: z.boolean().optional().describe("Set checked state"),
  }),
  execute: async ({ taskId, itemId, title, isChecked }) => {
    const etag = getETag(taskId, true);
    const itemUpdate: Record<string, any> = {
      "@odata.type": "#microsoft.graph.plannerChecklistItem",
    };
    if (title !== undefined) itemUpdate.title = title;
    if (isChecked !== undefined) itemUpdate.isChecked = isChecked;

    const body = {
      checklist: {
        [itemId]: itemUpdate,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Checklist item updated successfully";
    } catch (error: any) {
      throw new Error(`Update checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Delete checklist item
mcp.addTool({
  name: "delete-checklist-item",
  description: "Delete a checklist item from a Planner task",
  parameters: z.object({
    taskId: z.string().describe("The task ID"),
    itemId: z.string().describe("The checklist item ID to delete"),
  }),
  execute: async ({ taskId, itemId }) => {
    const etag = getETag(taskId, true);
    const body = {
      checklist: {
        [itemId]: null,
      },
    };

    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}/details`;
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method PATCH --url "${url}"`,
      `--headers "Content-Type=application/json" "If-Match=${escapedEtag}"`,
      `--body '${JSON.stringify(body)}'`,
    ];
    try {
      const result = execSync(args.join(" "), { encoding: "utf-8" });
      return result || "Checklist item deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete checklist item failed: ${error.message}`);
    }
  },
});

// Tool: Delete task
mcp.addTool({
  name: "delete-task",
  description: "Delete a Planner task. Auto-fetches ETag.",
  parameters: z.object({
    taskId: z.string().describe("The task ID to delete"),
  }),
  execute: async ({ taskId }) => {
    const etag = getETag(taskId);
    const url = `https://graph.microsoft.com/v1.0/planner/tasks/${taskId}`;
    // Escape the inner quotes in the ETag for shell: W/"..." -> W/\"...\"
    const escapedEtag = etag.replace(/"/g, '\\"');
    const args = [
      `az rest --method DELETE --url "${url}"`,
      `--headers "If-Match=${escapedEtag}"`,
    ];
    try {
      execSync(args.join(" "), { encoding: "utf-8" });
      return "Task deleted successfully";
    } catch (error: any) {
      throw new Error(`Delete failed: ${error.message}`);
    }
  },
});

// Tool: List buckets for a plan
mcp.addTool({
  name: "list-buckets",
  description: "List all buckets in a Planner plan",
  parameters: z.object({
    planId: z.string().describe("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`;
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: List plans for current user
mcp.addTool({
  name: "list-plans",
  description: "List all Planner plans accessible to the current user",
  parameters: z.object({}),
  execute: async () => {
    const url = "https://graph.microsoft.com/v1.0/me/planner/plans";
    const result = JSON.parse(azRest("GET", url));
    return JSON.stringify(result.value, null, 2);
  },
});

mcp.start({ transportType: "stdio" });
