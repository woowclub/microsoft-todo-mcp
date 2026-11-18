import { z } from "zod";
import { DateTimeTimeZone, Importance, TaskStatus } from "../lib/schema";

// Zod strips unrecognized Graph properties before serializing resources to MCP.
export const TaskListSchema = z
  .object({
    id: z.string(),
    displayName: z.string(),
    isOwner: z.boolean().optional(),
    isShared: z.boolean().optional(),
    wellknownListName: z.string().optional(),
  });
export type TaskList = z.infer<typeof TaskListSchema>;

export const ChecklistItemSchema = z
  .object({
    id: z.string(),
    displayName: z.string(),
    isChecked: z.boolean().optional(),
    createdDateTime: z.string().optional(),
  });
export type ChecklistItem = z.infer<typeof ChecklistItemSchema>;

export const TaskSchema = z
  .object({
    id: z.string(),
    title: z.string(),
    status: TaskStatus.optional(),
    importance: Importance.optional(),
    isReminderOn: z.boolean().optional(),
    createdDateTime: z.string().optional(),
    lastModifiedDateTime: z.string().optional(),
    body: z.object({ content: z.string(), contentType: z.string() }).optional(),
    dueDateTime: DateTimeTimeZone.nullish(),
    reminderDateTime: DateTimeTimeZone.nullish(),
    completedDateTime: DateTimeTimeZone.nullish(),
    categories: z.array(z.string()).optional(),
    linkedResources: z.array(z.object({
      id: z.string(),
      applicationName: z.string().optional(),
      displayName: z.string().optional(),
      webUrl: z.string().nullable().optional(),
      externalId: z.string().nullable().optional(),
    })).optional(),
    checklistItems: z.array(ChecklistItemSchema).optional(),
  });
export type Task = z.infer<typeof TaskSchema>;
