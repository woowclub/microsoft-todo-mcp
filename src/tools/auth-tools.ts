import { LOGIN_HINT } from "../auth/auth-state";
import { defineTool } from "./_shared";

export const authTools = [
  defineTool({
    name: "auth_status",
    description: "Show the local sign-in state and token expiry.",
    risk: "read",
    input: {},
    handler: async (_args, ctx) => ctx.auth.status(),
  }),
  defineTool({
    name: "sign_in",
    description: "Show instructions for signing in from your own terminal.",
    risk: "read",
    input: {},
    handler: async (_args, ctx) => ({ ...ctx.auth.status(), hint: LOGIN_HINT }),
  }),
];
