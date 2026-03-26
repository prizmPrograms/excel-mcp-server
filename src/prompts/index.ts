/**
 * Prompts management for Excel MCP Server
 * 
 * This module provides a centralized, extensible structure for managing
 * AI assistant guidelines (prompts). Each prompt is defined as a separate
 * module, making it easy to add, modify, or remove guidelines.
 * 
 * To add a new prompt:
 * 1. Create a new file in this directory (e.g., userform-static.ts)
 * 2. Export a `content` string with the guideline text
 * 3. Add an entry to the AVAILABLE_PROMPTS array below
 */

export interface PromptDefinition {
  name: string;
  description: string;
  getContent: () => Promise<string> | string;
}

/**
 * Available prompts for AI assistants
 * 
 * Add new prompts here as they are created.
 * The MCP server will automatically expose them to clients.
 */
export const AVAILABLE_PROMPTS: PromptDefinition[] = [
  {
    name: 'excel-vba-guidelines',
    description: 'Comprehensive guidelines for AI assistants using Excel MCP Server (advanced)',
    getContent: async () => {
      // ES Modules compatible import
      const module = await import('./guidelines-advanced.js');
      return module.content;
    },
  },
  // Future prompts can be added here:
  // {
  //   name: 'excel-userform-static',
  //   description: 'Guide for adding static user forms to Excel workbooks',
  //   getContent: () => require('./userform-static.js').content,
  // },
  // {
  //   name: 'excel-vba-guidelines-basic',
  //   description: 'Simplified guidelines for less capable AI models',
  //   getContent: () => require('./guidelines-basic.js').content,
  // },
];
