import { z } from 'zod';
import { getGraphClient } from '../api/graph-client.mjs';
import { createToolHandler } from '../api/retry.mjs';
import { apiCache, CacheKeys } from '../api/cache.mjs';
import { createPageHtml } from '../utils/html.mjs';
import { validateAndFetchResource } from '../utils/validation.mjs';

/**
 * Registers creation-related tools with the MCP server.
 * @param {McpServer} server - The MCP server instance.
 */
export function registerCreateTools(server) {
  server.tool(
    'createPage',
    {
      title: z
        .string()
        .min(1, { message: 'Title cannot be empty.' })
        .describe('The title for the new page.'),
      content: z
        .string()
        .min(1, { message: 'Content cannot be empty.' })
        .describe('The content for the new page (HTML or markdown-style).'),
    },
    createToolHandler(async ({ title, content }) => {
      const graphClient = getGraphClient();
      console.error(`Attempting to create page with title: "${title}"`);

      const sectionsResponse = await graphClient.api('/me/onenote/sections').get();
      if (!sectionsResponse.value || sectionsResponse.value.length === 0) {
        throw new Error('No sections found in your OneNote. Cannot create a page.');
      }
      const targetSectionId = sectionsResponse.value[0].id;
      const targetSectionName = sectionsResponse.value[0].displayName;

      const pageHtml = createPageHtml(title, content);

      const response = await graphClient
        .api(`/me/onenote/sections/${targetSectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);

      return {
        content: [
          {
            type: 'text',
            text: `✅ **Page Created Successfully!**
**Title:** ${response.title}
**Page ID:** ${response.id}
**In Section:** ${targetSectionName}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`,
          },
        ],
      };
    }, 'Error creating page')
  );

  server.tool(
    'createPageInSection',
    {
      sectionId: z
        .string()
        .min(1, { message: 'Section ID cannot be empty.' })
        .describe('The ID of the section to create the page in.'),
      title: z
        .string()
        .min(1, { message: 'Title cannot be empty.' })
        .describe('The title for the new page.'),
      content: z
        .string()
        .min(1, { message: 'Content cannot be empty.' })
        .describe('The content for the new page (HTML or markdown-style).'),
    },
    createToolHandler(async ({ sectionId, title, content }) => {
      const graphClient = getGraphClient();
      const { id: validatedSectionId, resource: sectionInfo } = await validateAndFetchResource(
        sectionId,
        'section',
        `/me/onenote/sections/${sectionId}`,
        'listSections or searchSections'
      );
      const targetSectionName = sectionInfo.displayName;

      console.error(
        `Attempting to create page with title: "${title}" in section: ${validatedSectionId}`
      );

      const pageHtml = createPageHtml(title, content);

      const response = await graphClient
        .api(`/me/onenote/sections/${validatedSectionId}/pages`)
        .header('Content-Type', 'application/xhtml+xml')
        .post(pageHtml);

      return {
        content: [
          {
            type: 'text',
            text: `✅ **Page Created Successfully!**
**Title:** ${response.title}
**Page ID:** ${response.id}
**In Section:** ${targetSectionName}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`,
          },
        ],
      };
    }, 'Error creating page in section')
  );

  server.tool(
    'createNotebook',
    {
      displayName: z
        .string()
        .min(1, { message: 'Notebook name cannot be empty.' })
        .describe('The name for the new notebook.'),
    },
    createToolHandler(async ({ displayName }) => {
      const graphClient = getGraphClient();
      console.error(`Creating notebook: "${displayName}"`);

      const response = await graphClient.api('/me/onenote/notebooks').post({ displayName });

      // Invalidate notebook list cache
      apiCache.invalidate(CacheKeys.notebooks());

      return {
        content: [
          {
            type: 'text',
            text: `✅ **Notebook Created Successfully!**
**Name:** ${response.displayName}
**Notebook ID:** ${response.id}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`,
          },
        ],
      };
    }, 'Error creating notebook')
  );

  server.tool(
    'createSection',
    {
      notebookId: z
        .string()
        .min(1, { message: 'Notebook ID cannot be empty.' })
        .describe('The ID of the notebook to create the section in.'),
      displayName: z
        .string()
        .min(1, { message: 'Section name cannot be empty.' })
        .describe('The name for the new section.'),
    },
    createToolHandler(async ({ notebookId, displayName }) => {
      const graphClient = getGraphClient();
      const { id: validatedNotebookId } = await validateAndFetchResource(
        notebookId,
        'notebook',
        `/me/onenote/notebooks/${notebookId}`
      );

      console.error(`Creating section "${displayName}" in notebook: ${validatedNotebookId}`);

      const response = await graphClient
        .api(`/me/onenote/notebooks/${validatedNotebookId}/sections`)
        .post({ displayName });

      // Invalidate sections cache for this notebook and general sections list
      apiCache.invalidate(CacheKeys.sections(validatedNotebookId));
      apiCache.invalidate(CacheKeys.sections());

      return {
        content: [
          {
            type: 'text',
            text: `✅ **Section Created Successfully!**
**Name:** ${response.displayName}
**Section ID:** ${response.id}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`,
          },
        ],
      };
    }, 'Error creating section')
  );

  server.tool(
    'createSectionGroup',
    {
      notebookId: z
        .string()
        .min(1, { message: 'Notebook ID cannot be empty.' })
        .describe('The ID of the notebook to create the section group in.'),
      displayName: z
        .string()
        .min(1, { message: 'Section group name cannot be empty.' })
        .describe('The name for the new section group.'),
    },
    createToolHandler(async ({ notebookId, displayName }) => {
      const graphClient = getGraphClient();
      const { id: validatedNotebookId } = await validateAndFetchResource(
        notebookId,
        'notebook',
        `/me/onenote/notebooks/${notebookId}`
      );

      console.error(`Creating section group "${displayName}" in notebook: ${validatedNotebookId}`);

      const response = await graphClient
        .api(`/me/onenote/notebooks/${validatedNotebookId}/sectionGroups`)
        .post({ displayName });

      return {
        content: [
          {
            type: 'text',
            text: `✅ **Section Group Created Successfully!**
**Name:** ${response.displayName}
**Section Group ID:** ${response.id}
**Created:** ${new Date(response.createdDateTime).toLocaleString()}`,
          },
        ],
      };
    }, 'Error creating section group')
  );
}
