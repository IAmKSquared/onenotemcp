import { z } from 'zod';
import { createToolHandler } from '../api/retry.mjs';
import { validateId } from '../utils/common.mjs';
import { logger } from '../utils/logger.mjs';

/**
 * Registers delete-related tools with the MCP server.
 *
 * SAFETY: All delete operations require explicit confirmation via the
 * `confirmDelete` parameter to prevent accidental data loss.
 *
 * NOTE: Only page deletion is supported by Microsoft Graph API.
 * Sections, section groups, and notebooks cannot be deleted via the API.
 * See: https://learn.microsoft.com/en-us/graph/api/page-delete
 * @param {McpServer} server - The MCP server instance.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 */
export function registerDeleteTools(server, session) {
  server.tool(
    'deletePage',
    {
      pageId: z.string().describe('The ID of the page to delete.'),
      confirmDelete: z
        .boolean()
        .describe('Must be true to confirm deletion. This action cannot be undone.'),
    },
    createToolHandler(
      session,
      async ({ pageId, confirmDelete }) => {
        if (!confirmDelete) {
          return {
            content: [
              {
                type: 'text',
                text: '⚠️ **Deletion cancelled.** Set `confirmDelete` to `true` to proceed with deletion.',
              },
            ],
          };
        }

        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');

        // Fetch page info first for confirmation message
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        logger.info(`Deleting page: "${pageInfo.title}" (ID: ${validatedPageId})`);

        await graphClient.api(`/me/onenote/pages/${validatedPageId}`).delete();

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Page Deleted**\nTitle: "${pageInfo.title}"\nDeleted: ${new Date().toLocaleString()}`,
            },
          ],
        };
      },
      'Failed to delete page'
    )
  );
}
