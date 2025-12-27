import { z } from 'zod';
import { createToolHandler } from '../api/retry.mjs';
import { patchPageContent } from '../utils/validation.mjs';
import { textToHtml, validateId, validateCsvData } from '../utils/common.mjs';
import { fetchPageContentAdvanced } from '../utils/validation.mjs';
import { HTTP_STATUS } from '../config/constants.mjs';
import { logger } from '../utils/logger.mjs';

/**
 * Registers write/update-related tools with the MCP server.
 *
 * NOTE ON CACHE INVALIDATION:
 * These operations modify page content but do NOT invalidate any caches because
 * pages are intentionally not cached. Pages have dynamic content with frequent
 * mutations, and caching them would risk serving stale data. See src/api/cache.mjs
 * for the complete caching strategy documentation.
 * @param {McpServer} server - The MCP server instance.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 */
export function registerWriteTools(server, session) {
  server.tool(
    'updatePageContent',
    {
      pageId: z.string().describe('The ID of the page to update.'),
      content: z.string().describe('New page content (HTML or markdown-style text).'),
      preserveTitle: z
        .boolean()
        .default(true)
        .describe('Keep the original title (default: true).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, content: newContent, preserveTitle }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');

        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        logger.info(`Updating content for page: "${pageInfo.title}" (ID: ${validatedPageId})`);

        const htmlContentForUpdate = textToHtml(newContent);
        const finalHtml = `
      <div>
        ${preserveTitle ? `<h1>${pageInfo.title}</h1>` : ''}
        ${htmlContentForUpdate}
        <hr>
        <p><em>Updated via OneNote MCP on ${new Date().toLocaleString()}</em></p>
      </div>
    `;

        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'body', action: 'replace', content: finalHtml }],
          'Update failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Page Content Updated!**\nPage: ${pageInfo.title}\nUpdated: ${new Date().toLocaleString()}\nContent Length: ${newContent.length} chars.`,
            },
          ],
        };
      },
      'Failed to update page content'
    )
  );

  server.tool(
    'appendToPage',
    {
      pageId: z.string().describe('The ID of the page to append content to.'),
      content: z.string().describe('Content to append (HTML or markdown-style).'),
      addTimestamp: z
        .boolean()
        .default(true)
        .describe('Add a timestamp (default: true).')
        .optional(),
      addSeparator: z
        .boolean()
        .default(true)
        .describe('Add a visual separator (default: true).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, content: newContent, addTimestamp, addSeparator }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        logger.info(`Appending content to page: "${pageInfo.title}" (ID: ${validatedPageId})`);

        const htmlContentToAppend = textToHtml(newContent);
        let appendHtml = '';
        if (addSeparator) appendHtml += '<hr>';
        if (addTimestamp) appendHtml += `<p><em>Added on ${new Date().toLocaleString()}</em></p>`;
        appendHtml += htmlContentToAppend;

        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'body', action: 'append', content: appendHtml }],
          'Append failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Content Appended!**\nPage: ${pageInfo.title}\nAppended: ${new Date().toLocaleString()}\nLength: ${newContent.length} chars.`,
            },
          ],
        };
      },
      'Failed to append content'
    )
  );

  server.tool(
    'updatePageTitle',
    {
      pageId: z.string().describe('The ID of the page whose title is to be updated.'),
      newTitle: z.string().describe('The new title for the page.'),
    },
    createToolHandler(
      session,
      async ({ pageId, newTitle }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        const oldTitle = pageInfo.title;
        logger.info(
          `Updating page title from "${oldTitle}" to "${newTitle}" for page ID "${validatedPageId}"`
        );

        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'title', action: 'replace', content: newTitle }],
          'Title update failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Page Title Updated!**\nOld Title: ${oldTitle}\nNew Title: ${newTitle}\nUpdated: ${new Date().toLocaleString()}`,
            },
          ],
        };
      },
      'Failed to update page title'
    )
  );

  server.tool(
    'replaceTextInPage',
    {
      pageId: z.string().describe('The ID of the page to modify.'),
      findText: z.string().describe('The text to find and replace.'),
      replaceText: z.string().describe('The text to replace with.'),
      caseSensitive: z
        .boolean()
        .default(false)
        .describe('Case-sensitive search (default: false).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, findText, replaceText, caseSensitive }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        const htmlContent = await fetchPageContentAdvanced(session, validatedPageId, 'httpDirect');
        logger.info(`Replacing text in page: "${pageInfo.title}" (ID: ${validatedPageId})`);

        const flags = caseSensitive ? 'g' : 'gi';
        const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
        const matches = (htmlContent.match(regex) || []).length;

        if (matches === 0) {
          return {
            content: [
              {
                type: 'text',
                text: `ℹ️ **No matches found** for "${findText}" in page: ${pageInfo.title}.`,
              },
            ],
          };
        }

        const updatedContent = htmlContent.replace(regex, replaceText);
        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'body', action: 'replace', content: `<div>${updatedContent}</div>` }],
          'Replace failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Text Replaced!**\nPage: ${pageInfo.title}\nFound: "${findText}" (${matches} occurrences)\nReplaced with: "${replaceText}".`,
            },
          ],
        };
      },
      'Failed to replace text'
    )
  );

  server.tool(
    'addNoteToPage',
    {
      pageId: z.string().describe('The ID of the page to add a note to.'),
      note: z.string().describe('The note/comment content.'),
      noteType: z
        .enum(['note', 'todo', 'important', 'question'])
        .default('note')
        .describe('Type of note (note, todo, important, question).')
        .optional(),
      position: z
        .enum(['top', 'bottom'])
        .default('bottom')
        .describe('Position to add the note (top or bottom).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, note, noteType, position }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        logger.info(
          `Adding ${noteType} to page: "${pageInfo.title}" (ID: ${validatedPageId}) at ${position}`
        );

        const icons = { note: '📝', todo: '✅', important: '🚨', question: '❓' };
        const colors = {
          note: '#e3f2fd',
          todo: '#e8f5e8',
          important: '#ffebee',
          question: '#fff3e0',
        };
        const noteHtml = `
      <div style="border-left: 4px solid #2196f3; background-color: ${colors[noteType]}; padding: 10px; margin: 10px 0;">
        <p><strong>${icons[noteType]} ${noteType.charAt(0).toUpperCase() + noteType.slice(1)}</strong> - <em>${new Date().toLocaleString()}</em></p>
        <p>${textToHtml(note)}</p>
      </div>`;

        const action = position === 'top' ? 'prepend' : 'append';
        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'body', action: action, content: noteHtml }],
          'Add note failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **${noteType.charAt(0).toUpperCase() + noteType.slice(1)} Added!**\nPage: ${pageInfo.title}\nPosition: ${position}.`,
            },
          ],
        };
      },
      'Failed to add note'
    )
  );

  server.tool(
    'addTableToPage',
    {
      pageId: z.string().describe('The ID of the page to add a table to.'),
      tableData: z
        .string()
        .describe(
          'Table data in CSV format (header row, then data rows). Note: Simple CSV only - quoted fields with commas are not supported.'
        ),
      title: z.string().describe('Optional title for the table.').optional(),
      position: z
        .enum(['top', 'bottom'])
        .default('bottom')
        .describe('Position to add the table (top or bottom).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, tableData, title, position }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        logger.info(
          `Adding table to page: "${pageInfo.title}" (ID: ${validatedPageId}) at ${position}`
        );

        const rows = validateCsvData(tableData);

        const headerRow = rows[0];
        const dataRows = rows.slice(1);
        let tableHtml = title ? `<h3>📊 ${textToHtml(title)}</h3>` : '';
        tableHtml += `<table style="border-collapse: collapse; width: 100%; margin: 10px 0;"><thead><tr style="background-color: #f5f5f5;">${headerRow.map((cell) => `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${textToHtml(cell)}</th>`).join('')}</tr></thead><tbody>${dataRows.map((row) => `<tr>${row.map((cell) => `<td style="border: 1px solid #ddd; padding: 8px;">${textToHtml(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;

        const action = position === 'top' ? 'prepend' : 'append';
        await patchPageContent(
          session,
          validatedPageId,
          [{ target: 'body', action: action, content: tableHtml }],
          'Add table failed'
        );

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Table Added!**\nPage: ${pageInfo.title}\nTitle: ${title || 'Untitled'}\nPosition: ${position}.`,
            },
          ],
        };
      },
      'Failed to add table'
    )
  );

  server.tool(
    'copyPage',
    {
      pageId: z.string().describe('The ID of the page to copy.'),
      targetSectionId: z.string().describe('The ID of the section to copy the page to.'),
    },
    createToolHandler(
      session,
      async ({ pageId, targetSectionId }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');
        const validatedSectionId = validateId(targetSectionId, 'section');

        logger.info(`Copying page ${validatedPageId} to section ${validatedSectionId}`);

        let pageInfo, targetSectionName;
        try {
          const [pageResult, sectionResult] = await Promise.all([
            graphClient.api(`/me/onenote/pages/${validatedPageId}`).get(),
            graphClient.api(`/me/onenote/sections/${validatedSectionId}`).get(),
          ]);
          pageInfo = pageResult;
          targetSectionName = sectionResult.displayName;
        } catch (error) {
          if (error.statusCode === HTTP_STATUS.NOT_FOUND) {
            if (error.message && error.message.includes('section')) {
              throw new Error(
                `Target section with ID "${validatedSectionId}" not found. Use listSections or searchSections to find valid section IDs.`
              );
            } else {
              throw new Error(`Page with ID "${validatedPageId}" not found.`);
            }
          }
          throw error;
        }

        const _copyResponse = await graphClient
          .api(`/me/onenote/pages/${validatedPageId}/copyToSection`)
          .post({ id: validatedSectionId });

        return {
          content: [
            {
              type: 'text',
              text: `✅ **Page Copy Initiated!**
**Original Page:** ${pageInfo.title}
**Target Section:** ${targetSectionName}

⚠️ Note: Copy is an asynchronous operation. The page will appear in the target section shortly.`,
            },
          ],
        };
      },
      'Error copying page'
    )
  );
}
