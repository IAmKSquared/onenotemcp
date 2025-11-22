import { z } from 'zod';
import { cachedApiCall, CacheKeys } from '../api/cache.mjs';
import { createToolHandler } from '../api/retry.mjs';
import { formatPageInfo, formatItemList } from '../utils/html.mjs';
import { validateAndFetchResource, fetchPageContentAdvanced } from '../utils/validation.mjs';
import {
  escapeODataString,
  validateId,
  extractTextSummary,
  extractReadableText,
} from '../utils/common.mjs';
import { DISPLAY_LIMITS } from '../config/constants.mjs';

/**
 * Registers read-related tools with the MCP server.
 * @param {McpServer} server - The MCP server instance.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 */
export function registerReadTools(server, session) {
  server.tool(
    'listNotebooks',
    {},
    createToolHandler(
      session,
      async () => {
        const graphClient = session.getGraphClient();
        const response = await cachedApiCall(
          CacheKeys.notebooks(),
          async () => await graphClient.api('/me/onenote/notebooks').get()
        );

        if (response.value && response.value.length > 0) {
          const notebookList = response.value.map((nb, i) => formatPageInfo(nb, i)).join('\n\n');
          return {
            content: [
              {
                type: 'text',
                text: `📚 **Your OneNote Notebooks** (${response.value.length} found):\n\n${notebookList}`,
              },
            ],
          };
        } else {
          return { content: [{ type: 'text', text: '📚 No OneNote notebooks found.' }] };
        }
      },
      'Failed to list notebooks'
    )
  );

  server.tool(
    'listSections',
    {
      notebookId: z.string().describe('The ID of the parent notebook.').optional(),
      sectionGroupId: z.string().describe('The ID of the parent section group.').optional(),
    },
    createToolHandler(
      session,
      async ({ notebookId, sectionGroupId }) => {
        const graphClient = session.getGraphClient();
        let endpoint = '/me/onenote/sections';

        if (notebookId) {
          endpoint = `/me/onenote/notebooks/${notebookId}/sections`;
        } else if (sectionGroupId) {
          endpoint = `/me/onenote/sectionGroups/${sectionGroupId}/sections`;
        }

        const response = await cachedApiCall(
          CacheKeys.sections(notebookId, sectionGroupId),
          async () => await graphClient.api(endpoint).get()
        );

        if (response.value && response.value.length > 0) {
          const list = response.value.map((item, i) => formatPageInfo(item, i)).join('\n\n');
          return {
            content: [
              {
                type: 'text',
                text: `📂 **Sections** (${response.value.length} found):\n\n${list}`,
              },
            ],
          };
        } else {
          return { content: [{ type: 'text', text: '📂 No sections found.' }] };
        }
      },
      'Failed to list sections'
    )
  );

  server.tool(
    'listSectionGroups',
    {
      notebookId: z.string().describe('The ID of the parent notebook.').optional(),
      sectionGroupId: z.string().describe('The ID of the parent section group.').optional(),
    },
    createToolHandler(
      session,
      async ({ notebookId, sectionGroupId }) => {
        const graphClient = session.getGraphClient();
        let endpoint = '/me/onenote/sectionGroups';
        if (notebookId) {
          endpoint = `/me/onenote/notebooks/${notebookId}/sectionGroups`;
        } else if (sectionGroupId) {
          endpoint = `/me/onenote/sectionGroups/${sectionGroupId}/sectionGroups`;
        }

        const response = await graphClient.api(endpoint).get();
        if (response.value && response.value.length > 0) {
          const list = response.value.map((item, i) => formatPageInfo(item, i)).join('\n\n');
          return {
            content: [
              {
                type: 'text',
                text: `📁 **Section Groups** (${response.value.length} found):\n\n${list}`,
              },
            ],
          };
        } else {
          return { content: [{ type: 'text', text: '📁 No section groups found.' }] };
        }
      },
      'Failed to list section groups'
    )
  );

  server.tool(
    'searchSections',
    {
      query: z.string().describe('The search term for section names.'),
    },
    createToolHandler(
      session,
      async ({ query }) => {
        const graphClient = session.getGraphClient();
        const escapedQuery = escapeODataString(query).toLowerCase();

        const response = await graphClient
          .api('/me/onenote/sections')
          .filter(`contains(tolower(displayName), '${escapedQuery}')`)
          .select('id,displayName,parentNotebook,parentSectionGroup')
          .top(50)
          .get();

        const sections = response.value || [];

        if (sections.length > 0) {
          const { list, more, limitWarning } = formatItemList(sections, 'results');
          return {
            content: [
              {
                type: 'text',
                text: `🔍 **Section Search Results** for "${query}" (${sections.length} found):\n\n${list}${more}${limitWarning}`,
              },
            ],
          };
        } else {
          return { content: [{ type: 'text', text: `🔍 No sections found matching "${query}".` }] };
        }
      },
      'Failed to search sections'
    )
  );

  server.tool(
    'listPagesInSection',
    {
      sectionId: z.string().describe('The ID of the section to list pages from.'),
    },
    createToolHandler(
      session,
      async ({ sectionId }) => {
        const graphClient = session.getGraphClient();
        const { id: validatedSectionId, resource: sectionInfo } = await validateAndFetchResource(
          session,
          sectionId,
          'section',
          `/me/onenote/sections/${sectionId}`,
          'listSections or searchSections'
        );
        const sectionName = sectionInfo.displayName;

        const response = await graphClient
          .api(`/me/onenote/sections/${validatedSectionId}/pages`)
          .select('id,title,lastModifiedDateTime')
          .top(50)
          .get();

        const pages = response.value || [];

        if (pages.length > 0) {
          const { list, more, limitWarning } = formatItemList(pages, 'pages');
          return {
            content: [
              {
                type: 'text',
                text: `📄 **Pages in Section "${sectionName}"** (${pages.length} found):\n\n${list}${more}${limitWarning}`,
              },
            ],
          };
        } else {
          return {
            content: [{ type: 'text', text: `📄 No pages found in section "${sectionName}".` }],
          };
        }
      },
      'Failed to list pages in section'
    )
  );

  server.tool(
    'searchPages',
    {
      query: z.string().describe('The search term for page titles.').optional(),
      modifiedAfter: z
        .string()
        .describe(
          'Filter pages modified after this date (ISO 8601 format: YYYY-MM-DD or YYYY-MM-DDTHH:MM:SSZ).'
        )
        .optional(),
      modifiedBefore: z
        .string()
        .describe(
          'Filter pages modified before this date (ISO 8601 format: YYYY-MM-DD or YYYY-MM-DDTHH:MM:SSZ).'
        )
        .optional(),
      notebookId: z
        .string()
        .describe('Filter pages within a specific notebook (must provide notebook ID).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ query, modifiedAfter, modifiedBefore, notebookId }) => {
        const graphClient = session.getGraphClient();
        const filterConditions = [];

        if (query) {
          const escapedQuery = escapeODataString(query).toLowerCase();
          filterConditions.push(`contains(tolower(title), '${escapedQuery}')`);
        }

        if (modifiedAfter) {
          filterConditions.push(`lastModifiedDateTime ge ${modifiedAfter}`);
        }

        if (modifiedBefore) {
          filterConditions.push(`lastModifiedDateTime le ${modifiedBefore}`);
        }

        if (notebookId) {
          const validatedNotebookId = validateId(notebookId, 'notebook');
          const sectionsResponse = await graphClient
            .api(`/me/onenote/notebooks/${validatedNotebookId}/sections`)
            .get();
          const sections = sectionsResponse.value || [];

          if (sections.length === 0) {
            return {
              content: [
                { type: 'text', text: `📄 No sections found in notebook. Cannot search pages.` },
              ],
            };
          }

          const sectionPagePromises = sections.map(async (section) => {
            let sectionRequest = graphClient
              .api(`/me/onenote/sections/${section.id}/pages`)
              .select('id,title,lastModifiedDateTime')
              .top(50);

            if (filterConditions.length > 0) {
              sectionRequest = sectionRequest.filter(filterConditions.join(' and '));
            }

            const sectionPages = await sectionRequest.get();
            return sectionPages.value || [];
          });

          const sectionResults = await Promise.all(sectionPagePromises);
          const allPages = sectionResults.flat();
          const pages = allPages.slice(0, 50);

          if (pages.length > 0) {
            const { list, more, limitWarning } = formatItemList(pages, 'pages');
            return {
              content: [
                {
                  type: 'text',
                  text: `🔍 **Search Results** in notebook (${pages.length} found):\n\n${list}${more}${limitWarning}`,
                },
              ],
            };
          } else {
            return {
              content: [{ type: 'text', text: `🔍 No pages found in notebook matching criteria.` }],
            };
          }
        }

        let request = graphClient
          .api('/me/onenote/pages')
          .select('id,title,lastModifiedDateTime')
          .top(50);

        if (filterConditions.length > 0) {
          request = request.filter(filterConditions.join(' and '));
        }

        const apiResponse = await request.get();
        const pages = apiResponse.value || [];

        if (pages.length > 0) {
          const { list, more, limitWarning } = formatItemList(pages, 'pages');

          let searchDesc = '';
          if (query) searchDesc += `"${query}"`;
          if (modifiedAfter) searchDesc += ` modified after ${modifiedAfter}`;
          if (modifiedBefore) searchDesc += ` modified before ${modifiedBefore}`;

          return {
            content: [
              {
                type: 'text',
                text: `🔍 **Search Results** ${searchDesc || ''} (${pages.length} found):\n\n${list}${more}${limitWarning}`,
              },
            ],
          };
        } else {
          return { content: [{ type: 'text', text: `🔍 No pages found matching criteria.` }] };
        }
      },
      'Failed to search pages'
    )
  );

  server.tool(
    'getPageContent',
    {
      pageId: z.string().describe('The ID of the page to retrieve content from.'),
      format: z
        .enum(['text', 'html', 'summary'])
        .default('text')
        .describe('Format of the content: text (readable), html (raw), or summary (brief).')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ pageId, format }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');

        const pageInfo = await graphClient.api(`/me/onenote/pages/${validatedPageId}`).get();
        const htmlContent = await fetchPageContentAdvanced(session, validatedPageId, 'httpDirect');
        let resultText = '';

        if (format === 'html') {
          resultText = `📄 **${pageInfo.title}** (HTML Format)\n\n${htmlContent}`;
        } else if (format === 'summary') {
          const summary = extractTextSummary(htmlContent);
          resultText = `📄 **${pageInfo.title}** (Summary)\n\n${summary}`;
        } else {
          const textContent = extractReadableText(htmlContent);
          resultText = `📄 **${pageInfo.title}**\n📅 Modified: ${new Date(pageInfo.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
        }
        return { content: [{ type: 'text', text: resultText }] };
      },
      'Failed to get page content'
    )
  );

  server.tool(
    'getPageByTitle',
    {
      title: z.string().describe('The title (or partial title) of the page to find.'),
      format: z
        .enum(['text', 'html', 'summary'])
        .default('text')
        .describe('Format of the content: text, html, or summary.')
        .optional(),
    },
    createToolHandler(
      session,
      async ({ title, format }) => {
        const graphClient = session.getGraphClient();
        const escapedTitle = escapeODataString(title).toLowerCase();
        const pagesResponse = await graphClient
          .api('/me/onenote/pages')
          .filter(`contains(tolower(title), '${escapedTitle}')`)
          .select('id,title,lastModifiedDateTime')
          .top(50)
          .get();

        const matchingPages = pagesResponse.value || [];

        if (matchingPages.length === 0) {
          const recentPages = await graphClient
            .api('/me/onenote/pages')
            .select('title')
            .top(10)
            .orderby('lastModifiedDateTime desc')
            .get();
          const availablePages = (recentPages.value || []).map((p) => `- ${p.title}`).join('\n');
          return {
            isError: true,
            content: [
              {
                type: 'text',
                text: `❌ No page found with title containing "${title}".\n\nRecent pages (up to 10):\n${availablePages || 'None'}`,
              },
            ],
          };
        }

        const matchingPage = matchingPages[0];
        const htmlContent = await fetchPageContentAdvanced(session, matchingPage.id, 'httpDirect');
        let resultText = '';
        if (format === 'html') {
          resultText = `📄 **${matchingPage.title}** (HTML Format)\n\n${htmlContent}`;
        } else if (format === 'summary') {
          const summary = extractTextSummary(htmlContent);
          resultText = `📄 **${matchingPage.title}** (Summary)\n\n${summary}`;
        } else {
          const textContent = extractReadableText(htmlContent);
          resultText = `📄 **${matchingPage.title}**\n📅 Modified: ${new Date(matchingPage.lastModifiedDateTime).toLocaleString()}\n\n${textContent}`;
        }

        if (matchingPages.length > 1) {
          resultText += `\n\n📌 Note: ${matchingPages.length} pages matched "${title}". Showing the first match.`;
          if (matchingPages.length === DISPLAY_LIMITS.API_RESULT_LIMIT) {
            resultText += `\n⚠️ Reached the ${DISPLAY_LIMITS.API_RESULT_LIMIT}-result limit. There may be additional matches not shown.`;
          }
        }

        return { content: [{ type: 'text', text: resultText }] };
      },
      'Failed to get page by title'
    )
  );
}
