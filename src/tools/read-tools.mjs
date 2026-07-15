import { z } from 'zod';
import { cachedApiCall, CacheKeys } from '../api/cache.mjs';
import { createToolHandler } from '../api/retry.mjs';
import { formatItemInfo, formatItemList, formatMetadata } from '../utils/html.mjs';
import { validateAndFetchResource, fetchPageContentAdvanced } from '../utils/validation.mjs';
import {
  escapeODataString,
  validateId,
  validateIsoDate,
  extractTextSummary,
  extractReadableText,
  collectSettledResults,
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
          async () =>
            await graphClient
              .api('/me/onenote/notebooks')
              .select('id,displayName,lastModifiedDateTime,createdDateTime,lastModifiedBy,links')
              .get()
        );

        if (response.value && response.value.length > 0) {
          const notebookList = response.value.map((nb, i) => formatItemInfo(nb, i)).join('\n\n');
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
        let cacheKey;

        if (notebookId) {
          const validatedId = validateId(notebookId, 'notebook');
          endpoint = `/me/onenote/notebooks/${validatedId}/sections`;
          cacheKey = CacheKeys.sections(validatedId, undefined);
        } else if (sectionGroupId) {
          const validatedId = validateId(sectionGroupId, 'sectionGroup');
          endpoint = `/me/onenote/sectionGroups/${validatedId}/sections`;
          cacheKey = CacheKeys.sections(undefined, validatedId);
        } else {
          cacheKey = CacheKeys.sections(undefined, undefined);
        }

        const response = await cachedApiCall(
          cacheKey,
          async () =>
            await graphClient
              .api(endpoint)
              .select(
                'id,displayName,lastModifiedDateTime,createdDateTime,lastModifiedBy,links,parentNotebook,parentSectionGroup'
              )
              .get()
        );

        if (response.value && response.value.length > 0) {
          const list = response.value.map((item, i) => formatItemInfo(item, i)).join('\n\n');
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
          const validatedNotebookId = validateId(notebookId, 'notebook');
          endpoint = `/me/onenote/notebooks/${validatedNotebookId}/sectionGroups`;
        } else if (sectionGroupId) {
          const validatedSectionGroupId = validateId(sectionGroupId, 'sectionGroup');
          endpoint = `/me/onenote/sectionGroups/${validatedSectionGroupId}/sectionGroups`;
        }

        const response = await graphClient
          .api(endpoint)
          .select(
            'id,displayName,lastModifiedDateTime,createdDateTime,lastModifiedBy,sectionGroupsUrl,sectionsUrl'
          )
          .get();
        if (response.value && response.value.length > 0) {
          const list = response.value.map((item, i) => formatItemInfo(item, i)).join('\n\n');
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
          .select(
            'id,displayName,lastModifiedDateTime,createdDateTime,lastModifiedBy,links,parentNotebook,parentSectionGroup'
          )
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
          .select('id,title,lastModifiedDateTime,createdDateTime,links')
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
          validateIsoDate(modifiedAfter, 'modifiedAfter');
          filterConditions.push(`lastModifiedDateTime ge ${modifiedAfter}`);
        }

        if (modifiedBefore) {
          validateIsoDate(modifiedBefore, 'modifiedBefore');
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
              .select('id,title,lastModifiedDateTime,createdDateTime,links')
              .top(50);

            if (filterConditions.length > 0) {
              sectionRequest = sectionRequest.filter(filterConditions.join(' and '));
            }

            const sectionPages = await sectionRequest.get();
            return sectionPages.value || [];
          });

          const allPages = await collectSettledResults(sectionPagePromises, 'section(s)');
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
          .select('id,title,lastModifiedDateTime,createdDateTime,links')
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
    'getRecentPages',
    {
      notebookId: z.string().describe('The notebook ID to get recent pages from.'),
      limit: z
        .number()
        .int()
        .min(1)
        .max(50)
        .default(10)
        .describe('Number of pages to return (default: 10, max: 50).'),
    },
    createToolHandler(
      session,
      async ({ notebookId, limit = 10 }) => {
        const graphClient = session.getGraphClient();
        const validatedNotebookId = validateId(notebookId, 'notebook');

        // Get notebook info for display name
        const notebookResponse = await graphClient
          .api(`/me/onenote/notebooks/${validatedNotebookId}`)
          .select('displayName')
          .get();
        const notebookName = notebookResponse.displayName;

        const sectionsResponse = await graphClient
          .api(`/me/onenote/notebooks/${validatedNotebookId}/sections`)
          .get();
        const sections = sectionsResponse.value || [];

        if (sections.length === 0) {
          return {
            content: [
              { type: 'text', text: `📄 No sections found in notebook "${notebookName}".` },
            ],
          };
        }

        // Fetch recent pages from each section in parallel (orderby ensures we get the most recent)
        const sectionPagePromises = sections.map(async (section) => {
          const sectionPages = await graphClient
            .api(`/me/onenote/sections/${section.id}/pages`)
            .select('id,title,lastModifiedDateTime,createdDateTime,links')
            .orderby('lastModifiedDateTime desc')
            .top(limit)
            .get();
          return sectionPages.value || [];
        });

        const allPages = await collectSettledResults(sectionPagePromises, 'section(s)');

        // Sort merged results by lastModifiedDateTime descending
        allPages.sort(
          (a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime)
        );

        const pages = allPages.slice(0, limit);

        if (pages.length > 0) {
          const { list, more, limitWarning } = formatItemList(pages, 'pages', limit, 50);
          return {
            content: [
              {
                type: 'text',
                text: `📄 **Recent Pages in "${notebookName}"** (${pages.length} found):\n\n${list}${more}${limitWarning}`,
              },
            ],
          };
        } else {
          return {
            content: [{ type: 'text', text: `📄 No pages found in notebook "${notebookName}".` }],
          };
        }
      },
      'Failed to get recent pages'
    )
  );

  server.tool(
    'getPageContent',
    {
      pageId: z.string().describe('The ID of the page to retrieve content from.'),
      format: z
        .enum(['text', 'html', 'summary'])
        .default('text')
        .describe('Format of the content: text (readable), html (raw), or summary (brief).'),
    },
    createToolHandler(
      session,
      async ({ pageId, format }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');

        const pageInfo = await graphClient
          .api(`/me/onenote/pages/${validatedPageId}`)
          .select('id,title,lastModifiedDateTime,createdDateTime,links')
          .get();
        const htmlContent = await fetchPageContentAdvanced(session, validatedPageId, 'httpDirect');
        let resultText = '';
        const metadata = formatMetadata(pageInfo);

        const metadataLine = metadata ? `\n${metadata}` : '';
        if (format === 'html') {
          resultText = `📄 **${pageInfo.title}** (HTML Format)${metadataLine}\n\n${htmlContent}`;
        } else if (format === 'summary') {
          const summary = extractTextSummary(htmlContent);
          resultText = `📄 **${pageInfo.title}** (Summary)${metadataLine}\n\n${summary}`;
        } else {
          const textContent = extractReadableText(htmlContent);
          resultText = `📄 **${pageInfo.title}**${metadataLine}\n\n${textContent}`;
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
        .describe('Format of the content: text, html, or summary.'),
    },
    createToolHandler(
      session,
      async ({ title, format }) => {
        const graphClient = session.getGraphClient();
        const escapedTitle = escapeODataString(title).toLowerCase();
        const pagesResponse = await graphClient
          .api('/me/onenote/pages')
          .filter(`contains(tolower(title), '${escapedTitle}')`)
          .select('id,title,lastModifiedDateTime,createdDateTime,links')
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
        const metadata = formatMetadata(matchingPage);
        const metadataLine = metadata ? `\n${metadata}` : '';
        let resultText = '';
        if (format === 'html') {
          resultText = `📄 **${matchingPage.title}** (HTML Format)${metadataLine}\n\n${htmlContent}`;
        } else if (format === 'summary') {
          const summary = extractTextSummary(htmlContent);
          resultText = `📄 **${matchingPage.title}** (Summary)${metadataLine}\n\n${summary}`;
        } else {
          const textContent = extractReadableText(htmlContent);
          resultText = `📄 **${matchingPage.title}**${metadataLine}\n\n${textContent}`;
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

  server.tool(
    'getPageLink',
    {
      pageId: z.string().describe('The ID of the page to get links for.'),
    },
    createToolHandler(
      session,
      async ({ pageId }) => {
        const graphClient = session.getGraphClient();
        const validatedPageId = validateId(pageId, 'page');

        const response = await graphClient
          .api(`/me/onenote/pages/${validatedPageId}`)
          .select('id,title,links')
          .get();

        // Graph API returns links as objects with href property
        const webUrl =
          response.links?.oneNoteWebUrl?.href || response.links?.oneNoteWebUrl || 'Not available';
        const clientUrl =
          response.links?.oneNoteClientUrl?.href ||
          response.links?.oneNoteClientUrl ||
          'Not available';

        return {
          content: [
            {
              type: 'text',
              text: `📄 **${response.title}**\n\n🌐 **Web URL (open in browser):**\n${webUrl}\n\n💻 **Desktop App URL:**\n${clientUrl}`,
            },
          ],
        };
      },
      'Failed to get page link'
    )
  );
}
