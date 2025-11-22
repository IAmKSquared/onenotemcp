import fetch from 'node-fetch';
import { validateId } from '../../utils.mjs';

/**
 * Fetches the content of a OneNote page.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 * @param {string} pageId - The ID of the page.
 * @param {'httpDirect' | 'direct'} [method] - The method to use for fetching.
 * @returns {Promise<string>} The HTML content of the page.
 */
export async function fetchPageContentAdvanced(session, pageId, method = 'httpDirect') {
  const graphClient = await session.ensureGraphClient();
  const accessToken = session.getAccessToken();

  if (method === 'httpDirect') {
    const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${encodeURIComponent(pageId)}/content`;
    const response = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!response.ok)
      throw new Error(
        `HTTP error fetching page content! Status: ${response.status} ${response.statusText}`
      );
    return await response.text();
  } else {
    // 'direct'
    return await graphClient.api(`/me/onenote/pages/${pageId}/content`).get();
  }
}

/**
 * Sends a PATCH request to update OneNote page content.
 * Handles the common pattern of PATCH operations for page modifications.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 * @param {string} pageId - The ID of the page to update.
 * @param {Array} operations - Array of operations (e.g., [{target: 'body', action: 'append', content: '...'}]).
 * @param {string} [errorPrefix] - Custom error message prefix. Defaults to 'PATCH operation failed'.
 * @returns {Promise<object>} The fetch response object.
 * @throws {Error} If the PATCH request fails.
 */
export async function patchPageContent(
  session,
  pageId,
  operations,
  errorPrefix = 'PATCH operation failed'
) {
  const accessToken = session.getAccessToken();
  const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${encodeURIComponent(pageId)}/content`;
  const response = await fetch(url, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(operations),
  });

  if (!response.ok) {
    throw new Error(`${errorPrefix}: ${response.status} ${response.statusText}`);
  }

  return response;
}

/**
 * Validates and fetches a OneNote resource, with helpful error messages for 404s.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 * @param {string} id - The ID of the resource to validate and fetch.
 * @param {string} resourceType - The type of resource (e.g., 'section', 'notebook', 'sectionGroup').
 * @param {string} endpoint - The Graph API endpoint to fetch the resource.
 * @param {string} [listToolSuggestion] - Optional suggestion for which tool to use to find valid IDs.
 * @returns {Promise<object>} Object with {id: validatedId, resource: fetchedResource}.
 * @throws {Error} If validation fails or resource is not found.
 */
export async function validateAndFetchResource(
  session,
  id,
  resourceType,
  endpoint,
  listToolSuggestion
) {
  const graphClient = await session.ensureGraphClient();
  const validatedId = validateId(id, resourceType);

  try {
    const resource = await graphClient.api(endpoint).get();
    return { id: validatedId, resource };
  } catch (error) {
    if (error.statusCode === 404) {
      const suggestion =
        listToolSuggestion ||
        `list${resourceType.charAt(0).toUpperCase() + resourceType.slice(1)}s`;
      throw new Error(
        `${resourceType.charAt(0).toUpperCase() + resourceType.slice(1)} with ID "${validatedId}" not found. Use ${suggestion} to find valid ${resourceType} IDs.`
      );
    }
    throw error;
  }
}
