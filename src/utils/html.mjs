import { textToHtml } from '../../utils.mjs';

/**
 * Formats OneNote page information for display.
 * @param {object} page - The OneNote page object from Graph API.
 * @param {number | null} [index] - Optional index for numbered lists.
 * @returns {string} Formatted page information string.
 */
export function formatPageInfo(page, index = null) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  const name = page.displayName || page.title || 'Untitled';
  return `${prefix}**${name}** (ID: ${page.id})`;
}

/**
 * Formats a list of items with pagination display (first 10 items + "X more" message).
 * @param {Array} items - The array of items to format.
 * @param {string} [itemType] - The type of items (e.g., 'pages', 'sections'). Defaults to 'items'.
 * @param {number} [maxDisplay] - Maximum number of items to display. Defaults to 10.
 * @param {number} [apiLimit] - The API result limit to check for warning. Defaults to 50.
 * @returns {object} Object with {list, more, limitWarning} strings.
 */
export function formatItemList(items, itemType = 'items', maxDisplay = 10, apiLimit = 50) {
  const displayItems = items.slice(0, maxDisplay);
  const list = displayItems.map((item, i) => formatPageInfo(item, i)).join('\n\n');
  const more =
    items.length > maxDisplay ? `\n\n... and ${items.length - maxDisplay} more ${itemType}.` : '';
  const limitWarning =
    items.length === apiLimit
      ? `\n\n⚠️ Note: Reached the ${apiLimit}-result limit. There may be additional matches not shown.`
      : '';
  return { list, more, limitWarning };
}

/**
 * Creates a complete HTML document for a OneNote page.
 * @param {string} title - The page title.
 * @param {string} content - The page content (plain text or markdown).
 * @returns {string} Complete HTML document ready for OneNote API.
 */
export function createPageHtml(title, content) {
  const htmlContent = textToHtml(content);
  return `<!DOCTYPE html>
<html>
<head>
  <title>${textToHtml(title)}</title>
  <meta charset="utf-8">
</head>
<body>
  <h1>${textToHtml(title)}</h1>
  ${htmlContent}
  <hr>
  <p><em>Created via OneNote MCP on ${new Date().toLocaleString()}</em></p>
</body>
</html>`;
}
