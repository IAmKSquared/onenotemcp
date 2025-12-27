import { textToHtml, sanitizeUrl } from '../utils/common.mjs';
import { DISPLAY_LIMITS } from '../config/constants.mjs';

/**
 * Formats an ISO 8601 timestamp to a localized string.
 * @param {string | null | undefined} isoTimestamp - The ISO 8601 timestamp string.
 * @returns {string} Localized date string, or empty string if input is invalid.
 */
export function formatTimestamp(isoTimestamp) {
  if (!isoTimestamp) return '';
  const date = new Date(isoTimestamp);
  if (isNaN(date.getTime())) return '';
  return date.toLocaleString();
}

/**
 * Extracts the display name from a Microsoft Graph identitySet object.
 * @param {object | null | undefined} identitySet - The identitySet object from Graph API.
 * @returns {string} The display name of the user or application, or empty string if unavailable.
 */
export function formatModifiedBy(identitySet) {
  if (!identitySet) return '';
  // Try user first, then application
  const displayName =
    identitySet.user?.displayName ||
    identitySet.user?.email ||
    identitySet.application?.displayName ||
    '';
  return displayName;
}

/**
 * Formats metadata line with timestamps and modified-by info.
 * @param {object} item - The item object with optional timestamp/identity fields.
 * @returns {string} Formatted metadata line, or empty string if no metadata available.
 */
export function formatMetadata(item) {
  const parts = [];

  // Modified timestamp and user
  if (item.lastModifiedDateTime) {
    let modified = `Modified: ${formatTimestamp(item.lastModifiedDateTime)}`;
    const modifiedBy = formatModifiedBy(item.lastModifiedBy);
    if (modifiedBy) {
      modified += ` by ${modifiedBy}`;
    }
    parts.push(modified);
  }

  // Created timestamp
  if (item.createdDateTime) {
    parts.push(`Created: ${formatTimestamp(item.createdDateTime)}`);
  }

  return parts.join(' | ');
}

/**
 * Formats OneNote item information for display (pages, sections, notebooks, section groups).
 * @param {object} item - The OneNote item object from Graph API.
 * @param {number | null} [index] - Optional index for numbered lists.
 * @returns {string} Formatted item information string.
 */
export function formatItemInfo(item, index = null) {
  const prefix = index !== null ? `${index + 1}. ` : '';
  const name = item.displayName || item.title || 'Untitled';

  // Graph API returns links as objects with href property - sanitize for defense in depth
  const webUrl = item.links?.oneNoteWebUrl?.href || item.links?.oneNoteWebUrl;
  const appUrl = item.links?.oneNoteClientUrl?.href || item.links?.oneNoteClientUrl;
  const safeWebUrl = webUrl && typeof webUrl === 'string' ? sanitizeUrl(webUrl) : '';
  const safeAppUrl = appUrl && typeof appUrl === 'string' ? sanitizeUrl(appUrl) : '';
  const webLink = safeWebUrl && safeWebUrl !== '#' ? `[Web](${safeWebUrl})` : '';
  const appLink = safeAppUrl && safeAppUrl !== '#' ? `[App](${safeAppUrl})` : '';
  const links = [webLink, appLink].filter(Boolean).join(' | ');

  // Build the output
  let output = `${prefix}**${name}** (ID: ${item.id})`;

  // Add metadata line if available
  const metadata = formatMetadata(item);
  if (metadata) {
    output += `\n   ${metadata}`;
  }

  // Add links on separate line if metadata exists, otherwise inline
  if (links) {
    if (metadata) {
      output += `\n   ${links}`;
    } else {
      output += ` - ${links}`;
    }
  }

  return output;
}

/**
 * Formats a list of items with pagination display + "X more" message.
 * @param {Array} items - The array of items to format.
 * @param {string} [itemType] - The type of items (e.g., 'pages', 'sections'). Defaults to 'items'.
 * @param {number} [maxDisplay] - Maximum number of items to display.
 * @param {number} [apiLimit] - The API result limit to check for warning.
 * @returns {object} Object with {list, more, limitWarning} strings.
 */
export function formatItemList(
  items,
  itemType = 'items',
  maxDisplay = DISPLAY_LIMITS.MAX_DISPLAY_ITEMS,
  apiLimit = DISPLAY_LIMITS.API_RESULT_LIMIT
) {
  const displayItems = items.slice(0, maxDisplay);
  const list = displayItems.map((item, i) => formatItemInfo(item, i)).join('\n\n');
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
