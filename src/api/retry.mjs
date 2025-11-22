/**
 * Retry helper with exponential backoff for transient failures.
 * @param {Function} fn - The async function to retry.
 * @param {number} maxRetries - Maximum number of retry attempts.
 * @param {number} baseDelay - Base delay in milliseconds for exponential backoff.
 * @returns {Promise} The result of the function call.
 */
export async function retryWithBackoff(fn, maxRetries = 3, baseDelay = 1000) {
  let lastError;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error;

      // Don't retry on these errors
      const nonRetryableErrors = [400, 401, 403, 404, 409];
      if (error.statusCode && nonRetryableErrors.includes(error.statusCode)) {
        throw error;
      }

      // Retry on rate limits (429) and server errors (500+)
      const shouldRetry =
        error.statusCode === 429 || // Rate limit
        (error.statusCode >= 500 && error.statusCode < 600) || // Server errors
        error.code === 'ETIMEDOUT' || // Timeout
        error.code === 'ECONNRESET'; // Connection reset

      if (!shouldRetry || attempt === maxRetries) {
        throw error;
      }

      // Calculate delay with exponential backoff
      const delay = baseDelay * Math.pow(2, attempt);
      console.error(
        `⏳ Retry attempt ${attempt + 1}/${maxRetries} after ${delay}ms (Error: ${error.statusCode || error.code})`
      );
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }

  throw lastError;
}

/**
 * Provides detailed, actionable error messages based on error type.
 * @param {Error} error - The error object from Graph API or other sources.
 * @param {string} errorPrefix - The prefix for the error message.
 * @returns {string} A user-friendly error message with guidance.
 */
export function getDetailedErrorMessage(error, errorPrefix) {
  const statusCode = error.statusCode || error.status;
  const errorMessage = error.message || 'Unknown error';

  // Authentication errors
  if (
    statusCode === 401 ||
    errorMessage.includes('authenticate') ||
    errorMessage.includes('Access token')
  ) {
    return `🔐 **Authentication Required**\nYour access token has expired or is invalid. Please run the 'authenticate' tool to sign in again.`;
  }

  // Permission errors
  if (statusCode === 403) {
    return `🔒 **Permission Denied**\nYou don't have permission to perform this action. Ensure your account has the required OneNote permissions (Notes.Read, Notes.ReadWrite, Notes.Create).`;
  }

  // Not found errors
  if (statusCode === 404) {
    return `❌ **Resource Not Found**\n${errorMessage}\n\nThe requested item doesn't exist. It may have been deleted or the ID is incorrect.`;
  }

  // Rate limit errors
  if (statusCode === 429) {
    return `⏱️ **Rate Limit Exceeded**\nToo many requests. Please wait a moment before trying again. (All retry attempts exhausted)`;
  }

  // Server errors
  if (statusCode >= 500 && statusCode < 600) {
    return `🔧 **Server Error** (${statusCode})\nMicrosoft's OneNote service is experiencing issues. Please try again in a few moments. (All retry attempts exhausted)`;
  }

  // Timeout errors
  if (error.code === 'ETIMEDOUT' || errorMessage.includes('timeout')) {
    return `⏰ **Request Timeout**\nThe request took too long to complete. Please check your network connection and try again.`;
  }

  // Network errors
  if (error.code === 'ECONNRESET' || error.code === 'ENOTFOUND' || error.code === 'ECONNREFUSED') {
    return `🌐 **Network Error**\nUnable to connect to Microsoft services. Please check your internet connection.`;
  }

  // Token expiration (specific message format from Azure)
  if (errorMessage.includes('token') && errorMessage.includes('expired')) {
    return `🔐 **Token Expired**\nYour authentication token has expired. Please run the 'authenticate' tool to sign in again.`;
  }

  // Default error message
  return `❌ **${errorPrefix}**\n${errorMessage}`;
}

/**
 * Creates a standardized tool handler with error handling and authentication checks.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 * @param {Function} handler - The async tool implementation function.
 * @param {string} errorPrefix - The prefix for error messages.
 * @returns {Function} A wrapped handler function compatible with McpServer.
 */
export function createToolHandler(session, handler, errorPrefix = 'Tool execution failed') {
  return async (args) => {
    try {
      // 'authenticate' tool is a special case that establishes the session,
      // so we don't check for ensureGraphClient() inside it to avoid recursion/errors.
      if (handler.name !== 'authenticateHandler') {
        await session.ensureGraphClient();
      }

      // Wrap handler with retry logic for transient failures
      return await retryWithBackoff(async () => {
        return await handler(args);
      });
    } catch (error) {
      const detailedMessage = getDetailedErrorMessage(error, errorPrefix);
      return { isError: true, content: [{ type: 'text', text: detailedMessage }] };
    }
  };
}
