import { authenticateWithDeviceCode } from '../auth/device-code-flow.mjs';

/**
 * Registers authentication-related tools with the MCP server.
 *
 * DESIGN NOTE: These tools intentionally DO NOT use createToolHandler wrapper.
 *
 * Rationale:
 * 1. Bootstrap Authentication: These tools establish authentication and cannot call
 *    ensureGraphClient() - they're creating the session that other tools depend on.
 *
 * 2. Interactive Operations: Unlike API tools, these handle user-driven auth flows,
 *    not HTTP API calls. Retry logic doesn't apply to interactive steps.
 *
 * 3. Custom Error Context: Auth errors require user-facing instructions (visit URL,
 *    enter code) rather than HTTP status-based error messages (401, 429, etc.).
 *
 * 4. Simple Error Handling: Auth flow errors are straightforward state issues
 *    (no token, auth pending) that don't benefit from the retry/backoff logic
 *    designed for transient API failures.
 *
 * This is an intentional architectural decision, not an oversight.
 * @param {McpServer} server - The MCP server instance.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 */
export function registerAuthTools(server, session) {
  /**
   * Initiates device code authentication flow.
   * Returns instructions for the user to complete authentication in their browser.
   */
  server.tool('authenticate', {}, async () => {
    try {
      const result = await authenticateWithDeviceCode(session);

      if (!result.success) {
        return {
          isError: true,
          content: [
            {
              type: 'text',
              text:
                result.error ||
                'Could not retrieve device code information. Please try again or check console logs.',
            },
          ],
        };
      }

      const authMessage = `🔐 **AUTHENTICATION REQUIRED**

Please complete the following steps:
1. **Open this URL in your browser:** https://microsoft.com/devicelogin
2. **Enter this code:** ${result.deviceCodeInfo.userCode}
3. **Sign in with your Microsoft account that has OneNote access.**
4. **After completing authentication, use the 'saveAccessToken' tool.**

Token will be saved automatically upon successful browser authentication.`;

      return { content: [{ type: 'text', text: authMessage }] };
    } catch (error) {
      return {
        isError: true,
        content: [{ type: 'text', text: `Authentication failed: ${error.message}` }],
      };
    }
  });

  /**
   * Loads and verifies the saved authentication token.
   * Called after the user completes browser authentication to confirm the token is valid.
   */
  server.tool('saveAccessToken', {}, async () => {
    try {
      await session.loadExistingToken();
      const graphClient = session.initializeGraphClient();

      if (graphClient) {
        const testResponse = await graphClient.api('/me').get();
        return {
          content: [
            {
              type: 'text',
              text: `✅ **Authentication Successful!**
    Token loaded and verified.
**Account Info:**
    - Name: ${testResponse.displayName || 'Unknown'}
    - Email: ${testResponse.userPrincipalName || 'Unknown'}
🚀 You can now use OneNote tools!`,
            },
          ],
        };
      } else {
        return {
          isError: true,
          content: [
            {
              type: 'text',
              text: `❌ **No Token Found.** Please run the 'authenticate' tool first.`,
            },
          ],
        };
      }
    } catch (error) {
      return {
        isError: true,
        content: [{ type: 'text', text: `Failed to load or verify token: ${error.message}` }],
      };
    }
  });
}
