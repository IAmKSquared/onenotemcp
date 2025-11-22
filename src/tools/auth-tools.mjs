import { authenticateWithDeviceCode } from '../auth/device-code-flow.mjs';
import { loadExistingToken } from '../auth/token-manager.mjs';
import { initializeGraphClient, getGraphClient } from '../api/graph-client.mjs';

/**
 * Registers authentication-related tools with the MCP server.
 * @param {McpServer} server - The MCP server instance.
 */
export function registerAuthTools(server) {
  server.tool('authenticate', {}, async () => {
    try {
      const result = await authenticateWithDeviceCode();

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

  server.tool('saveAccessToken', {}, async () => {
    try {
      await loadExistingToken();
      const graphClient = initializeGraphClient();

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
