import { DeviceCodeCredential } from '@azure/identity';
import { saveToken } from './token-manager.mjs';
import { TIMEOUTS, AUTH_CONFIG } from '../config/constants.mjs';
import { logger } from '../utils/logger.mjs';

// --- Configuration ---
const clientId = process.env.AZURE_CLIENT_ID || AUTH_CONFIG.DEFAULT_CLIENT_ID;
const scopes = AUTH_CONFIG.SCOPES;

/**
 * Initiates device code authentication flow.
 * @param {import('../session.mjs').OneNoteSession} session - The session instance.
 * @returns {Promise<{success: boolean, deviceCodeInfo: object | null, authPromise: Promise | null, error: string | null}>}
 */
export async function authenticateWithDeviceCode(session) {
  try {
    logger.info('Starting device code authentication...');
    let deviceCodeInfo = null;

    const credential = new DeviceCodeCredential({
      clientId: clientId,
      userPromptCallback: (info) => {
        deviceCodeInfo = info;
        logger.info(
          `\n=== AUTHENTICATION REQUIRED ===\n${info.message}\n================================\n`
        );
      },
    });

    const authPromise = credential.getToken(scopes);
    await new Promise((resolve) => setTimeout(resolve, TIMEOUTS.DEVICE_CODE_CALLBACK_MS)); // Allow time for userPromptCallback

    if (!deviceCodeInfo) {
      return {
        success: false,
        deviceCodeInfo: null,
        authPromise: null,
        error: 'Could not retrieve device code information.',
      };
    }

    // Set up background token handling
    authPromise
      .then(async (tokenResponse) => {
        session.setAccessToken(tokenResponse.token);
        const tokenData = {
          token: tokenResponse.token,
          clientId: clientId,
          scopes: scopes,
          createdAt: new Date().toISOString(),
          expiresOn: tokenResponse.expiresOnTimestamp
            ? new Date(tokenResponse.expiresOnTimestamp).toISOString()
            : null,
        };

        await saveToken(tokenData);
      })
      .catch((error) => {
        logger.error({ err: error }, 'Background authentication failed');
        session.setAuthError(error?.message || 'Authentication failed');
      });

    return {
      success: true,
      deviceCodeInfo,
      authPromise,
      error: null,
    };
  } catch (error) {
    return {
      success: false,
      deviceCodeInfo: null,
      authPromise: null,
      error: error.message,
    };
  }
}

/**
 * Gets the client ID being used.
 * @returns {string} The client ID.
 */
export function getClientId() {
  return clientId;
}

/**
 * Gets the scopes being requested.
 * @returns {string[]} The scopes array.
 */
export function getScopes() {
  return scopes;
}
