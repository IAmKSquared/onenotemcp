import { DeviceCodeCredential } from '@azure/identity';
import { setAccessToken } from './token-manager.mjs';
import { saveToken } from './token-manager.mjs';

// --- Configuration ---
const clientId = process.env.AZURE_CLIENT_ID || '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Default: Microsoft Graph Explorer App ID
const scopes = ['Notes.Read', 'Notes.ReadWrite', 'Notes.Create', 'User.Read'];

/**
 * Initiates device code authentication flow.
 * @returns {Promise<{success: boolean, deviceCodeInfo: object | null, authPromise: Promise | null, error: string | null}>}
 */
export async function authenticateWithDeviceCode() {
  try {
    console.error('Starting device code authentication...');
    let deviceCodeInfo = null;

    const credential = new DeviceCodeCredential({
      clientId: clientId,
      userPromptCallback: (info) => {
        deviceCodeInfo = info;
        console.error(
          `\n=== AUTHENTICATION REQUIRED ===\n${info.message}\n================================\n`
        );
      },
    });

    const authPromise = credential.getToken(scopes);
    await new Promise((resolve) => setTimeout(resolve, 2000)); // Allow time for userPromptCallback

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
        setAccessToken(tokenResponse.token);
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
        console.error(`Background authentication failed: ${error.message}`);
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
