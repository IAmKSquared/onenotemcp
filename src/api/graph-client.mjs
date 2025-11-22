import { Client } from '@microsoft/microsoft-graph-client';
import { getAccessToken, loadExistingToken } from '../auth/token-manager.mjs';
import { logger } from '../utils/logger.mjs';

// --- Global State ---
let graphClient = null;

/**
 * Initializes the Microsoft Graph client if an access token is available.
 * @returns {Client | null} The initialized Graph client or null.
 */
export function initializeGraphClient() {
  const accessToken = getAccessToken();
  if (accessToken && !graphClient) {
    graphClient = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
    logger.info('Microsoft Graph client initialized');
  }
  return graphClient;
}

/**
 * Gets the current graph client instance.
 * @returns {Client | null} The current graph client or null.
 */
export function getGraphClient() {
  return graphClient;
}

/**
 * Ensures the Graph client is initialized and authenticated.
 * Loads token if not present, then initializes client.
 * @throws {Error} If no access token is available after attempting to load.
 * @returns {Promise<Client>} The initialized and authenticated Graph client.
 */
export async function ensureGraphClient() {
  const accessToken = getAccessToken();
  if (!accessToken) {
    await loadExistingToken();
  }
  if (!getAccessToken()) {
    throw new Error(
      'No access token available. Please authenticate first using the "authenticate" tool.'
    );
  }
  if (!graphClient) {
    initializeGraphClient();
  }
  return graphClient;
}
