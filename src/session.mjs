import { Client } from '@microsoft/microsoft-graph-client';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { decrypt } from './auth/encryption.mjs';
import { TIME_CONVERSION } from './config/constants.mjs';
import { logger } from './utils/logger.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '..', '.access-token.txt');

/**
 * Encapsulates OneNote session state including authentication and graph client.
 */
export class OneNoteSession {
  /**
   *
   */
  constructor() {
    this.accessToken = null;
    this.graphClient = null;
  }

  /**
   * Gets the current access token.
   * @returns {string | null} The current access token or null.
   */
  getAccessToken() {
    return this.accessToken;
  }

  /**
   * Sets the access token.
   * @param {string} token - The access token to set.
   */
  setAccessToken(token) {
    this.accessToken = token;
    // Invalidate graph client when token changes
    this.graphClient = null;
  }

  /**
   * Gets the current graph client instance.
   * @returns {Client | null} The current graph client or null.
   */
  getGraphClient() {
    return this.graphClient;
  }

  /**
   * Initializes the Microsoft Graph client if an access token is available.
   * @returns {Client | null} The initialized Graph client or null.
   */
  initializeGraphClient() {
    if (this.accessToken && !this.graphClient) {
      this.graphClient = Client.init({
        authProvider: (done) => {
          done(null, this.accessToken);
        },
      });
      logger.info('Microsoft Graph client initialized');
    }
    return this.graphClient;
  }

  /**
   * Ensures the Graph client is initialized and authenticated.
   * Loads token if not present, then initializes client.
   * @throws {Error} If no access token is available after attempting to load.
   * @returns {Promise<Client>} The initialized and authenticated Graph client.
   */
  async ensureGraphClient() {
    if (!this.accessToken) {
      await this.loadExistingToken();
    }
    if (!this.accessToken) {
      throw new Error(
        'No access token available. Please authenticate first using the "authenticate" tool.'
      );
    }
    if (!this.graphClient) {
      this.initializeGraphClient();
    }
    return this.graphClient;
  }

  /**
   * Loads an existing access token from the local file system.
   * @returns {Promise<void>}
   */
  async loadExistingToken() {
    try {
      if (fs.existsSync(tokenFilePath)) {
        const fileContent = fs.readFileSync(tokenFilePath, 'utf8');
        let tokenDataStr;

        try {
          // Try to parse as JSON first to see if it's our encrypted format
          const parsed = JSON.parse(fileContent);
          if (parsed.iv && parsed.encryptedData) {
            tokenDataStr = await decrypt(parsed);
            logger.info('🔓 Decrypted token successfully');
          } else {
            // It's JSON but not encrypted (old format)
            tokenDataStr = fileContent;
            logger.warn('⚠️ Loaded unencrypted token (legacy format)');
          }
        } catch (_e) {
          // Not JSON, likely plain text token (very old format)
          tokenDataStr = fileContent;
          logger.warn('⚠️ Loaded raw text token (legacy format)');
        }

        try {
          const parsedToken = JSON.parse(tokenDataStr);

          // Check if token has expired
          if (parsedToken.expiresOn) {
            const expiryDate = new Date(parsedToken.expiresOn);
            const now = new Date();

            if (expiryDate <= now) {
              logger.warn(
                `⚠️ Token has expired (expired on: ${expiryDate.toLocaleString()}). Please re-authenticate`
              );
              return; // Don't set accessToken
            }

            const hoursUntilExpiry = Math.floor((expiryDate - now) / TIME_CONVERSION.MS_PER_HOUR);
            if (hoursUntilExpiry < 24) {
              logger.warn(`⏰ Token expires in ${hoursUntilExpiry} hours`);
            }
          }

          this.accessToken = parsedToken.token;
        } catch (_parseError) {
          this.accessToken = tokenDataStr; // Old format: plain token string
        }
      }
    } catch (error) {
      logger.error({ err: error }, 'Error loading token');
    }
  }

  /**
   * Checks if the session is authenticated.
   * @returns {boolean} True if authenticated, false otherwise.
   */
  isAuthenticated() {
    return this.accessToken !== null;
  }

  /**
   * Clears the session state.
   */
  clear() {
    this.accessToken = null;
    this.graphClient = null;
  }
}
