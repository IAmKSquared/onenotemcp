import { Client } from '@microsoft/microsoft-graph-client';
import { readFile } from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { decrypt, DecryptionError } from './auth/encryption.mjs';
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
    this.authError = null;
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
    // Clear any previous auth error on successful token set
    this.authError = null;
  }

  /**
   * Gets the current auth error, if any.
   * @returns {string | null} The auth error message or null.
   */
  getAuthError() {
    return this.authError;
  }

  /**
   * Sets an authentication error.
   * @param {string} error - The error message.
   */
  setAuthError(error) {
    this.authError = error;
    this.accessToken = null;
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
      // Provide more specific error if auth previously failed
      if (this.authError) {
        throw new Error(
          `Authentication failed: ${this.authError}. Please try authenticating again using the "authenticate" tool.`
        );
      }
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
    let fileContent;
    try {
      fileContent = await readFile(tokenFilePath, 'utf8');
    } catch (error) {
      if (error.code === 'ENOENT') {
        return; // File doesn't exist, nothing to load
      }
      logger.error({ err: error }, 'Error reading token file');
      return;
    }

    let tokenDataStr;

    try {
      // Try to parse as JSON first to see if it's our encrypted format
      const parsed = JSON.parse(fileContent);
      if (parsed.iv && parsed.encryptedData && parsed.authTag) {
        tokenDataStr = await decrypt(parsed);
        logger.info('🔓 Decrypted token successfully');
      } else {
        // It's JSON but not encrypted (old format)
        tokenDataStr = fileContent;
        logger.warn('⚠️ Loaded unencrypted token (legacy format)');
      }
    } catch (decryptError) {
      if (decryptError instanceof DecryptionError) {
        logger.error({ err: decryptError }, 'Failed to decrypt token - file may be corrupted');
        return;
      }
      // Not JSON, likely plain text token (very old format)
      tokenDataStr = fileContent;
      logger.warn('⚠️ Loaded raw text token (legacy format)');
    }

    // Parse the token structure
    let parsedToken;
    try {
      parsedToken = JSON.parse(tokenDataStr);
    } catch (_parseError) {
      // Old format: plain token string (no JSON structure)
      if (typeof tokenDataStr === 'string' && tokenDataStr.length > 0) {
        this.accessToken = tokenDataStr;
      }
      return;
    }

    if (!parsedToken?.token) {
      logger.error('Token file has invalid structure');
      return;
    }

    // Check if token has expired
    if (parsedToken.expiresOn) {
      const expiryDate = new Date(parsedToken.expiresOn);
      const now = new Date();

      if (expiryDate <= now) {
        logger.warn(
          `⚠️ Token has expired (expired on: ${expiryDate.toLocaleString()}). Please re-authenticate`
        );
        return;
      }

      const hoursUntilExpiry = Math.floor((expiryDate - now) / TIME_CONVERSION.MS_PER_HOUR);
      if (hoursUntilExpiry < 24) {
        logger.warn(`⏰ Token expires in ${hoursUntilExpiry} hours`);
      }
    }

    this.accessToken = parsedToken.token;
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
