// @napi-rs/keyring uses CJS exports; destructure after default import
import keyringPkg from '@napi-rs/keyring';
const { Entry } = keyringPkg;
import { readFile, writeFile, unlink } from 'fs/promises';
import { logger } from '../utils/logger.mjs';

const SERVICE_NAME = 'onenote-mcp';
const KEY_NAME = 'encryption-key';

/**
 * Secure key storage using OS-native keychain/credential manager.
 * Falls back to file-based storage if keyring is unavailable.
 */
export class KeyStorage {
  /**
   * Creates a new KeyStorage instance.
   * @param {string} fallbackFilePath - Path to fallback file if keyring unavailable.
   */
  constructor(fallbackFilePath) {
    this.entry = null;
    this.fallbackFilePath = fallbackFilePath;
    this.useFallback = false;

    try {
      this.entry = new Entry(SERVICE_NAME, KEY_NAME);
    } catch (error) {
      logger.warn({ err: error }, '⚠️ OS keychain unavailable, using encrypted file storage');
      this.useFallback = true;
    }
  }

  /**
   * Retrieves the encryption key from secure storage.
   * @returns {Promise<string|null>} The key as a hex string, or null if not found.
   */
  async getKey() {
    if (this.useFallback) {
      return await this._getKeyFromFile();
    }

    try {
      const key = await this.entry.getPassword();
      return key;
    } catch (_error) {
      // Key doesn't exist in keyring
      return null;
    }
  }

  /**
   * Stores the encryption key in secure storage.
   * @param {string} keyHex - The key as a hex string.
   * @returns {Promise<void>}
   */
  async setKey(keyHex) {
    if (this.useFallback) {
      return await this._setKeyToFile(keyHex);
    }

    try {
      await this.entry.setPassword(keyHex);
    } catch (error) {
      logger.warn(
        { err: error },
        '⚠️ Failed to store key in OS keychain, falling back to file storage'
      );
      this.useFallback = true;
      return await this._setKeyToFile(keyHex);
    }
  }

  /**
   * Deletes the encryption key from secure storage.
   * @returns {Promise<void>}
   */
  async deleteKey() {
    if (this.useFallback) {
      return await this._deleteKeyFromFile();
    }

    try {
      await this.entry.deletePassword();
    } catch (_error) {
      // Key might not exist, that's okay
    }
  }

  /**
   * Checks if a key exists in secure storage.
   * @returns {Promise<boolean>} True if key exists, false otherwise.
   */
  async hasKey() {
    const key = await this.getKey();
    return key !== null;
  }

  /**
   * Migrates key from file-based storage to OS keychain.
   * @returns {Promise<boolean>} True if migration occurred, false if not needed.
   */
  async migrate() {
    // If already using fallback, no migration possible
    if (this.useFallback) {
      return false;
    }

    let keyHex;
    try {
      // Read key from file (handles ENOENT if file doesn't exist)
      keyHex = (await readFile(this.fallbackFilePath, 'utf8')).trim();
    } catch (error) {
      if (error.code === 'ENOENT') {
        return false; // No file to migrate
      }
      logger.warn({ err: error }, '⚠️ Failed to read key file for migration');
      return false;
    }

    try {
      // Store in keyring
      await this.setKey(keyHex);

      // Delete old file
      await unlink(this.fallbackFilePath);

      logger.info('✅ Successfully migrated encryption key from file to OS keychain');
      return true;
    } catch (error) {
      logger.warn({ err: error }, '⚠️ Failed to migrate key');
      return false;
    }
  }

  /**
   * Gets key from fallback file storage.
   * @returns {Promise<string|null>} The key as hex string, or null if not found.
   * @private
   */
  async _getKeyFromFile() {
    try {
      return (await readFile(this.fallbackFilePath, 'utf8')).trim();
    } catch (error) {
      if (error.code === 'ENOENT') {
        return null; // File doesn't exist
      }
      logger.error({ err: error }, 'Error reading key file');
      return null;
    }
  }

  /**
   * Stores key to fallback file storage.
   * @param {string} keyHex - The key as hex string.
   * @returns {Promise<void>}
   * @private
   */
  async _setKeyToFile(keyHex) {
    try {
      await writeFile(this.fallbackFilePath, keyHex, { mode: 0o600 });
    } catch (error) {
      throw new Error(`Failed to write key file: ${error.message}`);
    }
  }

  /**
   * Deletes key from fallback file storage.
   * @returns {Promise<void>}
   * @private
   */
  async _deleteKeyFromFile() {
    try {
      await unlink(this.fallbackFilePath);
    } catch (error) {
      if (error.code !== 'ENOENT') {
        logger.error({ err: error }, 'Error deleting key file');
      }
      // ENOENT is fine - file already doesn't exist
    }
  }
}
