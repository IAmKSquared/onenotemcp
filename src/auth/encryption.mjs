import crypto from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { KeyStorage } from './key-storage.mjs';
import { ENCRYPTION } from '../config/constants.mjs';
import { logger } from '../utils/logger.mjs';

/**
 * Custom error class for decryption failures.
 * Use instanceof DecryptionError to detect decryption-specific errors.
 */
export class DecryptionError extends Error {
  /**
   *
   * @param message
   */
  constructor(message) {
    super(message);
    this.name = 'DecryptionError';
  }
}

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const keyFilePath = path.join(__dirname, '..', '..', '.local-secret-key');

const keyStorage = new KeyStorage(keyFilePath);

/**
 * Retrieves or generates the encryption key from secure storage.
 * @returns {Promise<Buffer>} The 32-byte encryption key.
 */
export async function getEncryptionKey() {
  // Attempt migration from file-based to keyring storage
  await keyStorage.migrate();

  // Try to get existing key
  const existingKey = await keyStorage.getKey();
  if (existingKey) {
    return Buffer.from(existingKey, 'hex');
  }

  // Generate new key
  const newKey = crypto.randomBytes(ENCRYPTION.KEY_LENGTH_BYTES);
  await keyStorage.setKey(newKey.toString('hex'));
  logger.info('🔑 Generated new encryption key and stored securely.');
  return newKey;
}

/**
 * Encrypts text using AES-256-GCM (authenticated encryption).
 * @param {string} text - The text to encrypt.
 * @returns {Promise<object>} The encrypted data { iv, encryptedData, authTag }.
 */
export async function encrypt(text) {
  if (typeof text !== 'string') {
    throw new Error('encrypt() requires a string input');
  }
  const key = await getEncryptionKey();
  const iv = crypto.randomBytes(ENCRYPTION.IV_LENGTH_BYTES);
  const cipher = crypto.createCipheriv(ENCRYPTION.ALGORITHM, key, iv);
  let encrypted = cipher.update(text);
  encrypted = Buffer.concat([encrypted, cipher.final()]);
  const authTag = cipher.getAuthTag();
  return {
    iv: iv.toString('hex'),
    encryptedData: encrypted.toString('hex'),
    authTag: authTag.toString('hex'),
  };
}

/**
 * Decrypts text using AES-256-GCM (authenticated encryption).
 * @param {object} text - The encrypted data object { iv, encryptedData, authTag }.
 * @returns {Promise<string>} The decrypted text.
 * @throws {DecryptionError} If the input object is structurally invalid (wrong type or missing fields).
 * @throws {Error} A native crypto error if GCM authentication fails (data was tampered with or the key is wrong).
 */
export async function decrypt(text) {
  // Validate input structure
  if (!text || typeof text !== 'object') {
    throw new DecryptionError('Invalid encrypted data: expected object');
  }
  if (!text.iv || !text.encryptedData || !text.authTag) {
    throw new DecryptionError(
      'Invalid encrypted data: missing required fields (iv, encryptedData, authTag)'
    );
  }

  const key = await getEncryptionKey();
  const iv = Buffer.from(text.iv, 'hex');
  const encryptedText = Buffer.from(text.encryptedData, 'hex');
  const authTag = Buffer.from(text.authTag, 'hex');
  const decipher = crypto.createDecipheriv(ENCRYPTION.ALGORITHM, key, iv);
  decipher.setAuthTag(authTag);
  let decrypted = decipher.update(encryptedText);
  decrypted = Buffer.concat([decrypted, decipher.final()]);
  return decrypted.toString();
}
