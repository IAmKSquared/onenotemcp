import crypto from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { KeyStorage } from './key-storage.mjs';
import { ENCRYPTION } from '../config/constants.mjs';
import { logger } from '../utils/logger.mjs';

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
 * @throws {Error} If the authentication tag is invalid (data was tampered with).
 */
export async function decrypt(text) {
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
