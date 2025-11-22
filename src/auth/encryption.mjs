import crypto from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { KeyStorage } from '../../key-storage.mjs';

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const keyFilePath = path.join(__dirname, '..', '..', '.local-secret-key');

// --- Encryption Configuration ---
const ALGORITHM = 'aes-256-cbc';
const IV_LENGTH = 16;
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
  const newKey = crypto.randomBytes(32);
  await keyStorage.setKey(newKey.toString('hex'));
  console.error('🔑 Generated new encryption key and stored securely.');
  return newKey;
}

/**
 * Encrypts text using AES-256-CBC.
 * @param {string} text - The text to encrypt.
 * @returns {Promise<object>} The encrypted data { iv, encryptedData }.
 */
export async function encrypt(text) {
  const key = await getEncryptionKey();
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  let encrypted = cipher.update(text);
  encrypted = Buffer.concat([encrypted, cipher.final()]);
  return { iv: iv.toString('hex'), encryptedData: encrypted.toString('hex') };
}

/**
 * Decrypts text using AES-256-CBC.
 * @param {object} text - The encrypted data object { iv, encryptedData }.
 * @returns {Promise<string>} The decrypted text.
 */
export async function decrypt(text) {
  const key = await getEncryptionKey();
  const iv = Buffer.from(text.iv, 'hex');
  const encryptedText = Buffer.from(text.encryptedData, 'hex');
  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  let decrypted = decipher.update(encryptedText);
  decrypted = Buffer.concat([decrypted, decipher.final()]);
  return decrypted.toString();
}
