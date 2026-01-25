import { writeFile } from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { encrypt } from './encryption.mjs';
import { logger } from '../utils/logger.mjs';

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '..', '..', '.access-token.txt');

/**
 * Saves an access token to the local file system (encrypted).
 * @param {object} tokenData - The token data object to save.
 * @returns {Promise<void>}
 */
export async function saveToken(tokenData) {
  const encryptedToken = await encrypt(JSON.stringify(tokenData));
  await writeFile(tokenFilePath, JSON.stringify(encryptedToken, null, 2), { mode: 0o600 });
  logger.info('🔒 Token saved securely (encrypted)');
}
