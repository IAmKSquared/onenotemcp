import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { encrypt, decrypt } from './encryption.mjs';

// --- Configuration ---
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, '..', '..', '.access-token.txt');

// --- Global State ---
let accessToken = null;

/**
 * Gets the current access token.
 * @returns {string | null} The current access token or null.
 */
export function getAccessToken() {
  return accessToken;
}

/**
 * Sets the access token.
 * @param {string} token - The access token to set.
 */
export function setAccessToken(token) {
  accessToken = token;
}

/**
 * Loads an existing access token from the local file system.
 * @returns {Promise<void>}
 */
export async function loadExistingToken() {
  try {
    if (fs.existsSync(tokenFilePath)) {
      const fileContent = fs.readFileSync(tokenFilePath, 'utf8');
      let tokenDataStr;

      try {
        // Try to parse as JSON first to see if it's our encrypted format
        const parsed = JSON.parse(fileContent);
        if (parsed.iv && parsed.encryptedData) {
          tokenDataStr = await decrypt(parsed);
          console.error('🔓 Decrypted token successfully.');
        } else {
          // It's JSON but not encrypted (old format)
          tokenDataStr = fileContent;
          console.error('⚠️  Loaded unencrypted token (legacy format).');
        }
      } catch (_e) {
        // Not JSON, likely plain text token (very old format)
        tokenDataStr = fileContent;
        console.error('⚠️  Loaded raw text token (legacy format).');
      }

      try {
        const parsedToken = JSON.parse(tokenDataStr);

        // Check if token has expired
        if (parsedToken.expiresOn) {
          const expiryDate = new Date(parsedToken.expiresOn);
          const now = new Date();

          if (expiryDate <= now) {
            console.error('⚠️  Token has expired. Please re-authenticate.');
            console.error(`   Expired on: ${expiryDate.toLocaleString()}`);
            return; // Don't set accessToken
          }

          const hoursUntilExpiry = Math.floor((expiryDate - now) / (1000 * 60 * 60));
          if (hoursUntilExpiry < 24) {
            console.error(`⏰ Token expires in ${hoursUntilExpiry} hours`);
          }
        }

        accessToken = parsedToken.token;
      } catch (_parseError) {
        accessToken = tokenDataStr; // Old format: plain token string
      }
    }
  } catch (error) {
    console.error(`Error loading token: ${error.message}`);
  }
}

/**
 * Saves an access token to the local file system (encrypted).
 * @param {object} tokenData - The token data object to save.
 * @returns {Promise<void>}
 */
export async function saveToken(tokenData) {
  // Encrypt before saving
  const encryptedToken = await encrypt(JSON.stringify(tokenData));
  fs.writeFileSync(tokenFilePath, JSON.stringify(encryptedToken, null, 2), {
    mode: 0o600,
  });
  console.error('🔒 Token saved securely (encrypted).');
}
