// Utility functions for encrypting/decrypting sensitive data
// In production, use proper encryption libraries and secure key management

class EncryptionService {
  private key: string;

  constructor() {
    // In production, this should come from environment variables or secure key management
    this.key = 'your-encryption-key-here';
  }

  // Simple encryption for demo purposes
  // In production, use proper encryption like AES-256
  encrypt(data: string): string {
    try {
      // This is a very basic encoding - NOT secure for production
      return btoa(data);
    } catch (error) {
      console.error('Encryption failed:', error);
      return data;
    }
  }

  decrypt(encryptedData: string): string {
    try {
      // This is a very basic decoding - NOT secure for production
      return atob(encryptedData);
    } catch (error) {
      console.error('Decryption failed:', error);
      return encryptedData;
    }
  }

  // Encrypt sensitive account data
  encryptAccountData(account: any): any {
    const sensitiveFields = ['accessToken', 'refreshToken'];
    const encryptedAccount = { ...account };

    sensitiveFields.forEach(field => {
      if (encryptedAccount[field]) {
        encryptedAccount[field] = this.encrypt(encryptedAccount[field]);
      }
    });

    return encryptedAccount;
  }

  // Decrypt sensitive account data
  decryptAccountData(encryptedAccount: any): any {
    const sensitiveFields = ['accessToken', 'refreshToken'];
    const decryptedAccount = { ...encryptedAccount };

    sensitiveFields.forEach(field => {
      if (decryptedAccount[field]) {
        decryptedAccount[field] = this.decrypt(decryptedAccount[field]);
      }
    });

    return decryptedAccount;
  }
}

export const encryptionService = new EncryptionService();
export default encryptionService;