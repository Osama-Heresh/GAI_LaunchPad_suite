import React, { createContext, useContext, useState, useEffect, ReactNode, useCallback } from 'react';
import { useAuth } from './AuthContext';

export interface SocialAccount {
  id: string;
  platform: string;
  username: string;
  displayName: string;
  profileImage?: string;
  isConnected: boolean;
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: Date;
  permissions: string[];
  lastSync: Date;
  status: 'active' | 'expired' | 'error' | 'pending';
  errorMessage?: string;
}

export interface ConnectionStatus {
  platform: string;
  isConnecting: boolean;
  error?: string;
}

interface SocialAccountsContextType {
  accounts: SocialAccount[];
  connectionStatus: ConnectionStatus[];
  connectAccount: (platform: string) => Promise<void>;
  disconnectAccount: (accountId: string) => Promise<void>;
  refreshAccount: (accountId: string) => Promise<void>;
  getAccountsByPlatform: (platform: string) => SocialAccount[];
  isAccountConnected: (platform: string) => boolean;
  getActiveAccount: (platform: string) => SocialAccount | null;
  isLoading: boolean;
}

const SocialAccountsContext = createContext<SocialAccountsContextType | undefined>(undefined);

export const useSocialAccounts = () => {
  const context = useContext(SocialAccountsContext);
  if (context === undefined) {
    throw new Error('useSocialAccounts must be used within a SocialAccountsProvider');
  }
  return context;
};

export const SocialAccountsProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const { user } = useAuth();
  const [accounts, setAccounts] = useState<SocialAccount[]>([]);
  const [connectionStatus, setConnectionStatus] = useState<ConnectionStatus[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  // Load accounts from localStorage when user changes
  useEffect(() => {
    if (user?.id) {
      try {
        const savedAccounts = localStorage.getItem(`social_accounts_${user.id}`);
        if (savedAccounts) {
          const parsedAccounts = JSON.parse(savedAccounts).map((account: any) => ({
            ...account,
            expiresAt: account.expiresAt ? new Date(account.expiresAt) : undefined,
            lastSync: new Date(account.lastSync)
          }));
          setAccounts(parsedAccounts);
        } else {
          setAccounts([]);
        }
      } catch (error) {
        console.error('Failed to load social accounts:', error);
        setAccounts([]);
      }
    } else {
      setAccounts([]);
    }
  }, [user?.id]);

  const saveAccounts = useCallback((newAccounts: SocialAccount[]) => {
    if (!user?.id) return;
    
    try {
      localStorage.setItem(`social_accounts_${user.id}`, JSON.stringify(newAccounts));
      setAccounts(newAccounts);
    } catch (error) {
      console.error('Failed to save social accounts:', error);
    }
  }, [user?.id]);

  const setConnectionStatusForPlatform = useCallback((platform: string, status: Partial<ConnectionStatus>) => {
    setConnectionStatus(prev => {
      const existing = prev.find(s => s.platform === platform);
      if (existing) {
        return prev.map(s => s.platform === platform ? { ...s, ...status } : s);
      } else {
        return [...prev, { platform, isConnecting: false, ...status }];
      }
    });
  }, []);

  const connectAccount = useCallback(async (platform: string): Promise<void> => {
    setConnectionStatusForPlatform(platform, { isConnecting: true, error: undefined });
    
    try {
      // Simulate OAuth flow - in real implementation, this would redirect to OAuth provider
      const result = await simulateOAuthFlow(platform);
      
      if (result.success) {
        const newAccount: SocialAccount = {
          id: `${platform}_${Date.now()}`,
          platform,
          username: result.username,
          displayName: result.displayName,
          profileImage: result.profileImage,
          isConnected: true,
          accessToken: result.accessToken,
          refreshToken: result.refreshToken,
          expiresAt: result.expiresAt,
          permissions: result.permissions,
          lastSync: new Date(),
          status: 'active'
        };

        const updatedAccounts = [...accounts.filter(a => a.platform !== platform), newAccount];
        saveAccounts(updatedAccounts);
        
        setConnectionStatusForPlatform(platform, { isConnecting: false });
      } else {
        throw new Error(result.error || 'Connection failed');
      }
    } catch (error: any) {
      console.error(`Failed to connect ${platform}:`, error);
      setConnectionStatusForPlatform(platform, { 
        isConnecting: false, 
        error: error.message || 'Connection failed' 
      });
      throw error;
    }
  }, [accounts, saveAccounts, setConnectionStatusForPlatform]);

  const disconnectAccount = useCallback(async (accountId: string): Promise<void> => {
    try {
      const account = accounts.find(a => a.id === accountId);
      if (!account) throw new Error('Account not found');

      // In real implementation, revoke tokens on the platform
      await revokeAccountAccess(account);

      const updatedAccounts = accounts.filter(a => a.id !== accountId);
      saveAccounts(updatedAccounts);
    } catch (error: any) {
      console.error('Failed to disconnect account:', error);
      throw error;
    }
  }, [accounts, saveAccounts]);

  const refreshAccount = useCallback(async (accountId: string): Promise<void> => {
    try {
      const account = accounts.find(a => a.id === accountId);
      if (!account) throw new Error('Account not found');

      // Simulate token refresh
      const refreshResult = await refreshAccountTokens(account);
      
      if (refreshResult.success) {
        const updatedAccounts = accounts.map(a => 
          a.id === accountId 
            ? {
                ...a,
                accessToken: refreshResult.accessToken,
                refreshToken: refreshResult.refreshToken,
                expiresAt: refreshResult.expiresAt,
                lastSync: new Date(),
                status: 'active' as const,
                errorMessage: undefined
              }
            : a
        );
        saveAccounts(updatedAccounts);
      } else {
        throw new Error(refreshResult.error || 'Refresh failed');
      }
    } catch (error: any) {
      console.error('Failed to refresh account:', error);
      
      // Mark account as having an error
      const updatedAccounts = accounts.map(a => 
        a.id === accountId 
          ? { ...a, status: 'error' as const, errorMessage: error.message }
          : a
      );
      saveAccounts(updatedAccounts);
      throw error;
    }
  }, [accounts, saveAccounts]);

  const getAccountsByPlatform = useCallback((platform: string): SocialAccount[] => {
    return accounts.filter(account => account.platform === platform);
  }, [accounts]);

  const isAccountConnected = useCallback((platform: string): boolean => {
    return accounts.some(account => account.platform === platform && account.isConnected);
  }, [accounts]);

  const getActiveAccount = useCallback((platform: string): SocialAccount | null => {
    const platformAccounts = accounts.filter(account => account.platform === platform);
    return platformAccounts.find(account => account.status === 'active') || null;
  }, [accounts]);

  const contextValue = {
    accounts,
    connectionStatus,
    connectAccount,
    disconnectAccount,
    refreshAccount,
    getAccountsByPlatform,
    isAccountConnected,
    getActiveAccount,
    isLoading
  };

  return (
    <SocialAccountsContext.Provider value={contextValue}>
      {children}
    </SocialAccountsContext.Provider>
  );
};

// Simulated OAuth flow - replace with real OAuth implementation
async function simulateOAuthFlow(platform: string): Promise<{
  success: boolean;
  username?: string;
  displayName?: string;
  profileImage?: string;
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: Date;
  permissions?: string[];
  error?: string;
}> {
  // Simulate network delay
  await new Promise(resolve => setTimeout(resolve, 2000));

  // Simulate random success/failure for demo
  if (Math.random() > 0.1) { // 90% success rate
    const mockData = {
      twitter: {
        username: '@johndoe',
        displayName: 'John Doe',
        profileImage: 'https://images.pexels.com/photos/220453/pexels-photo-220453.jpeg?auto=compress&cs=tinysrgb&w=100',
        permissions: ['read', 'write', 'manage_tweets']
      },
      facebook: {
        username: 'john.doe.123',
        displayName: 'John Doe',
        profileImage: 'https://images.pexels.com/photos/220453/pexels-photo-220453.jpeg?auto=compress&cs=tinysrgb&w=100',
        permissions: ['pages_manage_posts', 'pages_read_engagement']
      },
      instagram: {
        username: 'johndoe_official',
        displayName: 'John Doe',
        profileImage: 'https://images.pexels.com/photos/220453/pexels-photo-220453.jpeg?auto=compress&cs=tinysrgb&w=100',
        permissions: ['instagram_basic', 'instagram_content_publish']
      },
      linkedin: {
        username: 'john-doe-123',
        displayName: 'John Doe',
        profileImage: 'https://images.pexels.com/photos/220453/pexels-photo-220453.jpeg?auto=compress&cs=tinysrgb&w=100',
        permissions: ['w_member_social', 'r_liteprofile']
      }
    };

    const platformData = mockData[platform as keyof typeof mockData] || {
      username: `user_${platform}`,
      displayName: 'User Name',
      permissions: ['basic']
    };

    return {
      success: true,
      ...platformData,
      accessToken: `mock_access_token_${platform}_${Date.now()}`,
      refreshToken: `mock_refresh_token_${platform}_${Date.now()}`,
      expiresAt: new Date(Date.now() + 60 * 60 * 1000) // 1 hour from now
    };
  } else {
    return {
      success: false,
      error: 'User denied access or connection failed'
    };
  }
}

// Simulated token refresh
async function refreshAccountTokens(account: SocialAccount): Promise<{
  success: boolean;
  accessToken?: string;
  refreshToken?: string;
  expiresAt?: Date;
  error?: string;
}> {
  await new Promise(resolve => setTimeout(resolve, 1000));

  if (Math.random() > 0.05) { // 95% success rate
    return {
      success: true,
      accessToken: `refreshed_access_token_${account.platform}_${Date.now()}`,
      refreshToken: account.refreshToken,
      expiresAt: new Date(Date.now() + 60 * 60 * 1000)
    };
  } else {
    return {
      success: false,
      error: 'Token refresh failed - user may need to reconnect'
    };
  }
}

// Simulated token revocation
async function revokeAccountAccess(account: SocialAccount): Promise<void> {
  await new Promise(resolve => setTimeout(resolve, 500));
  console.log(`Revoked access for ${account.platform} account: ${account.username}`);
}