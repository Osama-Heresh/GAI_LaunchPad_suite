import React, { useState } from 'react';
import { 
  Plus, 
  Settings, 
  RefreshCw, 
  Trash2, 
  CheckCircle, 
  XCircle, 
  AlertTriangle,
  Twitter,
  Facebook,
  Instagram,
  Linkedin,
  Hash,
  MessageSquare,
  Shield,
  Clock,
  User
} from 'lucide-react';
import { useSocialAccounts, SocialAccount } from '../contexts/SocialAccountsContext';

interface SocialAccountManagerProps {
  onClose?: () => void;
}

const SocialAccountManager: React.FC<SocialAccountManagerProps> = ({ onClose }) => {
  const {
    accounts,
    connectionStatus,
    connectAccount,
    disconnectAccount,
    refreshAccount,
    isAccountConnected
  } = useSocialAccounts();

  const [showConfirmDisconnect, setShowConfirmDisconnect] = useState<string | null>(null);

  const platforms = [
    { id: 'twitter', name: 'Twitter (X)', icon: Twitter, color: 'from-sky-500 to-blue-500' },
    { id: 'facebook', name: 'Facebook', icon: Facebook, color: 'from-blue-600 to-blue-500' },
    { id: 'instagram', name: 'Instagram', icon: Instagram, color: 'from-pink-600 to-purple-500' },
    { id: 'linkedin', name: 'LinkedIn', icon: Linkedin, color: 'from-blue-700 to-blue-600' },
    { id: 'tiktok', name: 'TikTok', icon: Hash, color: 'from-pink-500 to-red-500' },
    { id: 'slack', name: 'Slack', icon: MessageSquare, color: 'from-purple-600 to-pink-600' },
    { id: 'discord', name: 'Discord', icon: MessageSquare, color: 'from-indigo-600 to-purple-600' }
  ];

  const handleConnect = async (platform: string) => {
    try {
      await connectAccount(platform);
    } catch (error) {
      console.error('Connection failed:', error);
    }
  };

  const handleDisconnect = async (accountId: string) => {
    try {
      await disconnectAccount(accountId);
      setShowConfirmDisconnect(null);
    } catch (error) {
      console.error('Disconnection failed:', error);
    }
  };

  const handleRefresh = async (accountId: string) => {
    try {
      await refreshAccount(accountId);
    } catch (error) {
      console.error('Refresh failed:', error);
    }
  };

  const getStatusIcon = (account: SocialAccount) => {
    switch (account.status) {
      case 'active':
        return <CheckCircle className="h-4 w-4 text-green-400" />;
      case 'expired':
        return <Clock className="h-4 w-4 text-yellow-400" />;
      case 'error':
        return <XCircle className="h-4 w-4 text-red-400" />;
      case 'pending':
        return <AlertTriangle className="h-4 w-4 text-orange-400" />;
      default:
        return <XCircle className="h-4 w-4 text-gray-400" />;
    }
  };

  const getStatusText = (account: SocialAccount) => {
    switch (account.status) {
      case 'active':
        return 'Connected';
      case 'expired':
        return 'Token Expired';
      case 'error':
        return 'Error';
      case 'pending':
        return 'Pending';
      default:
        return 'Disconnected';
    }
  };

  const isConnecting = (platform: string) => {
    return connectionStatus.find(s => s.platform === platform)?.isConnecting || false;
  };

  const getConnectionError = (platform: string) => {
    return connectionStatus.find(s => s.platform === platform)?.error;
  };

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700 max-h-[80vh] overflow-y-auto">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center space-x-3">
          <Shield className="h-6 w-6 text-blue-400" />
          <h2 className="text-xl font-semibold text-white">Social Media Accounts</h2>
        </div>
        {onClose && (
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-white transition-colors"
          >
            ✕
          </button>
        )}
      </div>

      <div className="space-y-6">
        {/* Connected Accounts */}
        {accounts.length > 0 && (
          <div>
            <h3 className="text-lg font-medium text-white mb-4">Connected Accounts</h3>
            <div className="space-y-3">
              {accounts.map((account) => {
                const platform = platforms.find(p => p.id === account.platform);
                if (!platform) return null;

                const Icon = platform.icon;

                return (
                  <div
                    key={account.id}
                    className="flex items-center justify-between p-4 bg-slate-700/50 rounded-lg border border-slate-600"
                  >
                    <div className="flex items-center space-x-4">
                      <div className={`p-2 rounded-lg bg-gradient-to-r ${platform.color}`}>
                        <Icon className="h-5 w-5 text-white" />
                      </div>
                      
                      <div className="flex items-center space-x-3">
                        {account.profileImage && (
                          <img
                            src={account.profileImage}
                            alt={account.displayName}
                            className="w-8 h-8 rounded-full"
                          />
                        )}
                        <div>
                          <div className="flex items-center space-x-2">
                            <span className="text-white font-medium">{account.displayName}</span>
                            {getStatusIcon(account)}
                          </div>
                          <div className="flex items-center space-x-2 text-sm text-slate-400">
                            <span>{account.username}</span>
                            <span>•</span>
                            <span>{getStatusText(account)}</span>
                          </div>
                          {account.errorMessage && (
                            <p className="text-xs text-red-400 mt-1">{account.errorMessage}</p>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center space-x-2">
                      {account.status === 'expired' || account.status === 'error' ? (
                        <button
                          onClick={() => handleRefresh(account.id)}
                          className="p-2 text-yellow-400 hover:text-yellow-300 transition-colors"
                          title="Refresh connection"
                        >
                          <RefreshCw className="h-4 w-4" />
                        </button>
                      ) : null}
                      
                      <button
                        onClick={() => setShowConfirmDisconnect(account.id)}
                        className="p-2 text-red-400 hover:text-red-300 transition-colors"
                        title="Disconnect account"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Available Platforms */}
        <div>
          <h3 className="text-lg font-medium text-white mb-4">Connect New Account</h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {platforms.map((platform) => {
              const isConnected = isAccountConnected(platform.id);
              const connecting = isConnecting(platform.id);
              const error = getConnectionError(platform.id);
              const Icon = platform.icon;

              return (
                <div key={platform.id} className="space-y-2">
                  <button
                    onClick={() => handleConnect(platform.id)}
                    disabled={connecting}
                    className={`w-full flex items-center space-x-3 p-4 rounded-lg border-2 transition-all ${
                      isConnected
                        ? 'border-green-500/50 bg-green-500/10 cursor-default'
                        : 'border-slate-600 hover:border-slate-500 bg-slate-700/50 hover:bg-slate-700'
                    } ${connecting ? 'opacity-50 cursor-not-allowed' : ''}`}
                  >
                    <div className={`p-2 rounded-lg bg-gradient-to-r ${platform.color}`}>
                      <Icon className="h-5 w-5 text-white" />
                    </div>
                    
                    <div className="flex-1 text-left">
                      <div className="text-white font-medium">{platform.name}</div>
                      <div className="text-sm text-slate-400">
                        {connecting ? 'Connecting...' : 
                         isConnected ? 'Connected' : 'Click to connect'}
                      </div>
                    </div>

                    {connecting && (
                      <div className="w-5 h-5 border-2 border-blue-400/30 border-t-blue-400 rounded-full animate-spin" />
                    )}

                    {isConnected && (
                      <CheckCircle className="h-5 w-5 text-green-400" />
                    )}

                    {!isConnected && !connecting && (
                      <Plus className="h-5 w-5 text-slate-400" />
                    )}
                  </button>

                  {error && (
                    <div className="text-sm text-red-400 px-4">
                      {error}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Security Notice */}
        <div className="p-4 bg-blue-900/20 border border-blue-500/30 rounded-lg">
          <div className="flex items-start space-x-3">
            <Shield className="h-5 w-5 text-blue-400 mt-0.5" />
            <div>
              <h4 className="text-blue-300 font-medium mb-1">Security & Privacy</h4>
              <p className="text-blue-200 text-sm leading-relaxed">
                Your social media credentials are encrypted and stored securely. We only request 
                the minimum permissions needed to post content. You can revoke access at any time.
              </p>
            </div>
          </div>
        </div>
      </div>

      {/* Disconnect Confirmation Modal */}
      {showConfirmDisconnect && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-800 rounded-xl p-6 w-full max-w-md border border-slate-700">
            <h3 className="text-lg font-semibold text-white mb-4">Disconnect Account</h3>
            <p className="text-slate-300 mb-6">
              Are you sure you want to disconnect this account? You'll need to reconnect it to post content.
            </p>
            <div className="flex space-x-3">
              <button
                onClick={() => handleDisconnect(showConfirmDisconnect)}
                className="flex-1 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition-colors"
              >
                Disconnect
              </button>
              <button
                onClick={() => setShowConfirmDisconnect(null)}
                className="flex-1 px-4 py-2 text-slate-300 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default SocialAccountManager;