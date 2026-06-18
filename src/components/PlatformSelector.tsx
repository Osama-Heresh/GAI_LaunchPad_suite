import React, { useState, useEffect } from 'react';
import { Twitter, Facebook, Instagram, Linkedin, MessageSquare, Hash, Briefcase, Settings, AlertCircle, Plus, Youtube } from 'lucide-react';
import { useSocialAccounts } from '../contexts/SocialAccountsContext';
import SocialAccountManager from './SocialAccountManager';

interface Platform {
  id: string;
  name: string;
  icon: React.ComponentType<any>;
  color: string;
  characterLimit: number;
}

interface PlatformSelectorProps {
  selectedPlatforms: string[];
  onPlatformToggle: (platformId: string) => void;
}

interface SocialCredential {
  id: string;
  platform: string;
  username: string;
  password: string;
  isActive: boolean;
  lastUpdated: string;
}

const PlatformSelector: React.FC<PlatformSelectorProps> = ({
  selectedPlatforms,
  onPlatformToggle
}) => {
  const { accounts, isAccountConnected, getActiveAccount } = useSocialAccounts();
  const [showAccountManager, setShowAccountManager] = useState(false);
  const [socialCredentials, setSocialCredentials] = useState<SocialCredential[]>([]);

  // Load credentials from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem('social_credentials');
      if (saved) {
        setSocialCredentials(JSON.parse(saved));
      }
    } catch (error) {
      console.error('Failed to load social credentials:', error);
    }
  }, []);

  const platforms: Platform[] = [
    {
      id: 'twitter',
      name: 'Twitter (X)',
      icon: Twitter,
      color: 'from-sky-500 to-blue-500',
      characterLimit: 280
    },
    {
      id: 'facebook',
      name: 'Facebook',
      icon: Facebook,
      color: 'from-blue-600 to-blue-500',
      characterLimit: 63206
    },
    {
      id: 'instagram',
      name: 'Instagram',
      icon: Instagram,
      color: 'from-pink-600 to-purple-500',
      characterLimit: 2200
    },
    {
      id: 'linkedin',
      name: 'LinkedIn',
      icon: Linkedin,
      color: 'from-blue-700 to-blue-600',
      characterLimit: 3000
    },
    {
      id: 'tiktok',
      name: 'TikTok',
      icon: Hash,
      color: 'from-pink-500 to-red-500',
      characterLimit: 150
    },
    {
      id: 'youtube',
      name: 'YouTube',
      icon: Youtube,
      color: 'from-red-600 to-red-500',
      characterLimit: 5000
    },
    {
      id: 'slack',
      name: 'Slack',
      icon: MessageSquare,
      color: 'from-purple-600 to-pink-600',
      characterLimit: 4000
    },
    {
      id: 'discord',
      name: 'Discord',
      icon: MessageSquare,
      color: 'from-indigo-600 to-purple-600',
      characterLimit: 2000
    }
  ];

  // Check if platform has credentials stored
  const hasCredentials = (platformId: string) => {
    const platformName = platforms.find(p => p.id === platformId)?.name.split(' ')[0]; // Get first word (Twitter, Facebook, etc.)
    return socialCredentials.some(cred => 
      cred.platform.toLowerCase() === platformName?.toLowerCase() && cred.isActive
    );
  };

  // Get credential info for platform
  const getCredentialInfo = (platformId: string) => {
    const platformName = platforms.find(p => p.id === platformId)?.name.split(' ')[0];
    return socialCredentials.find(cred => 
      cred.platform.toLowerCase() === platformName?.toLowerCase() && cred.isActive
    );
  };

  const handlePlatformToggle = (platformId: string) => {
    const hasStoredCredentials = hasCredentials(platformId);
    const connected = isAccountConnected(platformId);
    
    if (hasStoredCredentials || connected) {
      // If has credentials or connected, allow toggling selection
      onPlatformToggle(platformId);
    } else {
      // If no credentials, show account manager to add them
      setShowAccountManager(true);
    }
  };

  const connectedPlatforms = platforms.filter(p => hasCredentials(p.id) || isAccountConnected(p.id));
  const availablePlatforms = platforms.filter(p => !hasCredentials(p.id) && !isAccountConnected(p.id));

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-xl font-semibold text-white">Select Platforms to Post</h2>
        <button
          onClick={() => setShowAccountManager(true)}
          className="flex items-center space-x-2 px-3 py-2 text-slate-300 hover:text-white border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
        >
          <Settings className="h-4 w-4" />
          <span className="text-sm">Manage Accounts</span>
        </button>
      </div>

      {/* Connected Platforms */}
      {connectedPlatforms.length > 0 && (
        <div className="mb-6">
          <h3 className="text-sm font-medium text-slate-300 mb-3">Available Platforms</h3>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
            {connectedPlatforms.map((platform) => {
              const isSelected = selectedPlatforms.includes(platform.id);
              const activeAccount = getActiveAccount(platform.id);
              const credentialInfo = getCredentialInfo(platform.id);
              const Icon = platform.icon;
              
              return (
                <div key={platform.id} className="relative group">
                  <button
                    onClick={() => onPlatformToggle(platform.id)}
                    className={`w-full p-4 rounded-xl border-2 transition-all duration-200 ${
                      isSelected
                        ? 'border-blue-500 bg-blue-500/20 shadow-lg shadow-blue-500/25'
                        : 'border-slate-600 hover:border-slate-500 bg-slate-700/50 hover:bg-slate-700'
                    }`}
                    aria-pressed={isSelected}
                    title={`${platform.name} - ${platform.characterLimit} character limit`}
                  >
                    <div className="flex flex-col items-center space-y-2">
                      <div className={`p-3 rounded-lg bg-gradient-to-r ${platform.color}`}>
                        <Icon className="h-6 w-6 text-white" />
                      </div>
                      
                      <div className="text-center">
                        <span className="text-sm font-medium text-white">
                          {platform.name}
                        </span>
                        <div className="text-xs text-slate-400 mt-1">
                          {platform.characterLimit > 1000 
                            ? `${Math.round(platform.characterLimit / 1000)}k chars`
                            : `${platform.characterLimit} chars`
                          }
                        </div>
                        
                        {/* Show account info */}
                        {activeAccount && (
                          <div className="flex items-center justify-center space-x-1 mt-1">
                            {activeAccount.profileImage && (
                              <img
                                src={activeAccount.profileImage}
                                alt={activeAccount.displayName}
                                className="w-4 h-4 rounded-full"
                              />
                            )}
                            <span className="text-xs text-slate-400 truncate max-w-20">
                              {activeAccount.username}
                            </span>
                          </div>
                        )}
                        
                        {/* Show credential info if no OAuth account */}
                        {!activeAccount && credentialInfo && (
                          <div className="flex items-center justify-center space-x-1 mt-1">
                            <span className="text-xs text-green-400 truncate max-w-20">
                              {credentialInfo.username}
                            </span>
                          </div>
                        )}
                      </div>
                    </div>
                    
                    {/* Selection indicator */}
                    {isSelected && (
                      <div className="absolute -top-2 -right-2 w-6 h-6 bg-blue-500 rounded-full flex items-center justify-center">
                        <svg className="w-4 h-4 text-white" fill="currentColor" viewBox="0 0 20 20">
                          <path fillRule="evenodd" d="M16.707 5.293a1 1 0 010 1.414l-8 8a1 1 0 01-1.414 0l-4-4a1 1 0 011.414-1.414L8 12.586l7.293-7.293a1 1 0 011.414 0z" clipRule="evenodd" />
                        </svg>
                      </div>
                    )}
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Available Platforms to Connect */}
      {availablePlatforms.length > 0 && (
        <div className="mb-4">
          <h3 className="text-sm font-medium text-slate-300 mb-3">Connect More Platforms</h3>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
            {availablePlatforms.map((platform) => {
              const Icon = platform.icon;
              
              return (
                <div key={platform.id} className="relative group">
                  <button
                    onClick={() => setShowAccountManager(true)}
                    className="w-full p-4 rounded-xl border-2 border-dashed border-slate-600 hover:border-slate-500 bg-slate-700/30 hover:bg-slate-700/50 transition-all duration-200 opacity-70 hover:opacity-100"
                    title={`Connect your ${platform.name} account`}
                  >
                    <div className="flex flex-col items-center space-y-2">
                      <div className={`p-3 rounded-lg bg-gradient-to-r ${platform.color} opacity-60`}>
                        <Icon className="h-6 w-6 text-white" />
                      </div>
                      
                      <div className="text-center">
                        <span className="text-sm font-medium text-slate-400">
                          {platform.name}
                        </span>
                        <div className="text-xs text-slate-500 mt-1">
                          Click to connect
                        </div>
                      </div>
                    </div>
                    
                    {/* Plus icon indicator */}
                    <div className="absolute -top-2 -right-2 w-6 h-6 bg-slate-600 rounded-full flex items-center justify-center">
                      <Plus className="w-4 h-4 text-slate-300" />
                    </div>
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* No platforms connected state */}
      {connectedPlatforms.length === 0 && (
        <div className="text-center py-8">
          <div className="w-16 h-16 bg-slate-700/50 rounded-full flex items-center justify-center mx-auto mb-4">
            <Briefcase className="h-8 w-8 text-slate-400" />
          </div>
          <h3 className="text-lg font-medium text-white mb-2">No Connected Platforms</h3>
          <p className="text-slate-400 mb-4">
            Connect your social media accounts to start posting
          </p>
          <button
            onClick={() => setShowAccountManager(true)}
            className="inline-flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <Plus className="h-4 w-4" />
            <span>Connect Platforms</span>
          </button>
        </div>
      )}
      
      {/* Selected platforms summary */}
      {selectedPlatforms.length > 0 && (
        <div className="mt-4 p-4 bg-slate-700/50 rounded-lg">
          <p className="text-sm text-slate-300">
            Selected for posting: <span className="text-white font-medium">
              {selectedPlatforms.length} platform{selectedPlatforms.length > 1 ? 's' : ''}
            </span>
          </p>
          <div className="flex flex-wrap gap-2 mt-2">
            {selectedPlatforms.map((platformId) => {
              const platform = platforms.find(p => p.id === platformId);
              const activeAccount = getActiveAccount(platformId);
              const credentialInfo = getCredentialInfo(platformId);
              if (!platform) return null;
              
              const displayName = activeAccount?.username || credentialInfo?.username || 'Connected';
              
              return (
                <div
                  key={platformId}
                  className="flex items-center space-x-2 px-3 py-1 bg-blue-500/20 text-blue-300 text-xs rounded-md border border-blue-500/30"
                >
                  <span>{platform.name}</span>
                  <span className="text-blue-400">({displayName})</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Helpful tip */}
      {connectedPlatforms.length > 0 && selectedPlatforms.length === 0 && (
        <div className="mt-4 p-3 bg-orange-900/20 border border-orange-500/30 rounded-lg">
          <div className="flex items-center space-x-2">
            <AlertCircle className="h-4 w-4 text-orange-400" />
            <span className="text-orange-300 text-sm">
              Select at least one platform to publish your post
            </span>
          </div>
        </div>
      )}

      {/* Account Manager Modal */}
      {showAccountManager && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-4xl max-h-[90vh] overflow-hidden">
            <SocialAccountManager onClose={() => setShowAccountManager(false)} />
          </div>
        </div>
      )}
    </div>
  );
};

export default PlatformSelector;