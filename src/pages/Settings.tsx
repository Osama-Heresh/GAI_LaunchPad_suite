import React, { useState, useEffect } from 'react';
import { User, Bell, Shield, Palette, Globe, CreditCard, Eye, EyeOff, Plus, Edit, Trash2, Key, ExternalLink, Sparkles, AlertCircle } from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';
import { useAuth } from '../contexts/AuthContext';

interface SocialCredential {
  id: string;
  platform: string;
  username: string;
  password: string;
  isActive: boolean;
  lastUpdated: string;
}

interface AIApiKey {
  id: string;
  service: 'openai' | 'replicate' | 'stability' | 'runway';
  apiKey: string;
  isActive: boolean;
  lastUpdated: string;
}

const Settings: React.FC = () => {
  const { t, language, setLanguage } = useLanguage();
  const { user } = useAuth();
  const [activeTab, setActiveTab] = useState('profile');
  const [showPasswords, setShowPasswords] = useState<Record<string, boolean>>({});
  const [showApiKeys, setShowApiKeys] = useState<Record<string, boolean>>({});
  const [showAddCredential, setShowAddCredential] = useState(false);
  const [showAddApiKey, setShowAddApiKey] = useState(false);
  const [editingCredential, setEditingCredential] = useState<SocialCredential | null>(null);
  const [editingApiKey, setEditingApiKey] = useState<AIApiKey | null>(null);

  const tabs = [
    { id: 'profile', name: 'Profile', icon: User },
    { id: 'notifications', name: 'Notifications', icon: Bell },
    { id: 'security', name: 'Security', icon: Shield },
    { id: 'appearance', name: 'Appearance', icon: Palette },
    { id: 'language', name: 'Language', icon: Globe },
    { id: 'billing', name: 'Billing', icon: CreditCard },
  ];

  const [profile, setProfile] = useState({
    name: user?.name || '',
    email: user?.email || '',
    company: user?.company || '',
    phone: '',
    timezone: 'UTC',
  });

  const [notifications, setNotifications] = useState({
    email: true,
    push: true,
    marketing: false,
    updates: true,
  });

  // Load credentials from localStorage
  const [socialCredentials, setSocialCredentials] = useState<SocialCredential[]>(() => {
    try {
      const saved = localStorage.getItem('social_credentials');
      return saved ? JSON.parse(saved) : [
        {
          id: '1',
          platform: 'Twitter',
          username: 'john_doe_business',
          password: 'twitter_password_123',
          isActive: true,
          lastUpdated: '2024-01-10'
        },
        {
          id: '2',
          platform: 'Facebook',
          username: 'john.doe.business@email.com',
          password: 'facebook_secure_pass',
          isActive: true,
          lastUpdated: '2024-01-08'
        },
        {
          id: '3',
          platform: 'LinkedIn',
          username: 'john.doe.professional',
          password: 'linkedin_password_456',
          isActive: false,
          lastUpdated: '2023-12-15'
        }
      ];
    } catch {
      return [];
    }
  });

  // Load AI API keys from localStorage
  const [aiApiKeys, setAiApiKeys] = useState<AIApiKey[]>(() => {
    try {
      const saved = localStorage.getItem('ai_api_keys');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [newCredential, setNewCredential] = useState({
    platform: '',
    username: '',
    password: ''
  });

  const [newApiKey, setNewApiKey] = useState({
    service: 'openai' as const,
    apiKey: ''
  });

  const platforms = [
    'Twitter', 'Facebook', 'Instagram', 'LinkedIn', 'TikTok', 'YouTube', 'Pinterest', 'Snapchat'
  ];

  const aiServices = [
    {
      id: 'openai',
      name: 'OpenAI (DALL-E)',
      description: 'Generate high-quality images with DALL-E 3',
      getKeyUrl: 'https://platform.openai.com/api-keys',
      placeholder: 'sk-...',
      icon: '🤖'
    },
    {
      id: 'replicate',
      name: 'Replicate',
      description: 'Generate videos and run ML models',
      getKeyUrl: 'https://replicate.com/account/api-tokens',
      placeholder: 'r8_...',
      icon: '🔄'
    },
    {
      id: 'stability',
      name: 'Stability AI',
      description: 'Generate images with Stable Diffusion',
      getKeyUrl: 'https://platform.stability.ai/account/keys',
      placeholder: 'sk-...',
      icon: '🎨'
    },
    {
      id: 'runway',
      name: 'Runway ML',
      description: 'Generate videos with Gen-2',
      getKeyUrl: 'https://runwayml.com/account',
      placeholder: 'rw_...',
      icon: '🎬'
    }
  ];

  // Save credentials to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem('social_credentials', JSON.stringify(socialCredentials));
  }, [socialCredentials]);

  // Save AI API keys to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem('ai_api_keys', JSON.stringify(aiApiKeys));
  }, [aiApiKeys]);

  const togglePasswordVisibility = (credentialId: string) => {
    setShowPasswords(prev => ({
      ...prev,
      [credentialId]: !prev[credentialId]
    }));
  };

  const toggleApiKeyVisibility = (keyId: string) => {
    setShowApiKeys(prev => ({
      ...prev,
      [keyId]: !prev[keyId]
    }));
  };

  const handleAddCredential = () => {
    if (!newCredential.platform || !newCredential.username || !newCredential.password) return;

    const credential: SocialCredential = {
      id: Date.now().toString(),
      platform: newCredential.platform,
      username: newCredential.username,
      password: newCredential.password,
      isActive: true,
      lastUpdated: new Date().toISOString().split('T')[0]
    };

    setSocialCredentials(prev => [...prev, credential]);
    setNewCredential({ platform: '', username: '', password: '' });
    setShowAddCredential(false);
  };

  const handleAddApiKey = () => {
    if (!newApiKey.service || !newApiKey.apiKey) return;

    // Check if key already exists for this service
    const existingKey = aiApiKeys.find(key => key.service === newApiKey.service);
    if (existingKey) {
      // Update existing key
      setAiApiKeys(prev =>
        prev.map(key =>
          key.service === newApiKey.service
            ? { ...key, apiKey: newApiKey.apiKey, lastUpdated: new Date().toISOString().split('T')[0] }
            : key
        )
      );
    } else {
      // Add new key
      const apiKey: AIApiKey = {
        id: Date.now().toString(),
        service: newApiKey.service,
        apiKey: newApiKey.apiKey,
        isActive: true,
        lastUpdated: new Date().toISOString().split('T')[0]
      };
      setAiApiKeys(prev => [...prev, apiKey]);
    }

    setNewApiKey({ service: 'openai', apiKey: '' });
    setShowAddApiKey(false);
  };

  const handleEditCredential = (credential: SocialCredential) => {
    setEditingCredential(credential);
    setNewCredential({
      platform: credential.platform,
      username: credential.username,
      password: credential.password
    });
  };

  const handleEditApiKey = (apiKey: AIApiKey) => {
    setEditingApiKey(apiKey);
    setNewApiKey({
      service: apiKey.service,
      apiKey: apiKey.apiKey
    });
  };

  const handleUpdateCredential = () => {
    if (!editingCredential || !newCredential.username || !newCredential.password) return;

    setSocialCredentials(prev =>
      prev.map(cred =>
        cred.id === editingCredential.id
          ? {
              ...cred,
              username: newCredential.username,
              password: newCredential.password,
              lastUpdated: new Date().toISOString().split('T')[0]
            }
          : cred
      )
    );

    setEditingCredential(null);
    setNewCredential({ platform: '', username: '', password: '' });
  };

  const handleUpdateApiKey = () => {
    if (!editingApiKey || !newApiKey.apiKey) return;

    setAiApiKeys(prev =>
      prev.map(key =>
        key.id === editingApiKey.id
          ? {
              ...key,
              apiKey: newApiKey.apiKey,
              lastUpdated: new Date().toISOString().split('T')[0]
            }
          : key
      )
    );

    setEditingApiKey(null);
    setNewApiKey({ service: 'openai', apiKey: '' });
  };

  const handleDeleteCredential = (credentialId: string) => {
    setSocialCredentials(prev => prev.filter(cred => cred.id !== credentialId));
  };

  const handleDeleteApiKey = (keyId: string) => {
    setAiApiKeys(prev => prev.filter(key => key.id !== keyId));
  };

  const toggleCredentialStatus = (credentialId: string) => {
    setSocialCredentials(prev =>
      prev.map(cred =>
        cred.id === credentialId
          ? { ...cred, isActive: !cred.isActive }
          : cred
      )
    );
  };

  const toggleApiKeyStatus = (keyId: string) => {
    setAiApiKeys(prev =>
      prev.map(key =>
        key.id === keyId
          ? { ...key, isActive: !key.isActive }
          : key
      )
    );
  };

  const getPlatformColor = (platform: string) => {
    const colors: Record<string, string> = {
      'Twitter': 'text-blue-400',
      'Facebook': 'text-blue-600',
      'Instagram': 'text-pink-500',
      'LinkedIn': 'text-blue-700',
      'TikTok': 'text-red-500',
      'YouTube': 'text-red-600',
      'Pinterest': 'text-red-400',
      'Snapchat': 'text-yellow-400'
    };
    return colors[platform] || 'text-gray-500';
  };

  const getServiceColor = (service: string) => {
    const colors: Record<string, string> = {
      'openai': 'text-green-500',
      'replicate': 'text-blue-500',
      'stability': 'text-purple-500',
      'runway': 'text-orange-500'
    };
    return colors[service] || 'text-gray-500';
  };

  const getApiKeyForService = (service: string) => {
    return aiApiKeys.find(key => key.service === service && key.isActive);
  };

  return (
    <div className="space-y-6">
      <h1 className="text-3xl font-bold gradient-text">{t('nav.settings')}</h1>

      <div className="glass-card rounded-xl shadow-sm overflow-hidden">
        <div className="flex border-b border-white/10">
          {tabs.map((tab) => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id)}
              className={`flex items-center space-x-2 px-6 py-4 text-sm font-medium border-b-2 transition-colors ${
                activeTab === tab.id
                  ? 'border-blue-500 text-blue-400'
                  : 'border-transparent text-white/60 hover:text-white/80'
              }`}
            >
              <tab.icon className="h-4 w-4" />
              <span>{tab.name}</span>
            </button>
          ))}
        </div>

        <div className="p-6">
          {activeTab === 'profile' && (
            <div className="space-y-6">
              <h2 className="text-lg font-semibold text-white">Profile Information</h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-2">
                    Full Name
                  </label>
                  <input
                    type="text"
                    value={profile.name}
                    onChange={(e) => setProfile(prev => ({ ...prev, name: e.target.value }))}
                    className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-2">
                    Email Address
                  </label>
                  <input
                    type="email"
                    value={profile.email}
                    onChange={(e) => setProfile(prev => ({ ...prev, email: e.target.value }))}
                    className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-2">
                    Company
                  </label>
                  <input
                    type="text"
                    value={profile.company}
                    onChange={(e) => setProfile(prev => ({ ...prev, company: e.target.value }))}
                    className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-2">
                    Phone Number
                  </label>
                  <input
                    type="tel"
                    value={profile.phone}
                    onChange={(e) => setProfile(prev => ({ ...prev, phone: e.target.value }))}
                    className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  />
                </div>
              </div>
              <button className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
                Save Changes
              </button>
            </div>
          )}

          {activeTab === 'notifications' && (
            <div className="space-y-6">
              <h2 className="text-lg font-semibold text-white">Notification Preferences</h2>
              <div className="space-y-4">
                {Object.entries(notifications).map(([key, value]) => (
                  <div key={key} className="flex items-center justify-between">
                    <div>
                      <h3 className="text-sm font-medium text-white capitalize">
                        {key} Notifications
                      </h3>
                      <p className="text-sm text-white/60">
                        Receive {key} notifications
                      </p>
                    </div>
                    <button
                      onClick={() => setNotifications(prev => ({ ...prev, [key]: !value }))}
                      className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${
                        value ? 'bg-blue-600' : 'bg-white/20'
                      }`}
                    >
                      <span
                        className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${
                          value ? 'translate-x-6' : 'translate-x-1'
                        }`}
                      />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {activeTab === 'security' && (
            <div className="space-y-8">
              {/* AI API Keys Section */}
              <div>
                <div className="flex items-center justify-between mb-6">
                  <div className="flex items-center space-x-3">
                    <Sparkles className="h-6 w-6 text-purple-400" />
                    <div>
                      <h2 className="text-lg font-semibold text-white">AI Service API Keys</h2>
                      <p className="text-sm text-white/60">Manage your personal AI service API keys for content generation</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setShowAddApiKey(true)}
                    className="flex items-center space-x-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Add API Key</span>
                  </button>
                </div>

                {/* Trial Notice */}
                <div className="glass-card p-4 bg-blue-500/10 border border-blue-500/20 mb-6">
                  <div className="flex items-start space-x-3">
                    <AlertCircle className="h-5 w-5 text-blue-400 mt-0.5" />
                    <div>
                      <h4 className="text-blue-300 font-medium mb-1">Trial Mode Active</h4>
                      <p className="text-blue-200 text-sm leading-relaxed">
                        You're currently using system API keys for AI generation. Add your own API keys to avoid usage limits and have full control over your AI generation costs.
                      </p>
                    </div>
                  </div>
                </div>

                {/* AI Services Grid */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                  {aiServices.map((service) => {
                    const userKey = getApiKeyForService(service.id);
                    const hasUserKey = !!userKey;
                    
                    return (
                      <div key={service.id} className="glass-card p-4 bg-white/5">
                        <div className="flex items-start justify-between mb-3">
                          <div className="flex items-center space-x-3">
                            <span className="text-2xl">{service.icon}</span>
                            <div>
                              <h3 className={`font-medium ${getServiceColor(service.id)}`}>
                                {service.name}
                              </h3>
                              <p className="text-white/60 text-sm">{service.description}</p>
                            </div>
                          </div>
                          <div className="flex items-center space-x-2">
                            <span className={`px-2 py-1 rounded-full text-xs ${
                              hasUserKey 
                                ? 'bg-green-500/20 text-green-400' 
                                : 'bg-orange-500/20 text-orange-400'
                            }`}>
                              {hasUserKey ? 'Personal Key' : 'Trial Mode'}
                            </span>
                          </div>
                        </div>

                        {hasUserKey ? (
                          <div className="space-y-3">
                            <div className="flex items-center space-x-2">
                              <span className="text-white/60 text-sm">API Key:</span>
                              <div className="flex items-center space-x-2 flex-1">
                                <span className="text-white font-mono text-sm flex-1">
                                  {showApiKeys[userKey.id] ? userKey.apiKey : '••••••••••••••••'}
                                </span>
                                <button
                                  onClick={() => toggleApiKeyVisibility(userKey.id)}
                                  className="text-white/60 hover:text-white"
                                >
                                  {showApiKeys[userKey.id] ? 
                                    <EyeOff className="h-4 w-4" /> : 
                                    <Eye className="h-4 w-4" />
                                  }
                                </button>
                              </div>
                            </div>
                            
                            <div className="flex items-center justify-between">
                              <span className="text-white/60 text-xs">
                                Updated: {userKey.lastUpdated}
                              </span>
                              <div className="flex items-center space-x-2">
                                <button
                                  onClick={() => toggleApiKeyStatus(userKey.id)}
                                  className={`px-3 py-1 rounded-full text-xs font-medium ${
                                    userKey.isActive 
                                      ? 'bg-green-500/20 text-green-400' 
                                      : 'bg-red-500/20 text-red-400'
                                  }`}
                                >
                                  {userKey.isActive ? 'Active' : 'Inactive'}
                                </button>
                                
                                <button
                                  onClick={() => handleEditApiKey(userKey)}
                                  className="p-1 text-white/60 hover:text-white"
                                  title="Edit"
                                >
                                  <Edit className="h-4 w-4" />
                                </button>
                                
                                <button
                                  onClick={() => handleDeleteApiKey(userKey.id)}
                                  className="p-1 text-red-400 hover:text-red-300"
                                  title="Delete"
                                >
                                  <Trash2 className="h-4 w-4" />
                                </button>
                              </div>
                            </div>
                          </div>
                        ) : (
                          <div className="space-y-3">
                            <p className="text-white/60 text-sm">
                              Add your personal API key to use this service without limits.
                            </p>
                            <div className="flex items-center space-x-2">
                              <a
                                href={service.getKeyUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="flex items-center space-x-1 text-blue-400 hover:text-blue-300 text-sm"
                              >
                                <span>Get API Key</span>
                                <ExternalLink className="h-3 w-3" />
                              </a>
                              <span className="text-white/40">•</span>
                              <button
                                onClick={() => {
                                  setNewApiKey({ service: service.id as any, apiKey: '' });
                                  setShowAddApiKey(true);
                                }}
                                className="text-purple-400 hover:text-purple-300 text-sm"
                              >
                                Add Key
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Account Security Section */}
              <div>
                <div className="flex items-center justify-between mb-4">
                  <h2 className="text-lg font-semibold text-white">Account Security & Social Media Credentials</h2>
                  <button
                    onClick={() => setShowAddCredential(true)}
                    className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
                  >
                    <Plus className="h-4 w-4" />
                    <span>Add Account</span>
                  </button>
                </div>

                {/* Password Security Section */}
                <div className="glass-card p-4 bg-white/5 mb-6">
                  <div className="flex items-center space-x-3 mb-4">
                    <Key className="h-5 w-5 text-yellow-400" />
                    <h3 className="text-md font-medium text-white">Account Security</h3>
                  </div>
                  <div className="space-y-3">
                    <button className="w-full text-left px-4 py-3 bg-white/10 rounded-lg hover:bg-white/15 transition-colors">
                      <div className="flex items-center justify-between">
                        <span className="text-white">Change Password</span>
                        <span className="text-white/60 text-sm">Last changed 30 days ago</span>
                      </div>
                    </button>
                    <button className="w-full text-left px-4 py-3 bg-white/10 rounded-lg hover:bg-white/15 transition-colors">
                      <div className="flex items-center justify-between">
                        <span className="text-white">Two-Factor Authentication</span>
                        <span className="text-green-400 text-sm">Enabled</span>
                      </div>
                    </button>
                  </div>
                </div>

                {/* Social Media Credentials */}
                <div className="space-y-4">
                  <h3 className="text-md font-medium text-white">Social Media Account Credentials</h3>
                  
                  {socialCredentials.length === 0 ? (
                    <div className="text-center py-8 glass-card bg-white/5">
                      <Shield className="h-12 w-12 text-white/40 mx-auto mb-3" />
                      <p className="text-white/60">No social media credentials stored</p>
                      <p className="text-white/40 text-sm mt-1">Add your social media account credentials to get started</p>
                    </div>
                  ) : (
                    <div className="space-y-3">
                      {socialCredentials.map((credential) => (
                        <div key={credential.id} className="glass-card p-4 bg-white/5">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center space-x-4">
                              <div className={`text-lg font-bold ${getPlatformColor(credential.platform)}`}>
                                {credential.platform}
                              </div>
                              <div>
                                <div className="text-white font-medium">{credential.username}</div>
                                <div className="text-white/60 text-sm">
                                  Last updated: {credential.lastUpdated}
                                </div>
                              </div>
                            </div>
                            
                            <div className="flex items-center space-x-3">
                              <div className="flex items-center space-x-2">
                                <span className="text-white/60 text-sm">Password:</span>
                                <div className="flex items-center space-x-2">
                                  <span className="text-white font-mono text-sm">
                                    {showPasswords[credential.id] ? credential.password : '••••••••••'}
                                  </span>
                                  <button
                                    onClick={() => togglePasswordVisibility(credential.id)}
                                    className="text-white/60 hover:text-white"
                                  >
                                    {showPasswords[credential.id] ? 
                                      <EyeOff className="h-4 w-4" /> : 
                                      <Eye className="h-4 w-4" />
                                    }
                                  </button>
                                </div>
                              </div>
                              
                              <div className="flex items-center space-x-2">
                                <button
                                  onClick={() => toggleCredentialStatus(credential.id)}
                                  className={`px-3 py-1 rounded-full text-xs font-medium ${
                                    credential.isActive 
                                      ? 'bg-green-500/20 text-green-400' 
                                      : 'bg-red-500/20 text-red-400'
                                  }`}
                                >
                                  {credential.isActive ? 'Active' : 'Inactive'}
                                </button>
                                
                                <button
                                  onClick={() => handleEditCredential(credential)}
                                  className="p-2 text-white/60 hover:text-white"
                                  title="Edit"
                                >
                                  <Edit className="h-4 w-4" />
                                </button>
                                
                                <button
                                  onClick={() => handleDeleteCredential(credential.id)}
                                  className="p-2 text-red-400 hover:text-red-300"
                                  title="Delete"
                                >
                                  <Trash2 className="h-4 w-4" />
                                </button>
                              </div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {activeTab === 'language' && (
            <div className="space-y-6">
              <h2 className="text-lg font-semibold text-white">Language Settings</h2>
              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-white/80 mb-2">
                    Interface Language
                  </label>
                  <select
                    value={language}
                    onChange={(e) => setLanguage(e.target.value as 'en' | 'ar')}
                    className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                  >
                    <option value="en" className="bg-slate-800 text-white">English</option>
                    <option value="ar" className="bg-slate-800 text-white">العربية (Arabic)</option>
                  </select>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'billing' && (
            <div className="space-y-6">
              <h2 className="text-lg font-semibold text-white">Billing & Subscription</h2>
              <div className="glass-card p-4 bg-blue-500/10 border border-blue-500/20">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-medium text-blue-300">Pro Plan</h3>
                    <p className="text-sm text-blue-200">$29/month - Active until Jan 31, 2024</p>
                  </div>
                  <button className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
                    Manage Subscription
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Add/Edit Credential Modal */}
      {(showAddCredential || editingCredential) && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="glass-card p-6 w-full max-w-md border border-white/20">
            <h2 className="text-lg font-semibold text-white mb-4">
              {editingCredential ? 'Edit Social Media Account' : 'Add Social Media Account'}
            </h2>
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Platform
                </label>
                <select
                  value={newCredential.platform}
                  onChange={(e) => setNewCredential(prev => ({ ...prev, platform: e.target.value }))}
                  disabled={!!editingCredential}
                  className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent disabled:opacity-50"
                >
                  <option value="" className="bg-slate-800 text-white">Select Platform</option>
                  {platforms.map(platform => (
                    <option key={platform} value={platform} className="bg-slate-800 text-white">{platform}</option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Username/Email
                </label>
                <input
                  type="text"
                  value={newCredential.username}
                  onChange={(e) => setNewCredential(prev => ({ ...prev, username: e.target.value }))}
                  placeholder="Enter your username or email"
                  className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white placeholder-white/40 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
              
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  Password
                </label>
                <input
                  type="password"
                  value={newCredential.password}
                  onChange={(e) => setNewCredential(prev => ({ ...prev, password: e.target.value }))}
                  placeholder="Enter your password"
                  className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white placeholder-white/40 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
            </div>
            
            <div className="flex space-x-3 mt-6">
              <button
                onClick={editingCredential ? handleUpdateCredential : handleAddCredential}
                className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
              >
                {editingCredential ? 'Update Account' : 'Add Account'}
              </button>
              <button
                onClick={() => {
                  setShowAddCredential(false);
                  setEditingCredential(null);
                  setNewCredential({ platform: '', username: '', password: '' });
                }}
                className="px-4 py-2 text-white/60 border border-white/20 rounded-lg hover:bg-white/10"
              >
                Cancel
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Add/Edit API Key Modal */}
      {(showAddApiKey || editingApiKey) && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="glass-card p-6 w-full max-w-md border border-white/20">
            <h2 className="text-lg font-semibold text-white mb-4">
              {editingApiKey ? 'Edit API Key' : 'Add AI Service API Key'}
            </h2>
            
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  AI Service
                </label>
                <select
                  value={newApiKey.service}
                  onChange={(e) => setNewApiKey(prev => ({ ...prev, service: e.target.value as any }))}
                  disabled={!!editingApiKey}
                  className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white focus:ring-2 focus:ring-blue-500 focus:border-transparent disabled:opacity-50"
                >
                  {aiServices.map(service => (
                    <option key={service.id} value={service.id} className="bg-slate-800 text-white">
                      {service.name}
                    </option>
                  ))}
                </select>
              </div>
              
              <div>
                <label className="block text-sm font-medium text-white/80 mb-2">
                  API Key
                </label>
                <input
                  type="password"
                  value={newApiKey.apiKey}
                  onChange={(e) => setNewApiKey(prev => ({ ...prev, apiKey: e.target.value }))}
                  placeholder={aiServices.find(s => s.id === newApiKey.service)?.placeholder || 'Enter your API key'}
                  className="w-full px-3 py-2 bg-white/10 border border-white/20 rounded-lg text-white placeholder-white/40 focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>

              <div className="glass-card p-3 bg-blue-500/10 border border-blue-500/20">
                <div className="flex items-start space-x-2">
                  <ExternalLink className="h-4 w-4 text-blue-400 mt-0.5" />
                  <div>
                    <p className="text-blue-300 text-sm font-medium">Get your API key:</p>
                    <a
                      href={aiServices.find(s => s.id === newApiKey.service)?.getKeyUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="text-blue-400 hover:text-blue-300 text-sm underline"
                    >
                      {aiServices.find(s => s.id === newApiKey.service)?.getKeyUrl}
                    </a>
                  </div>
                </div>
              </div>
            </div>
            
            <div className="flex space-x-3 mt-6">
              <button
                onClick={editingApiKey ? handleUpdateApiKey : handleAddApiKey}
                className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
              >
                {editingApiKey ? 'Update Key' : 'Add Key'}
              </button>
              <button
                onClick={() => {
                  setShowAddApiKey(false);
                  setEditingApiKey(null);
                  setNewApiKey({ service: 'openai', apiKey: '' });
                }}
                className="px-4 py-2 text-white/60 border border-white/20 rounded-lg hover:bg-white/10"
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

export default Settings;