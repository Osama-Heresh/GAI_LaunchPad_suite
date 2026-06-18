import React, { useState, useRef } from 'react';
import { Upload, Image, Video, Type, Sparkles, X, Settings, Palette, AlertTriangle, ExternalLink, CheckCircle, Clock } from 'lucide-react';
import { aiService, AIGenerationOptions } from '../services/aiService';

interface MediaUploadProps {
  activeTab: 'image' | 'video' | 'text';
  onTabChange: (tab: 'image' | 'video' | 'text') => void;
  onMediaUpload: (file: File) => void;
  onAIGenerate: (prompt: string, type: 'image' | 'video') => void;
  uploadedMedia: string | null;
  onRemoveMedia: () => void;
}

const MediaUpload: React.FC<MediaUploadProps> = ({
  activeTab,
  onTabChange,
  onMediaUpload,
  onAIGenerate,
  uploadedMedia,
  onRemoveMedia
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const [showAIPanel, setShowAIPanel] = useState(false);
  const [aiPrompt, setAiPrompt] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [aiOptions, setAiOptions] = useState<Partial<AIGenerationOptions>>({
    style: 'photorealistic',
    size: '1024x1024',
    quality: 'standard',
    model: 'dall-e-3'
  });
  const [showAdvancedOptions, setShowAdvancedOptions] = useState(false);
  const [generationError, setGenerationError] = useState<string | null>(null);
  const [generationProgress, setGenerationProgress] = useState<string>('');
  const [apiAvailability, setApiAvailability] = useState<any>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Check API availability when AI panel opens
  React.useEffect(() => {
    if (showAIPanel) {
      aiService.checkAPIAvailability().then(setApiAvailability);
    }
  }, [showAIPanel]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const files = Array.from(e.dataTransfer.files);
    const file = files[0];
    
    if (file && (
      (activeTab === 'image' && file.type.startsWith('image/')) ||
      (activeTab === 'video' && file.type.startsWith('video/'))
    )) {
      onMediaUpload(file);
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onMediaUpload(file);
    }
  };

  const handleAIGenerate = async () => {
    if (!aiPrompt.trim()) {
      setGenerationError('Please enter a prompt to generate content.');
      return;
    }
    
    setIsGenerating(true);
    setGenerationError(null);
    setGenerationProgress('Initializing...');
    
    try {
      console.log(`Starting ${activeTab} generation with prompt:`, aiPrompt);
      
      // Update progress for video generation
      if (activeTab === 'video') {
        setGenerationProgress('Preparing video generation...');
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        setGenerationProgress('Processing with AI models...');
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        setGenerationProgress('Finalizing video...');
      } else {
        setGenerationProgress('Generating image...');
      }
      
      // Optimize the prompt
      const optimizedPrompt = aiService.optimizePrompt(
        aiPrompt, 
        activeTab as 'image' | 'video', 
        aiOptions.style
      );

      const options: AIGenerationOptions = {
        prompt: optimizedPrompt,
        type: activeTab as 'image' | 'video',
        ...aiOptions
      };

      console.log('Generation options:', options);

      const result = await aiService.generateContent(options);
      
      console.log('Generation result:', result);
      
      if (result.success && result.url) {
        setGenerationProgress('Processing generated content...');
        
        try {
          // For video URLs, we can use them directly
          if (activeTab === 'video') {
            // Test if the video URL is accessible
            const testVideo = document.createElement('video');
            testVideo.preload = 'metadata';
            
            const videoLoadPromise = new Promise((resolve, reject) => {
              testVideo.onloadedmetadata = () => {
                console.log('Video metadata loaded successfully:', result.url);
                resolve(result.url);
              };
              
              testVideo.onerror = (error) => {
                console.warn('Video loading failed, will use direct URL:', error);
                resolve(result.url); // Still resolve, we'll handle it
              };
              
              // Set timeout for video loading
              setTimeout(() => {
                console.log('Video loading timeout, proceeding with URL');
                resolve(result.url);
              }, 5000);
            });
            
            testVideo.src = result.url;
            
            await videoLoadPromise;
            
            // Create a blob URL for better handling
            try {
              const response = await fetch(result.url);
              if (response.ok) {
                const blob = await response.blob();
                const file = new File([blob], `ai-generated-video.mp4`, {
                  type: 'video/mp4'
                });
                onMediaUpload(file);
              } else {
                // If fetch fails, use the direct URL
                console.log('Using direct video URL');
                onAIGenerate(aiPrompt, 'video');
              }
            } catch (fetchError) {
              console.warn('Fetch failed, using direct URL:', fetchError);
              onAIGenerate(aiPrompt, 'video');
            }
          } else {
            // For images, create a blob from the URL
            const response = await fetch(result.url);
            const blob = await response.blob();
            const file = new File([blob], `ai-generated-image.${result.url.includes('.png') ? 'png' : 'jpg'}`, {
              type: blob.type || 'image/png'
            });
            
            onMediaUpload(file);
          }
          
          setShowAIPanel(false);
          setAiPrompt('');
          setGenerationError(null);
          setGenerationProgress('');
          
          console.log(`${activeTab} generation completed successfully`);
        } catch (fileError) {
          console.error('Error processing generated content:', fileError);
          setGenerationError(`Generated ${activeTab} successfully but failed to process the file. Please try again.`);
        }
      } else {
        console.error('Generation failed:', result.error);
        setGenerationError(result.error || `Failed to generate ${activeTab}. Please try again.`);
      }
    } catch (error: any) {
      console.error('AI generation failed:', error);
      setGenerationError(error.message || `An unexpected error occurred while generating ${activeTab}. Please try again.`);
    } finally {
      setIsGenerating(false);
      setGenerationProgress('');
    }
  };

  const tabs = [
    { id: 'image', label: 'Image', icon: Image },
    { id: 'video', label: 'Video', icon: Video },
    { id: 'text', label: 'Text Only', icon: Type }
  ];

  const imageStyles = [
    { value: 'photorealistic', label: 'Photorealistic' },
    { value: 'artistic', label: 'Artistic' },
    { value: 'minimalist', label: 'Minimalist' },
    { value: 'vintage', label: 'Vintage' },
    { value: 'futuristic', label: 'Futuristic' }
  ];

  const imageSizes = [
    { value: '1024x1024', label: 'Square (1024×1024)' },
    { value: '1792x1024', label: 'Landscape (1792×1024)' },
    { value: '1024x1792', label: 'Portrait (1024×1792)' }
  ];

  const imageQualities = [
    { value: 'standard', label: 'Standard' },
    { value: 'hd', label: 'HD' }
  ];

  const imageModels = [
    { value: 'dall-e-3', label: 'DALL-E 3 (Latest)' },
    { value: 'dall-e-2', label: 'DALL-E 2' }
  ];

  const getAvailableServices = () => {
    if (!apiAvailability) return [];
    
    const services = [];
    if (activeTab === 'image') {
      if (apiAvailability.openai.available) {
        services.push({
          name: 'OpenAI DALL-E',
          status: apiAvailability.openai.isTrialMode ? 'Trial' : 'Personal Key',
          color: apiAvailability.openai.isTrialMode ? 'text-orange-400' : 'text-green-400'
        });
      }
      if (apiAvailability.stability.available) {
        services.push({
          name: 'Stability AI',
          status: apiAvailability.stability.isTrialMode ? 'Trial' : 'Personal Key',
          color: apiAvailability.stability.isTrialMode ? 'text-orange-400' : 'text-green-400'
        });
      }
    } else if (activeTab === 'video') {
      if (apiAvailability.replicate.available) {
        services.push({
          name: 'Replicate',
          status: apiAvailability.replicate.isTrialMode ? 'Demo Mode' : 'Personal Key',
          color: apiAvailability.replicate.isTrialMode ? 'text-blue-400' : 'text-green-400'
        });
      }
      if (apiAvailability.runway.available) {
        services.push({
          name: 'Runway ML',
          status: apiAvailability.runway.isTrialMode ? 'Demo Mode' : 'Personal Key',
          color: apiAvailability.runway.isTrialMode ? 'text-blue-400' : 'text-green-400'
        });
      }
    }
    
    return services;
  };

  const hasAvailableServices = () => {
    // For video, we always have demo mode available
    if (activeTab === 'video') return true;
    
    const services = getAvailableServices();
    return services.length > 0;
  };

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700">
      <h2 className="text-xl font-semibold text-white mb-4">Media Content</h2>
      
      {/* Tab Navigation */}
      <div className="flex space-x-1 mb-6 bg-slate-700/50 rounded-lg p-1">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            onClick={() => onTabChange(tab.id as any)}
            className={`flex items-center space-x-2 px-4 py-2 rounded-md text-sm font-medium transition-all ${
              activeTab === tab.id
                ? 'bg-blue-600 text-white shadow-lg'
                : 'text-slate-300 hover:text-white hover:bg-slate-600/50'
            }`}
            aria-pressed={activeTab === tab.id}
          >
            <tab.icon className="h-4 w-4" />
            <span>{tab.label}</span>
          </button>
        ))}
      </div>

      {/* Content Area */}
      {activeTab !== 'text' && (
        <div className="space-y-4">
          {!uploadedMedia ? (
            <>
              {/* Upload Area */}
              <div
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                className={`border-2 border-dashed rounded-xl p-8 text-center transition-all ${
                  isDragging
                    ? 'border-blue-500 bg-blue-500/10'
                    : 'border-slate-600 hover:border-slate-500'
                }`}
              >
                <Upload className="h-12 w-12 text-slate-400 mx-auto mb-4" />
                <p className="text-slate-300 mb-2">
                  Drag and drop your {activeTab} here, or
                </p>
                <button
                  onClick={() => fileInputRef.current?.click()}
                  className="text-blue-400 hover:text-blue-300 font-medium"
                >
                  browse files
                </button>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept={activeTab === 'image' ? 'image/*' : 'video/*'}
                  onChange={handleFileSelect}
                  className="hidden"
                />
              </div>

              {/* AI Generation Button */}
              <div className="text-center">
                <button
                  onClick={() => setShowAIPanel(true)}
                  className="inline-flex items-center space-x-2 px-6 py-3 bg-gradient-to-r from-purple-600 to-pink-600 text-white rounded-lg hover:from-purple-500 hover:to-pink-500 transition-all shadow-lg hover:shadow-xl"
                >
                  <Sparkles className="h-5 w-5" />
                  <span>Generate with AI</span>
                </button>
                <p className="text-xs text-slate-400 mt-2">
                  {activeTab === 'video' 
                    ? 'Powered by Replicate & Runway ML (Demo mode available)'
                    : 'Powered by OpenAI DALL-E & Stability AI'
                  }
                </p>
              </div>
            </>
          ) : (
            /* Media Preview */
            <div className="relative">
              {activeTab === 'image' ? (
                <img
                  src={uploadedMedia}
                  alt="Uploaded content"
                  className="w-full h-64 object-cover rounded-lg"
                />
              ) : (
                <video
                  src={uploadedMedia}
                  controls
                  className="w-full h-64 rounded-lg"
                />
              )}
              <button
                onClick={onRemoveMedia}
                className="absolute top-2 right-2 p-2 bg-red-600 text-white rounded-full hover:bg-red-700 transition-colors"
                aria-label="Remove media"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
          )}
        </div>
      )}

      {activeTab === 'text' && (
        <div className="text-center py-12">
          <Type className="h-16 w-16 text-slate-400 mx-auto mb-4" />
          <p className="text-slate-300">Text-only post selected</p>
          <p className="text-slate-500 text-sm mt-2">
            Add your content in the post editor below
          </p>
        </div>
      )}

      {/* Enhanced AI Generation Panel */}
      {showAIPanel && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-800 rounded-xl p-6 w-full max-w-2xl border border-slate-700 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-lg font-semibold text-white flex items-center space-x-2">
                <Sparkles className="h-5 w-5 text-purple-400" />
                <span>Generate {activeTab} with AI</span>
              </h3>
              <button
                onClick={() => {
                  setShowAIPanel(false);
                  setGenerationError(null);
                  setGenerationProgress('');
                }}
                className="text-slate-400 hover:text-white"
                aria-label="Close panel"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            
            <div className="space-y-6">
              {/* API Status */}
              {apiAvailability && (
                <div className="space-y-3">
                  <h4 className="text-sm font-medium text-slate-300">Available Services</h4>
                  {hasAvailableServices() ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                      {getAvailableServices().map((service, index) => (
                        <div key={index} className="flex items-center justify-between p-2 bg-slate-700/50 rounded-lg">
                          <span className="text-white text-sm">{service.name}</span>
                          <span className={`text-xs px-2 py-1 rounded-full ${
                            service.status === 'Trial' || service.status === 'Demo Mode'
                              ? 'bg-blue-500/20 text-blue-400' 
                              : 'bg-green-500/20 text-green-400'
                          }`}>
                            {service.status}
                          </span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="p-4 bg-red-900/20 border border-red-500/30 rounded-lg">
                      <div className="flex items-start space-x-3">
                        <AlertTriangle className="h-5 w-5 text-red-400 mt-0.5" />
                        <div>
                          <h4 className="text-red-300 font-medium mb-1">No API Keys Available</h4>
                          <p className="text-red-200 text-sm leading-relaxed mb-3">
                            You need to add AI service API keys to generate {activeTab}s. 
                            Go to Settings → Security to add your API keys.
                          </p>
                          <a
                            href="/settings"
                            className="inline-flex items-center space-x-1 text-red-400 hover:text-red-300 text-sm"
                          >
                            <span>Go to Settings</span>
                            <ExternalLink className="h-3 w-3" />
                          </a>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* Demo Mode Notice for Videos */}
              {activeTab === 'video' && (
                <div className="p-4 bg-blue-900/20 border border-blue-500/30 rounded-lg">
                  <div className="flex items-start space-x-3">
                    <CheckCircle className="h-5 w-5 text-blue-400 mt-0.5" />
                    <div>
                      <h4 className="text-blue-300 font-medium mb-1">Video Generation Available</h4>
                      <p className="text-blue-200 text-sm leading-relaxed">
                        Video generation works in demo mode with high-quality sample videos. Add your own API keys for real AI-generated videos.
                      </p>
                    </div>
                  </div>
                </div>
              )}

              {/* Generation Progress */}
              {isGenerating && generationProgress && (
                <div className="p-4 bg-blue-900/20 border border-blue-500/30 rounded-lg">
                  <div className="flex items-center space-x-3">
                    <div className="w-5 h-5 border-2 border-blue-400/30 border-t-blue-400 rounded-full animate-spin" />
                    <div>
                      <h4 className="text-blue-300 font-medium">Generating {activeTab}...</h4>
                      <p className="text-blue-200 text-sm">{generationProgress}</p>
                    </div>
                  </div>
                </div>
              )}

              {/* Prompt Input */}
              <div>
                <label className="block text-sm font-medium text-slate-300 mb-2">
                  Describe what you want to create
                </label>
                <textarea
                  value={aiPrompt}
                  onChange={(e) => setAiPrompt(e.target.value)}
                  placeholder={`e.g., "A modern office workspace with natural lighting and plants" for ${activeTab}`}
                  rows={4}
                  className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
                />
              </div>

              {/* Style Selection */}
              {activeTab === 'image' && (
                <div>
                  <label className="block text-sm font-medium text-slate-300 mb-2">
                    Style
                  </label>
                  <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
                    {imageStyles.map((style) => (
                      <button
                        key={style.value}
                        onClick={() => setAiOptions(prev => ({ ...prev, style: style.value }))}
                        className={`p-3 rounded-lg border text-sm transition-all ${
                          aiOptions.style === style.value
                            ? 'border-purple-500 bg-purple-500/20 text-purple-300'
                            : 'border-slate-600 text-slate-300 hover:border-slate-500'
                        }`}
                      >
                        {style.label}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* Advanced Options Toggle */}
              <button
                onClick={() => setShowAdvancedOptions(!showAdvancedOptions)}
                className="flex items-center space-x-2 text-slate-300 hover:text-white transition-colors"
              >
                <Settings className="h-4 w-4" />
                <span>Advanced Options</span>
              </button>

              {/* Advanced Options */}
              {showAdvancedOptions && activeTab === 'image' && (
                <div className="space-y-4 p-4 bg-slate-700/30 rounded-lg border border-slate-600">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-sm font-medium text-slate-300 mb-2">
                        Size
                      </label>
                      <select
                        value={aiOptions.size}
                        onChange={(e) => setAiOptions(prev => ({ ...prev, size: e.target.value as any }))}
                        className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        {imageSizes.map((size) => (
                          <option key={size.value} value={size.value}>
                            {size.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-slate-300 mb-2">
                        Quality
                      </label>
                      <select
                        value={aiOptions.quality}
                        onChange={(e) => setAiOptions(prev => ({ ...prev, quality: e.target.value as any }))}
                        className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        {imageQualities.map((quality) => (
                          <option key={quality.value} value={quality.value}>
                            {quality.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div className="md:col-span-2">
                      <label className="block text-sm font-medium text-slate-300 mb-2">
                        Model
                      </label>
                      <select
                        value={aiOptions.model}
                        onChange={(e) => setAiOptions(prev => ({ ...prev, model: e.target.value }))}
                        className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        {imageModels.map((model) => (
                          <option key={model.value} value={model.value}>
                            {model.label}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              )}

              {/* Error Display */}
              {generationError && (
                <div className="p-4 bg-red-900/50 border border-red-500/50 rounded-lg">
                  <div className="flex items-start space-x-3">
                    <AlertTriangle className="h-5 w-5 text-red-400 mt-0.5" />
                    <div>
                      <h4 className="text-red-300 font-medium mb-1">Generation Error</h4>
                      <p className="text-red-200 text-sm">{generationError}</p>
                    </div>
                  </div>
                </div>
              )}
              
              {/* Action Buttons */}
              <div className="flex space-x-3">
                <button
                  onClick={handleAIGenerate}
                  disabled={!aiPrompt.trim() || isGenerating}
                  className="flex-1 flex items-center justify-center space-x-2 px-4 py-3 bg-gradient-to-r from-purple-600 to-pink-600 text-white rounded-lg hover:from-purple-500 hover:to-pink-500 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                >
                  {isGenerating ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      <span>Generating...</span>
                    </>
                  ) : (
                    <>
                      <Sparkles className="h-4 w-4" />
                      <span>Generate {activeTab}</span>
                    </>
                  )}
                </button>
                <button
                  onClick={() => {
                    setShowAIPanel(false);
                    setGenerationError(null);
                    setGenerationProgress('');
                  }}
                  className="px-4 py-3 text-slate-300 border border-slate-600 rounded-lg hover:bg-slate-700 transition-colors"
                >
                  Cancel
                </button>
              </div>

              {/* API Status */}
              <div className="text-xs text-slate-400 text-center">
                <p>Using professional AI services for high-quality generation</p>
                <p className="mt-1">
                  {activeTab === 'video' 
                    ? 'Videos: Replicate Stable Video Diffusion • Runway ML (Demo mode available)'
                    : 'Images: OpenAI DALL-E 3 • Stability AI'
                  }
                </p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default MediaUpload;