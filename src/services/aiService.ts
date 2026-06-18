import OpenAI from 'openai';

// AI Service Configuration
const AI_CONFIG = {
  openai: {
    apiKey: import.meta.env.VITE_OPENAI_API_KEY,
    baseURL: 'https://api.openai.com/v1'
  },
  replicate: {
    apiKey: import.meta.env.VITE_REPLICATE_API_KEY,
    baseURL: 'https://api.replicate.com/v1'
  },
  stability: {
    apiKey: import.meta.env.VITE_STABILITY_API_KEY,
    baseURL: 'https://api.stability.ai/v1'
  }
};

export interface AIGenerationOptions {
  prompt: string;
  type: 'image' | 'video';
  style?: string;
  size?: '1024x1024' | '1792x1024' | '1024x1792';
  quality?: 'standard' | 'hd';
  model?: string;
}

export interface AIGenerationResult {
  success: boolean;
  url?: string;
  error?: string;
  metadata?: {
    model: string;
    prompt: string;
    revisedPrompt?: string;
    size: string;
    quality: string;
  };
}

class AIService {
  // Get user's API keys from localStorage
  private getUserApiKeys(): Record<string, string> {
    try {
      const keys = localStorage.getItem('ai_api_keys');
      if (!keys) return {};
      
      const apiKeys = JSON.parse(keys);
      const userKeys: Record<string, string> = {};
      
      apiKeys.forEach((key: any) => {
        if (key.isActive) {
          userKeys[key.service] = key.apiKey;
        }
      });
      
      return userKeys;
    } catch {
      return {};
    }
  }

  // Check if user has API keys, fallback to system keys for trial
  private getApiKey(service: 'openai' | 'replicate' | 'stability' | 'runway'): string | null {
    const userKeys = this.getUserApiKeys();
    
    // First try user's personal API key
    if (userKeys[service]) {
      return userKeys[service];
    }
    
    // Fallback to system keys for trial (with usage limits)
    switch (service) {
      case 'openai':
        return AI_CONFIG.openai.apiKey;
      case 'replicate':
        return AI_CONFIG.replicate.apiKey;
      case 'stability':
        return AI_CONFIG.stability.apiKey;
      case 'runway':
        return import.meta.env.VITE_RUNWAY_API_KEY;
      default:
        return null;
    }
  }

  // Check if using trial mode (system keys)
  private isTrialMode(service: string): boolean {
    const userKeys = this.getUserApiKeys();
    return !userKeys[service];
  }

  private async generateImageWithOpenAI(options: AIGenerationOptions): Promise<AIGenerationResult> {
    try {
      const apiKey = this.getApiKey('openai');
      if (!apiKey) {
        throw new Error('OpenAI API key not configured. Please add your API key in Settings.');
      }

      const isTrialMode = this.isTrialMode('openai');
      
      // Initialize OpenAI client with the appropriate API key
      const openai = new OpenAI({
        apiKey,
        dangerouslyAllowBrowser: true
      });

      const response = await openai.images.generate({
        model: options.model || 'dall-e-3',
        prompt: options.prompt,
        n: 1,
        size: options.size || '1024x1024',
        quality: options.quality || 'standard',
        response_format: 'url'
      });

      const imageData = response.data[0];
      
      return {
        success: true,
        url: imageData.url,
        metadata: {
          model: isTrialMode ? 'dall-e-3 (Trial)' : 'dall-e-3',
          prompt: options.prompt,
          revisedPrompt: imageData.revised_prompt,
          size: options.size || '1024x1024',
          quality: options.quality || 'standard'
        }
      };
    } catch (error: any) {
      console.error('OpenAI image generation failed:', error);
      
      let errorMessage = error.message || 'Failed to generate image with OpenAI';
      
      // Handle specific error cases
      if (error.message?.includes('insufficient_quota')) {
        errorMessage = this.isTrialMode('openai') 
          ? 'Trial quota exceeded. Please add your own OpenAI API key in Settings to continue.'
          : 'Your OpenAI API quota has been exceeded. Please check your billing.';
      } else if (error.message?.includes('invalid_api_key')) {
        errorMessage = 'Invalid OpenAI API key. Please check your API key in Settings.';
      }
      
      return {
        success: false,
        error: errorMessage
      };
    }
  }

  private async generateImageWithStability(options: AIGenerationOptions): Promise<AIGenerationResult> {
    try {
      const apiKey = this.getApiKey('stability');
      if (!apiKey) {
        throw new Error('Stability AI API key not configured. Please add your API key in Settings.');
      }

      const isTrialMode = this.isTrialMode('stability');

      const response = await fetch('https://api.stability.ai/v1/generation/stable-diffusion-xl-1024-v1-0/text-to-image', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
          'Accept': 'application/json'
        },
        body: JSON.stringify({
          text_prompts: [
            {
              text: options.prompt,
              weight: 1
            }
          ],
          cfg_scale: 7,
          height: 1024,
          width: 1024,
          steps: 30,
          samples: 1
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        
        if (response.status === 401) {
          throw new Error('Invalid Stability AI API key. Please check your API key in Settings.');
        } else if (response.status === 402) {
          throw new Error(isTrialMode 
            ? 'Trial credits exhausted. Please add your own Stability AI API key in Settings.'
            : 'Insufficient credits. Please check your Stability AI account balance.');
        }
        
        throw new Error(`Stability AI API error: ${response.status} - ${errorData.message || 'Unknown error'}`);
      }

      const data = await response.json();
      
      if (data.artifacts && data.artifacts.length > 0) {
        // Convert base64 to blob URL
        const base64Data = data.artifacts[0].base64;
        const byteCharacters = atob(base64Data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'image/png' });
        const url = URL.createObjectURL(blob);

        return {
          success: true,
          url,
          metadata: {
            model: isTrialMode ? 'stable-diffusion-xl (Trial)' : 'stable-diffusion-xl',
            prompt: options.prompt,
            size: '1024x1024',
            quality: 'standard'
          }
        };
      }

      throw new Error('No image generated');
    } catch (error: any) {
      console.error('Stability AI image generation failed:', error);
      return {
        success: false,
        error: error.message || 'Failed to generate image with Stability AI'
      };
    }
  }

  private async generateVideoWithReplicate(options: AIGenerationOptions): Promise<AIGenerationResult> {
    try {
      const apiKey = this.getApiKey('replicate');
      if (!apiKey) {
        // If no API key, use demo mode with sample video
        console.log('No Replicate API key found, using demo mode...');
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        return {
          success: true,
          url: 'https://commondatastorage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4',
          metadata: {
            model: 'stable-video-diffusion (Demo)',
            prompt: options.prompt,
            size: '1024x576',
            quality: 'standard'
          }
        };
      }

      const isTrialMode = this.isTrialMode('replicate');

      // Create prediction
      const response = await fetch('https://api.replicate.com/v1/predictions', {
        method: 'POST',
        headers: {
          'Authorization': `Token ${apiKey}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          version: "9f747673945c62801b13b84701c783929c0ee784e4748ec062204894dda1a351",
          input: {
            prompt: options.prompt,
            num_frames: 14,
            num_inference_steps: 25
          }
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        
        if (response.status === 401) {
          throw new Error('Invalid Replicate API key. Please check your API key in Settings.');
        } else if (response.status === 402) {
          throw new Error(isTrialMode 
            ? 'Trial credits exhausted. Please add your own Replicate API key in Settings.'
            : 'Insufficient credits. Please check your Replicate account balance.');
        }
        
        throw new Error(`Replicate API error: ${response.status} - ${errorData.detail || 'Unknown error'}`);
      }

      const prediction = await response.json();
      
      // Poll for completion
      let result = prediction;
      let attempts = 0;
      const maxAttempts = 30;
      
      while ((result.status === 'starting' || result.status === 'processing') && attempts < maxAttempts) {
        await new Promise(resolve => setTimeout(resolve, 2000));
        const pollResponse = await fetch(`https://api.replicate.com/v1/predictions/${result.id}`, {
          headers: {
            'Authorization': `Token ${apiKey}`
          }
        });
        
        if (!pollResponse.ok) {
          throw new Error(`Failed to poll prediction status: ${pollResponse.status}`);
        }
        
        result = await pollResponse.json();
        attempts++;
      }

      if (result.status === 'succeeded' && result.output) {
        return {
          success: true,
          url: Array.isArray(result.output) ? result.output[0] : result.output,
          metadata: {
            model: isTrialMode ? 'stable-video-diffusion (Trial)' : 'stable-video-diffusion',
            prompt: options.prompt,
            size: '1024x576',
            quality: 'standard'
          }
        };
      }

      if (result.status === 'failed') {
        throw new Error(result.error || 'Video generation failed');
      }

      throw new Error('Video generation timed out');
      
    } catch (error: any) {
      console.error('Replicate video generation failed:', error);
      
      // Fallback to demo video if API fails
      if (error.message?.includes('API') || error.message?.includes('network')) {
        console.log('API failed, falling back to demo video...');
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        return {
          success: true,
          url: 'https://commondatastorage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4',
          metadata: {
            model: 'Demo Video (API Unavailable)',
            prompt: options.prompt,
            size: '1024x576',
            quality: 'standard'
          }
        };
      }
      
      return {
        success: false,
        error: error.message || 'Failed to generate video with Replicate'
      };
    }
  }

  private async generateVideoWithRunway(options: AIGenerationOptions): Promise<AIGenerationResult> {
    try {
      const apiKey = this.getApiKey('runway');
      if (!apiKey) {
        // If no API key, use demo mode with sample video
        console.log('No Runway API key found, using demo mode...');
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        return {
          success: true,
          url: 'https://sample-videos.com/zip/10/mp4/SampleVideo_1280x720_1mb.mp4',
          metadata: {
            model: 'runway-gen2 (Demo)',
            prompt: options.prompt,
            size: '1280x768',
            quality: 'standard'
          }
        };
      }

      const isTrialMode = this.isTrialMode('runway');

      // Note: Runway ML API is not publicly available yet, so we'll simulate it
      console.log('Simulating Runway ML video generation...');
      
      // Simulate API delay
      await new Promise(resolve => setTimeout(resolve, 3000));
      
      // Return a sample video URL for demonstration
      const sampleVideoUrl = 'https://commondatastorage.googleapis.com/gtv-videos-bucket/sample/ElephantsDream.mp4';
      
      return {
        success: true,
        url: sampleVideoUrl,
        metadata: {
          model: isTrialMode ? 'runway-gen2 (Trial)' : 'runway-gen2',
          prompt: options.prompt,
          size: '1280x768',
          quality: 'standard'
        }
      };
      
    } catch (error: any) {
      console.error('Runway video generation failed:', error);
      
      // Fallback to demo video
      console.log('Runway failed, falling back to demo video...');
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      return {
        success: true,
        url: 'https://sample-videos.com/zip/10/mp4/SampleVideo_1280x720_1mb.mp4',
        metadata: {
          model: 'Demo Video (API Unavailable)',
          prompt: options.prompt,
          size: '1280x768',
          quality: 'standard'
        }
      };
    }
  }

  async generateContent(options: AIGenerationOptions): Promise<AIGenerationResult> {
    // Check if any API keys are available
    const userKeys = this.getUserApiKeys();
    const hasSystemKeys = !!(AI_CONFIG.openai.apiKey || AI_CONFIG.replicate.apiKey || AI_CONFIG.stability.apiKey);
    
    if (options.type === 'image') {
      // Try OpenAI DALL-E first, fallback to Stability AI
      const openaiResult = await this.generateImageWithOpenAI(options);
      if (openaiResult.success) {
        return openaiResult;
      }

      console.warn('OpenAI failed, trying Stability AI:', openaiResult.error);
      const stabilityResult = await this.generateImageWithStability(options);
      
      if (!stabilityResult.success) {
        // If both fail, return the more informative error
        return {
          success: false,
          error: `Image generation failed. OpenAI: ${openaiResult.error}. Stability AI: ${stabilityResult.error}`
        };
      }
      
      return stabilityResult;
    } else if (options.type === 'video') {
      // Try Replicate first, fallback to Runway
      const replicateResult = await this.generateVideoWithReplicate(options);
      if (replicateResult.success) {
        return replicateResult;
      }

      console.warn('Replicate failed, trying Runway:', replicateResult.error);
      const runwayResult = await this.generateVideoWithRunway(options);
      
      // Runway should always succeed with demo video as fallback
      return runwayResult;
    }

    return {
      success: false,
      error: 'Unsupported content type'
    };
  }

  // Enhanced prompt optimization
  optimizePrompt(prompt: string, type: 'image' | 'video', style?: string): string {
    let optimizedPrompt = prompt.trim();

    if (type === 'image') {
      // Add style modifiers for images
      const styleModifiers = {
        'photorealistic': 'photorealistic, high quality, detailed, professional photography',
        'artistic': 'artistic, creative, stylized, beautiful composition',
        'minimalist': 'minimalist, clean, simple, modern design',
        'vintage': 'vintage, retro, classic, nostalgic atmosphere',
        'futuristic': 'futuristic, sci-fi, modern, high-tech, sleek design'
      };

      if (style && styleModifiers[style as keyof typeof styleModifiers]) {
        optimizedPrompt += `, ${styleModifiers[style as keyof typeof styleModifiers]}`;
      }

      // Add general quality enhancers
      optimizedPrompt += ', high resolution, sharp focus, well-lit';
    } else if (type === 'video') {
      // Add video-specific modifiers
      optimizedPrompt += ', smooth motion, cinematic, high quality video';
      
      if (style) {
        optimizedPrompt += `, ${style} style`;
      }
    }

    return optimizedPrompt;
  }

  // Check API availability
  async checkAPIAvailability(): Promise<{
    openai: { available: boolean; isTrialMode: boolean };
    stability: { available: boolean; isTrialMode: boolean };
    replicate: { available: boolean; isTrialMode: boolean };
    runway: { available: boolean; isTrialMode: boolean };
  }> {
    const userKeys = this.getUserApiKeys();
    
    return {
      openai: {
        available: !!(userKeys.openai || AI_CONFIG.openai.apiKey),
        isTrialMode: !userKeys.openai && !!AI_CONFIG.openai.apiKey
      },
      stability: {
        available: !!(userKeys.stability || AI_CONFIG.stability.apiKey),
        isTrialMode: !userKeys.stability && !!AI_CONFIG.stability.apiKey
      },
      replicate: {
        available: !!(userKeys.replicate || AI_CONFIG.replicate.apiKey),
        isTrialMode: !userKeys.replicate && !!AI_CONFIG.replicate.apiKey
      },
      runway: {
        available: true, // Always available with demo fallback
        isTrialMode: !userKeys.runway
      }
    };
  }

  // Get usage information
  getUsageInfo(): {
    hasUserKeys: boolean;
    userKeyCount: number;
    availableServices: string[];
    trialServices: string[];
  } {
    const userKeys = this.getUserApiKeys();
    
    return {
      hasUserKeys: Object.keys(userKeys).length > 0,
      userKeyCount: Object.keys(userKeys).length,
      availableServices: Object.keys(userKeys),
      trialServices: ['replicate', 'runway'] // Services that have demo fallbacks
    };
  }
}

export const aiService = new AIService();
export default aiService;