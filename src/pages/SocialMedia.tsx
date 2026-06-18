import React, { useState, useCallback, useMemo } from 'react';
import MediaUpload from '../components/MediaUpload';
import TextOverlay from '../components/TextOverlay';
import PostEditor from '../components/PostEditor';
import PlatformSelector from '../components/PlatformSelector';
import SchedulingPanel from '../components/SchedulingPanel';
import ToastContainer from '../components/ToastContainer';
import { ToastProps } from '../components/Toast';
import { useLanguage } from '../contexts/LanguageContext';

const SocialMedia: React.FC = () => {
  const { t } = useLanguage();
  
  // Media state
  const [activeTab, setActiveTab] = useState<'image' | 'video' | 'text'>('image');
  const [uploadedMedia, setUploadedMedia] = useState<string | null>(null);
  
  // Text overlay state
  const [overlayText, setOverlayText] = useState('');
  const [overlayPosition, setOverlayPosition] = useState({ x: 50, y: 50 });
  const [overlayFontSize, setOverlayFontSize] = useState(24);
  const [overlayColor, setOverlayColor] = useState('#FFFFFF');
  const [overlayFontFamily, setOverlayFontFamily] = useState('Inter');
  
  // Post content state
  const [postContent, setPostContent] = useState('');
  
  // Platform selection state
  const [selectedPlatforms, setSelectedPlatforms] = useState<string[]>([]);
  
  // Scheduling state
  const [postNow, setPostNow] = useState(true);
  const [scheduleLater, setScheduleLater] = useState(false);
  const [scheduledDate, setScheduledDate] = useState('');
  const [scheduledTime, setScheduledTime] = useState('');
  
  // UI state
  const [isPublishing, setIsPublishing] = useState(false);
  const [toasts, setToasts] = useState<ToastProps[]>([]);

  // Character limits for different platforms - static object
  const characterLimits = useMemo(() => ({
    twitter: 280,
    facebook: 63206,
    instagram: 2200,
    linkedin: 3000,
    tiktok: 150,
    slack: 4000,
    discord: 2000
  }), []);

  // Toast management
  const addToast = useCallback((toast: Omit<ToastProps, 'id' | 'onClose'>) => {
    const id = Date.now().toString();
    setToasts(prev => [...prev, { ...toast, id, onClose: removeToast }]);
  }, []);

  const removeToast = useCallback((id: string) => {
    setToasts(prev => prev.filter(toast => toast.id !== id));
  }, []);

  // Media handling
  const handleMediaUpload = useCallback((file: File) => {
    const url = URL.createObjectURL(file);
    setUploadedMedia(url);
    addToast({
      type: 'success',
      message: `${activeTab} uploaded successfully!`
    });
  }, [activeTab, addToast]);

  const handleAIGenerate = useCallback(async (prompt: string, type: 'image' | 'video') => {
    console.log(`AI Generate called with prompt: "${prompt}" and type: "${type}"`);
    
    // Show loading toast
    const loadingToastId = Date.now().toString();
    setToasts(prev => [...prev, {
      id: loadingToastId,
      type: 'loading',
      message: `Generating ${type} with AI...`,
      onClose: removeToast
    }]);

    try {
      // Simulate AI generation delay
      await new Promise(resolve => setTimeout(resolve, 3000));
      
      // Remove loading toast
      setToasts(prev => prev.filter(t => t.id !== loadingToastId));
      
      // Create mock media URL based on type
      let mockUrl: string;
      if (type === 'image') {
        mockUrl = 'https://images.pexels.com/photos/3184418/pexels-photo-3184418.jpeg?auto=compress&cs=tinysrgb&w=800';
      } else {
        // Use a reliable video URL for demo
        mockUrl = 'https://commondatastorage.googleapis.com/gtv-videos-bucket/sample/BigBuckBunny.mp4';
      }
      
      console.log(`Setting media URL: ${mockUrl}`);
      setUploadedMedia(mockUrl);
      
      addToast({
        type: 'success',
        message: `AI ${type} generated successfully! ${type === 'video' ? '(Demo video)' : ''}`
      });
      
    } catch (error: any) {
      console.error('AI generation simulation failed:', error);
      
      // Remove loading toast
      setToasts(prev => prev.filter(t => t.id !== loadingToastId));
      
      addToast({
        type: 'error',
        message: `Failed to generate ${type}. Please try again.`
      });
    }
  }, [addToast, removeToast]);

  const handleRemoveMedia = useCallback(() => {
    if (uploadedMedia) {
      URL.revokeObjectURL(uploadedMedia);
    }
    setUploadedMedia(null);
    setOverlayText('');
  }, [uploadedMedia]);

  // Platform handling
  const handlePlatformToggle = useCallback((platformId: string) => {
    setSelectedPlatforms(prev => 
      prev.includes(platformId)
        ? prev.filter(id => id !== platformId)
        : [...prev, platformId]
    );
  }, []);

  // Publishing
  const canPublish = useMemo(() => 
    (postContent.trim() || uploadedMedia) && selectedPlatforms.length > 0,
    [postContent, uploadedMedia, selectedPlatforms.length]
  );

  const handlePublish = useCallback(async () => {
    if (!canPublish) return;

    setIsPublishing(true);
    addToast({
      type: 'loading',
      message: postNow ? 'Publishing post...' : 'Scheduling post...'
    });

    // Simulate API call
    setTimeout(() => {
      setIsPublishing(false);
      setToasts(prev => prev.filter(t => t.type !== 'loading'));
      
      addToast({
        type: 'success',
        message: postNow 
          ? `Post published to ${selectedPlatforms.length} platform${selectedPlatforms.length > 1 ? 's' : ''}!`
          : `Post scheduled for ${new Date(`${scheduledDate}T${scheduledTime}`).toLocaleString()}!`
      });

      // Reset form
      setPostContent('');
      setUploadedMedia(null);
      setOverlayText('');
      setSelectedPlatforms([]);
      setPostNow(true);
      setScheduleLater(false);
      setScheduledDate('');
      setScheduledTime('');
    }, 2000);
  }, [canPublish, postNow, selectedPlatforms.length, scheduledDate, scheduledTime, addToast]);

  return (
    <div className="space-y-8">
      {/* Page Header */}
      <div className="flex items-center justify-between">
        <h1 className="text-3xl font-bold gradient-text">{t('social.title')}</h1>
      </div>

      {/* Main Content */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Left Column */}
        <div className="space-y-8">
          {/* Media Upload */}
          <MediaUpload
            activeTab={activeTab}
            onTabChange={setActiveTab}
            onMediaUpload={handleMediaUpload}
            onAIGenerate={handleAIGenerate}
            uploadedMedia={uploadedMedia}
            onRemoveMedia={handleRemoveMedia}
          />

          {/* Text Overlay (only show if media is uploaded) */}
          {uploadedMedia && activeTab !== 'text' && (
            <TextOverlay
              text={overlayText}
              onTextChange={setOverlayText}
              position={overlayPosition}
              onPositionChange={setOverlayPosition}
              fontSize={overlayFontSize}
              onFontSizeChange={setOverlayFontSize}
              color={overlayColor}
              onColorChange={setOverlayColor}
              fontFamily={overlayFontFamily}
              onFontFamilyChange={setOverlayFontFamily}
            />
          )}

          {/* Post Editor */}
          <PostEditor
            content={postContent}
            onContentChange={setPostContent}
            selectedPlatforms={selectedPlatforms}
            characterLimits={characterLimits}
          />
        </div>

        {/* Right Column */}
        <div className="space-y-8">
          {/* Platform Selector */}
          <PlatformSelector
            selectedPlatforms={selectedPlatforms}
            onPlatformToggle={handlePlatformToggle}
          />

          {/* Scheduling Panel */}
          <SchedulingPanel
            postNow={postNow}
            onPostNowChange={setPostNow}
            scheduleLater={scheduleLater}
            onScheduleLaterChange={setScheduleLater}
            scheduledDate={scheduledDate}
            onScheduledDateChange={setScheduledDate}
            scheduledTime={scheduledTime}
            onScheduledTimeChange={setScheduledTime}
            onPublish={handlePublish}
            canPublish={canPublish}
            isPublishing={isPublishing}
          />
        </div>
      </div>

      {/* Toast Notifications */}
      <ToastContainer toasts={toasts} onRemoveToast={removeToast} />
    </div>
  );
};

export default SocialMedia;