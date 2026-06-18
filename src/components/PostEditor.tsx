import React, { useState, useRef, useEffect } from 'react';
import { Bold, Italic, Link, Hash, AtSign } from 'lucide-react';

interface PostEditorProps {
  content: string;
  onContentChange: (content: string) => void;
  selectedPlatforms: string[];
  characterLimits: Record<string, number>;
}

const PostEditor: React.FC<PostEditorProps> = ({
  content,
  onContentChange,
  selectedPlatforms,
  characterLimits
}) => {
  const [isFocused, setIsFocused] = useState(false);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const insertText = (before: string, after: string = '') => {
    const textarea = textareaRef.current;
    if (!textarea) return;

    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const selectedText = content.substring(start, end);
    const newText = content.substring(0, start) + before + selectedText + after + content.substring(end);
    
    onContentChange(newText);
    
    // Restore cursor position
    setTimeout(() => {
      textarea.focus();
      textarea.setSelectionRange(start + before.length, end + before.length);
    }, 0);
  };

  const formatButtons = [
    { icon: Bold, action: () => insertText('**', '**'), label: 'Bold' },
    { icon: Italic, action: () => insertText('_', '_'), label: 'Italic' },
    { icon: Hash, action: () => insertText('#'), label: 'Hashtag' },
    { icon: AtSign, action: () => insertText('@'), label: 'Mention' },
    { icon: Link, action: () => insertText('[', '](url)'), label: 'Link' }
  ];

  const getCharacterCount = (platform: string) => {
    const limit = characterLimits[platform] || 280;
    const remaining = limit - content.length;
    return { used: content.length, limit, remaining };
  };

  const getWorstCase = () => {
    if (selectedPlatforms.length === 0) return { used: content.length, limit: 280, remaining: 280 - content.length };
    
    const limits = selectedPlatforms.map(platform => characterLimits[platform] || 280);
    const minLimit = Math.min(...limits);
    return { used: content.length, limit: minLimit, remaining: minLimit - content.length };
  };

  const worstCase = getWorstCase();

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700">
      <h2 className="text-xl font-semibold text-white mb-4">Post Content</h2>
      
      {/* Formatting Toolbar */}
      <div className="flex items-center space-x-2 mb-4 p-2 bg-slate-700/50 rounded-lg">
        {formatButtons.map((button, index) => (
          <button
            key={index}
            onClick={button.action}
            className="p-2 text-slate-300 hover:text-white hover:bg-slate-600 rounded-md transition-colors"
            title={button.label}
            aria-label={button.label}
          >
            <button.icon className="h-4 w-4" />
          </button>
        ))}
      </div>

      {/* Text Editor */}
      <div className="relative">
        <textarea
          ref={textareaRef}
          value={content}
          onChange={(e) => onContentChange(e.target.value)}
          onFocus={() => setIsFocused(true)}
          onBlur={() => setIsFocused(false)}
          placeholder="What's happening? Share your thoughts..."
          rows={6}
          className={`w-full px-4 py-3 bg-slate-700 border rounded-lg text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:border-transparent resize-none transition-all ${
            isFocused ? 'border-blue-500 ring-blue-500/50' : 'border-slate-600'
          } ${worstCase.remaining < 0 ? 'border-red-500 ring-red-500/50' : ''}`}
        />
        
        {/* Character Counter */}
        <div className="absolute bottom-3 right-3 flex items-center space-x-2">
          <div className={`text-sm font-medium ${
            worstCase.remaining < 0 ? 'text-red-400' : 
            worstCase.remaining < 20 ? 'text-yellow-400' : 'text-slate-400'
          }`}>
            {worstCase.used}/{worstCase.limit}
          </div>
          <div className={`w-8 h-8 rounded-full border-2 flex items-center justify-center ${
            worstCase.remaining < 0 ? 'border-red-400' : 
            worstCase.remaining < 20 ? 'border-yellow-400' : 'border-slate-500'
          }`}>
            <div
              className={`w-6 h-6 rounded-full ${
                worstCase.remaining < 0 ? 'bg-red-400' : 
                worstCase.remaining < 20 ? 'bg-yellow-400' : 'bg-slate-500'
              }`}
              style={{
                transform: `scale(${Math.min(worstCase.used / worstCase.limit, 1)})`
              }}
            />
          </div>
        </div>
      </div>

      {/* Platform-specific character counts */}
      {selectedPlatforms.length > 1 && (
        <div className="mt-4 grid grid-cols-2 md:grid-cols-3 gap-3">
          {selectedPlatforms.map((platform) => {
            const count = getCharacterCount(platform);
            return (
              <div
                key={platform}
                className={`p-3 rounded-lg border ${
                  count.remaining < 0 ? 'border-red-500/50 bg-red-500/10' :
                  count.remaining < 20 ? 'border-yellow-500/50 bg-yellow-500/10' :
                  'border-slate-600 bg-slate-700/50'
                }`}
              >
                <div className="text-sm font-medium text-white capitalize">{platform}</div>
                <div className={`text-xs ${
                  count.remaining < 0 ? 'text-red-400' :
                  count.remaining < 20 ? 'text-yellow-400' : 'text-slate-400'
                }`}>
                  {count.used}/{count.limit} ({count.remaining > 0 ? '+' : ''}{count.remaining})
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Content Preview */}
      {content && (
        <div className="mt-4 p-4 bg-slate-700/50 rounded-lg">
          <h4 className="text-sm font-medium text-slate-300 mb-2">Preview</h4>
          <div className="text-white text-sm leading-relaxed whitespace-pre-wrap">
            {content}
          </div>
        </div>
      )}
    </div>
  );
};

export default PostEditor;