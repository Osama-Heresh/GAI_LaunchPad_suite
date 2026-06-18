import React, { useState } from 'react';
import { Type, Palette, Move } from 'lucide-react';

interface TextOverlayProps {
  text: string;
  onTextChange: (text: string) => void;
  position: { x: number; y: number };
  onPositionChange: (position: { x: number; y: number }) => void;
  fontSize: number;
  onFontSizeChange: (size: number) => void;
  color: string;
  onColorChange: (color: string) => void;
  fontFamily: string;
  onFontFamilyChange: (font: string) => void;
}

const TextOverlay: React.FC<TextOverlayProps> = ({
  text,
  onTextChange,
  position,
  onPositionChange,
  fontSize,
  onFontSizeChange,
  color,
  onColorChange,
  fontFamily,
  onFontFamilyChange
}) => {
  const [isEditing, setIsEditing] = useState(false);

  const fonts = [
    'Inter',
    'Arial',
    'Helvetica',
    'Georgia',
    'Times New Roman',
    'Courier New',
    'Verdana',
    'Impact'
  ];

  const presetColors = [
    '#FFFFFF', '#000000', '#FF0000', '#00FF00', '#0000FF',
    '#FFFF00', '#FF00FF', '#00FFFF', '#FFA500', '#800080'
  ];

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700">
      <div className="flex items-center space-x-2 mb-4">
        <Type className="h-5 w-5 text-blue-400" />
        <h3 className="text-lg font-semibold text-white">Text Overlay</h3>
      </div>

      <div className="space-y-4">
        {/* Text Input */}
        <div>
          <label className="block text-sm font-medium text-slate-300 mb-2">
            Overlay Text
          </label>
          <textarea
            value={text}
            onChange={(e) => onTextChange(e.target.value)}
            placeholder="Enter text to overlay on media..."
            rows={3}
            className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
          />
        </div>

        {/* Font Family */}
        <div>
          <label className="block text-sm font-medium text-slate-300 mb-2">
            Font Family
          </label>
          <select
            value={fontFamily}
            onChange={(e) => onFontFamilyChange(e.target.value)}
            className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            {fonts.map((font) => (
              <option key={font} value={font} style={{ fontFamily: font }}>
                {font}
              </option>
            ))}
          </select>
        </div>

        {/* Font Size */}
        <div>
          <label className="block text-sm font-medium text-slate-300 mb-2">
            Font Size: {fontSize}px
          </label>
          <input
            type="range"
            min="12"
            max="72"
            value={fontSize}
            onChange={(e) => onFontSizeChange(Number(e.target.value))}
            className="w-full h-2 bg-slate-600 rounded-lg appearance-none cursor-pointer slider"
          />
        </div>

        {/* Color Picker */}
        <div>
          <label className="block text-sm font-medium text-slate-300 mb-2">
            Text Color
          </label>
          <div className="flex items-center space-x-3">
            <input
              type="color"
              value={color}
              onChange={(e) => onColorChange(e.target.value)}
              className="w-12 h-10 rounded-lg border border-slate-600 bg-slate-700 cursor-pointer"
            />
            <div className="flex flex-wrap gap-2">
              {presetColors.map((presetColor) => (
                <button
                  key={presetColor}
                  onClick={() => onColorChange(presetColor)}
                  className={`w-6 h-6 rounded border-2 ${
                    color === presetColor ? 'border-blue-400' : 'border-slate-600'
                  }`}
                  style={{ backgroundColor: presetColor }}
                  aria-label={`Select color ${presetColor}`}
                />
              ))}
            </div>
          </div>
        </div>

        {/* Position Controls */}
        <div>
          <label className="block text-sm font-medium text-slate-300 mb-2">
            Position
          </label>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-xs text-slate-400 mb-1">X Position</label>
              <input
                type="range"
                min="0"
                max="100"
                value={position.x}
                onChange={(e) => onPositionChange({ ...position, x: Number(e.target.value) })}
                className="w-full h-2 bg-slate-600 rounded-lg appearance-none cursor-pointer slider"
              />
              <span className="text-xs text-slate-400">{position.x}%</span>
            </div>
            <div>
              <label className="block text-xs text-slate-400 mb-1">Y Position</label>
              <input
                type="range"
                min="0"
                max="100"
                value={position.y}
                onChange={(e) => onPositionChange({ ...position, y: Number(e.target.value) })}
                className="w-full h-2 bg-slate-600 rounded-lg appearance-none cursor-pointer slider"
              />
              <span className="text-xs text-slate-400">{position.y}%</span>
            </div>
          </div>
        </div>

        {/* Preview */}
        {text && (
          <div className="mt-4">
            <label className="block text-sm font-medium text-slate-300 mb-2">
              Preview
            </label>
            <div className="relative bg-slate-700 rounded-lg h-32 overflow-hidden">
              <div
                className="absolute pointer-events-none"
                style={{
                  left: `${position.x}%`,
                  top: `${position.y}%`,
                  transform: 'translate(-50%, -50%)',
                  fontSize: `${fontSize}px`,
                  color: color,
                  fontFamily: fontFamily,
                  textShadow: '2px 2px 4px rgba(0,0,0,0.5)',
                  whiteSpace: 'nowrap'
                }}
              >
                {text}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default TextOverlay;