import React, { useState } from 'react';
import { Calendar, Clock, Send, CheckCircle } from 'lucide-react';

interface SchedulingPanelProps {
  postNow: boolean;
  onPostNowChange: (value: boolean) => void;
  scheduleLater: boolean;
  onScheduleLaterChange: (value: boolean) => void;
  scheduledDate: string;
  onScheduledDateChange: (date: string) => void;
  scheduledTime: string;
  onScheduledTimeChange: (time: string) => void;
  onPublish: () => void;
  canPublish: boolean;
  isPublishing: boolean;
}

const SchedulingPanel: React.FC<SchedulingPanelProps> = ({
  postNow,
  onPostNowChange,
  scheduleLater,
  onScheduleLaterChange,
  scheduledDate,
  onScheduledDateChange,
  scheduledTime,
  onScheduledTimeChange,
  onPublish,
  canPublish,
  isPublishing
}) => {
  const handlePostNowChange = (checked: boolean) => {
    onPostNowChange(checked);
    if (checked) {
      onScheduleLaterChange(false);
    }
  };

  const handleScheduleLaterChange = (checked: boolean) => {
    onScheduleLaterChange(checked);
    if (checked) {
      onPostNowChange(false);
      // Set default date/time if not already set
      if (!scheduledDate) {
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        onScheduledDateChange(tomorrow.toISOString().split('T')[0]);
      }
      if (!scheduledTime) {
        onScheduledTimeChange('09:00');
      }
    }
  };

  const getMinDate = () => {
    return new Date().toISOString().split('T')[0];
  };

  const getMinTime = () => {
    const now = new Date();
    const selectedDate = new Date(scheduledDate);
    const today = new Date().toISOString().split('T')[0];
    
    if (scheduledDate === today) {
      // If selected date is today, minimum time is current time + 5 minutes
      now.setMinutes(now.getMinutes() + 5);
      return now.toTimeString().slice(0, 5);
    }
    return '00:00';
  };

  return (
    <div className="bg-slate-800 rounded-xl p-6 border border-slate-700">
      <h2 className="text-xl font-semibold text-white mb-4">Scheduling & Publishing</h2>
      
      <div className="space-y-4">
        {/* Post Now Option */}
        <div className="flex items-center space-x-3">
          <input
            type="checkbox"
            id="postNow"
            checked={postNow}
            onChange={(e) => handlePostNowChange(e.target.checked)}
            className="w-4 h-4 text-blue-600 bg-slate-700 border-slate-600 rounded focus:ring-blue-500 focus:ring-2"
          />
          <label htmlFor="postNow" className="flex items-center space-x-2 text-white cursor-pointer">
            <Send className="h-4 w-4 text-green-400" />
            <span>Post now</span>
          </label>
        </div>

        {/* Schedule Later Option */}
        <div className="flex items-center space-x-3">
          <input
            type="checkbox"
            id="scheduleLater"
            checked={scheduleLater}
            onChange={(e) => handleScheduleLaterChange(e.target.checked)}
            className="w-4 h-4 text-blue-600 bg-slate-700 border-slate-600 rounded focus:ring-blue-500 focus:ring-2"
          />
          <label htmlFor="scheduleLater" className="flex items-center space-x-2 text-white cursor-pointer">
            <Calendar className="h-4 w-4 text-blue-400" />
            <span>Schedule for later</span>
          </label>
        </div>

        {/* Date/Time Picker */}
        {scheduleLater && (
          <div className="ml-7 space-y-3 p-4 bg-slate-700/50 rounded-lg border border-slate-600">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-300 mb-2">
                  Date
                </label>
                <input
                  type="date"
                  value={scheduledDate}
                  onChange={(e) => onScheduledDateChange(e.target.value)}
                  min={getMinDate()}
                  className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-300 mb-2">
                  Time
                </label>
                <input
                  type="time"
                  value={scheduledTime}
                  onChange={(e) => onScheduledTimeChange(e.target.value)}
                  min={getMinTime()}
                  className="w-full px-3 py-2 bg-slate-700 border border-slate-600 rounded-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>
            </div>
            
            {scheduledDate && scheduledTime && (
              <div className="flex items-center space-x-2 text-sm text-slate-300">
                <Clock className="h-4 w-4" />
                <span>
                  Scheduled for: {new Date(`${scheduledDate}T${scheduledTime}`).toLocaleString()}
                </span>
              </div>
            )}
          </div>
        )}

        {/* Publish Button */}
        <div className="pt-4 border-t border-slate-700">
          <button
            onClick={onPublish}
            disabled={!canPublish || isPublishing}
            className={`w-full flex items-center justify-center space-x-2 px-6 py-4 rounded-xl font-semibold text-lg transition-all ${
              canPublish && !isPublishing
                ? 'bg-gradient-to-r from-blue-600 to-purple-600 hover:from-blue-500 hover:to-purple-500 text-white shadow-lg hover:shadow-xl'
                : 'bg-slate-700 text-slate-400 cursor-not-allowed'
            }`}
          >
            {isPublishing ? (
              <>
                <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                <span>Publishing...</span>
              </>
            ) : postNow ? (
              <>
                <Send className="h-5 w-5" />
                <span>Publish Now</span>
              </>
            ) : scheduleLater ? (
              <>
                <Calendar className="h-5 w-5" />
                <span>Schedule Post</span>
              </>
            ) : (
              <>
                <CheckCircle className="h-5 w-5" />
                <span>Publish</span>
              </>
            )}
          </button>
          
          {!canPublish && (
            <p className="text-sm text-slate-400 text-center mt-2">
              Select at least one platform and add content to publish
            </p>
          )}
        </div>
      </div>
    </div>
  );
};

export default SchedulingPanel;