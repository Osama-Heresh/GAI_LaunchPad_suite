import React, { useEffect } from 'react';
import { CheckCircle, XCircle, AlertCircle, X } from 'lucide-react';

export interface ToastProps {
  id: string;
  type: 'success' | 'error' | 'loading';
  message: string;
  onClose: (id: string) => void;
  duration?: number;
}

const Toast: React.FC<ToastProps> = ({ id, type, message, onClose, duration = 5000 }) => {
  useEffect(() => {
    if (type !== 'loading' && duration > 0) {
      const timer = setTimeout(() => {
        onClose(id);
      }, duration);
      
      return () => clearTimeout(timer);
    }
  }, [id, type, duration, onClose]);

  const getIcon = () => {
    switch (type) {
      case 'success':
        return <CheckCircle className="h-5 w-5 text-green-400" />;
      case 'error':
        return <XCircle className="h-5 w-5 text-red-400" />;
      case 'loading':
        return <div className="w-5 h-5 border-2 border-blue-400/30 border-t-blue-400 rounded-full animate-spin" />;
      default:
        return <AlertCircle className="h-5 w-5 text-yellow-400" />;
    }
  };

  const getStyles = () => {
    switch (type) {
      case 'success':
        return 'bg-green-900/50 border-green-500/50 text-green-100';
      case 'error':
        return 'bg-red-900/50 border-red-500/50 text-red-100';
      case 'loading':
        return 'bg-blue-900/50 border-blue-500/50 text-blue-100';
      default:
        return 'bg-yellow-900/50 border-yellow-500/50 text-yellow-100';
    }
  };

  return (
    <div
      className={`flex items-center space-x-3 p-4 rounded-lg border backdrop-blur-sm ${getStyles()} animate-in slide-in-from-right-full duration-300`}
      role="alert"
      aria-live="polite"
    >
      {getIcon()}
      <span className="flex-1 text-sm font-medium">{message}</span>
      {type !== 'loading' && (
        <button
          onClick={() => onClose(id)}
          className="text-current hover:opacity-70 transition-opacity"
          aria-label="Close notification"
        >
          <X className="h-4 w-4" />
        </button>
      )}
    </div>
  );
};

export default Toast;