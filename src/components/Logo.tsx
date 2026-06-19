import React from 'react';

interface LogoProps {
  className?: string;
  size?: 'sm' | 'md' | 'lg';
}

export const Logo: React.FC<LogoProps> = ({ className = '', size = 'sm' }) => {
  const dimensionClass = {
    sm: 'h-8 w-8',
    md: 'h-12 w-12',
    lg: 'h-16 w-16',
  }[size];

  return (
    <svg
      className={`${dimensionClass} ${className} transition-transform duration-300 hover:scale-105`}
      viewBox="0 0 100 100"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <defs>
        {/* Core background gradient */}
        <linearGradient id="logoGrad" x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor="#3B82F6" /> {/* Indigo blue */}
          <stop offset="50%" stopColor="#8B5CF6" /> {/* Deep Violet */}
          <stop offset="100%" stopColor="#EC4899" /> {/* vibrant Pink */}
        </linearGradient>
        
        {/* Shadow effect */}
        <filter id="logoGlow" x="-10%" y="-10%" width="120%" height="120%">
          <feDropShadow dx="0" dy="4" stdDeviation="6" floodColor="#8B5CF6" floodOpacity="0.4" />
        </filter>
      </defs>

      {/* Futuristic Geometric Backdrop */}
      <rect
        x="10"
        y="10"
        width="80"
        height="80"
        rx="22"
        fill="url(#logoGrad)"
        filter="url(#logoGlow)"
        className="opacity-95"
      />

      {/* Inner Tech-Grid Lines (adds a sophisticated blueprint detail) */}
      <circle cx="50" cy="50" r="30" stroke="white" strokeWidth="1" strokeDasharray="4 4" className="opacity-20" />
      <line x1="20" y1="50" x2="80" y2="50" stroke="white" strokeWidth="1" strokeDasharray="3 3" className="opacity-15" />
      <line x1="50" y1="20" x2="50" y2="80" stroke="white" strokeWidth="1" strokeDasharray="3 3" className="opacity-15" />

      {/* Styled LaunchPad Rocket / Launch Triangle */}
      {/* Dynamic upward chevron representing acceleration/launch, coupled with abstract engine flame */}
      <path
        d="M50 24L74 65H60L50 48L40 65H26L50 24Z"
        fill="white"
        className="drop-shadow-sm"
      />
      
      {/* Inner glowing core core */}
      <path
        d="M50 35L62 56H54L50 48L46 56H38L50 35Z"
        fill="#FFFFFF"
        className="opacity-90"
      />

      {/* Active Propulsion Flare Point */}
      <circle cx="50" cy="62" r="4" fill="#FBBF24" className="animate-pulse" />
    </svg>
  );
};

export default Logo;
