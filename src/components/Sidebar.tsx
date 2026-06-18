import React from 'react';
import { NavLink } from 'react-router-dom';
import {
  LayoutDashboard,
  Share2,
  CheckSquare,
  FileText,
  Users,
  Settings,
  PlayCircle,
  X,
  UsersRound
} from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';
import { useAuth } from '../contexts/AuthContext';

interface SidebarProps {
  isOpen: boolean;
  onClose: () => void;
}

const Sidebar: React.FC<SidebarProps> = ({ isOpen, onClose }) => {
  const { t, isRTL } = useLanguage();
  const { user } = useAuth();

  const navigation = [
    { name: t('nav.dashboard'), href: '/dashboard', icon: LayoutDashboard },
    { name: t('nav.socialMedia'), href: '/social-media', icon: Share2 },
    { name: t('nav.tasks'), href: '/tasks', icon: CheckSquare },
    { name: 'Team', href: '/team', icon: UsersRound },
    { name: t('nav.templates'), href: '/templates', icon: FileText },
    ...(user?.role === 'admin' ? [{ name: t('nav.users'), href: '/users', icon: Users }] : []),
    { name: t('nav.training'), href: '/training', icon: PlayCircle },
    { name: t('nav.settings'), href: '/settings', icon: Settings },
  ];

  return (
    <>
      {/* Mobile overlay */}
      {isOpen && (
        <div 
          className="fixed inset-0 bg-black/50 backdrop-blur-sm z-40 lg:hidden"
          onClick={onClose}
        />
      )}
      
      {/* Sidebar */}
      <div className={`fixed top-0 ${isRTL ? 'right-0' : 'left-0'} z-50 h-full w-64 glass-card-dark border-r border-white/10 transform transition-transform duration-300 ease-in-out ${
        isOpen ? 'translate-x-0' : isRTL ? 'translate-x-full' : '-translate-x-full'
      } lg:translate-x-0`}>
        <div className="flex items-center justify-between h-16 px-4 border-b border-white/10">
          <div className="flex items-center space-x-3">
            <img 
              src="/Screenshot_29-6-2025_15959_www.design.com.jpeg" 
              alt="LaunchPad Suite" 
              className="h-8 w-8 object-contain"
            />
            <span className="text-sm font-bold gradient-text">LaunchPad Suite</span>
          </div>
          <button
            onClick={onClose}
            className="lg:hidden p-2 rounded-lg hover:bg-white/10 text-white/70 hover:text-white transition-colors"
          >
            <X className="h-4 w-4" />
          </button>
        </div>
        
        <nav className="mt-6 px-3">
          <div className="space-y-2">
            {navigation.map((item) => (
              <NavLink
                key={item.name}
                to={item.href}
                className={({ isActive }) =>
                  `group flex items-center px-3 py-3 text-base font-medium rounded-xl transition-all duration-200 ${
                    isActive
                      ? 'bg-gradient-to-r from-blue-500/20 to-purple-500/20 text-white border border-blue-500/30 neon-glow'
                      : 'text-white/70 hover:text-white hover:bg-white/10'
                  }`
                }
                onClick={() => onClose()}
              >
                <item.icon className={`${isRTL ? 'ml-3' : 'mr-3'} h-5 w-5 flex-shrink-0`} />
                <span className="truncate text-base">{item.name}</span>
              </NavLink>
            ))}
          </div>
        </nav>
        
        <div className="absolute bottom-0 w-full p-4 border-t border-white/10">
          <div className="glass-card p-3 text-center">
            <div className="text-sm text-white/60">
              LaunchPad Suite v2.0
            </div>
            <div className="text-sm text-white/40 mt-1">
              Your Business Hub
            </div>
          </div>
        </div>
      </div>
    </>
  );
};

export default Sidebar;