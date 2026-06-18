import React from 'react';
import { Menu, Bell, User, Globe, LogOut, Search } from 'lucide-react';
import { useAuth } from '../contexts/AuthContext';
import { useLanguage } from '../contexts/LanguageContext';

interface HeaderProps {
  onMenuClick: () => void;
}

const Header: React.FC<HeaderProps> = ({ onMenuClick }) => {
  const { user, logout } = useAuth();
  const { language, setLanguage, t, isRTL } = useLanguage();

  const toggleLanguage = () => {
    setLanguage(language === 'en' ? 'ar' : 'en');
  };

  return (
    <header className="glass-card-dark border-b border-white/10 sticky top-0 z-30">
      <div className="flex items-center justify-between h-20 px-6">
        <div className="flex items-center space-x-4">
          <button
            onClick={onMenuClick}
            className="lg:hidden p-2 rounded-xl hover:bg-white/10 text-white/70 hover:text-white transition-colors"
          >
            <Menu className="h-5 w-5" />
          </button>
          
          {/* Search Bar */}
          <div className="hidden md:flex relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-white/40" />
            <input
              type="text"
              placeholder="Search..."
              className="pl-10 pr-4 py-2 w-80 bg-white/10 border border-white/20 rounded-xl text-white placeholder-white/40 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-transparent backdrop-blur-sm"
            />
          </div>
        </div>
        
        <div className="flex items-center space-x-4">
          <button
            onClick={toggleLanguage}
            className="p-2 rounded-xl hover:bg-white/10 transition-colors text-white/70 hover:text-white"
            title="Change Language"
          >
            <Globe className="h-5 w-5" />
            <span className="ml-1 text-sm">
              {language.toUpperCase()}
            </span>
          </button>
          
          <button className="p-2 rounded-xl hover:bg-white/10 transition-colors relative text-white/70 hover:text-white">
            <Bell className="h-5 w-5" />
            <span className="absolute -top-1 -right-1 h-3 w-3 bg-gradient-to-r from-pink-500 to-red-500 rounded-full animate-pulse"></span>
          </button>
          
          <div className="flex items-center space-x-3">
            <div className={`text-sm ${isRTL ? 'text-left' : 'text-right'} text-white`}>
              <div className="font-medium">{user?.name}</div>
              <div className="text-white/60 text-xs">{user?.company}</div>
            </div>
            <div className="h-10 w-10 bg-gradient-to-r from-blue-500 to-purple-600 rounded-xl flex items-center justify-center ring-2 ring-white/20">
              <User className="h-5 w-5 text-white" />
            </div>
          </div>
          
          <button
            onClick={logout}
            className="p-2 rounded-xl hover:bg-white/10 transition-colors text-white/70 hover:text-white"
            title={t('nav.logout')}
          >
            <LogOut className="h-5 w-5" />
          </button>
        </div>
      </div>
    </header>
  );
};

export default Header;