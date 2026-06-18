import React from 'react';
import { useNavigate } from 'react-router-dom';
import { Users, CheckSquare, Calendar, FileText, TrendingUp, Activity, ArrowUp, ArrowDown } from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';
import { useAuth } from '../contexts/AuthContext';

const Dashboard: React.FC = () => {
  const { t } = useLanguage();
  const { user } = useAuth();
  const navigate = useNavigate();

  const stats = [
    { 
      name: t('dashboard.activeUsers'), 
      value: '24', 
      change: '+12%',
      trend: 'up',
      icon: Users, 
      gradient: 'from-blue-500 to-cyan-500' 
    },
    { 
      name: t('dashboard.pendingTasks'), 
      value: '12', 
      change: '-8%',
      trend: 'down',
      icon: CheckSquare, 
      gradient: 'from-green-500 to-emerald-500' 
    },
    { 
      name: t('dashboard.scheduledPosts'), 
      value: '8', 
      change: '+24%',
      trend: 'up',
      icon: Calendar, 
      gradient: 'from-purple-500 to-pink-500' 
    },
    { 
      name: t('dashboard.templates'), 
      value: '156', 
      change: '+5%',
      trend: 'up',
      icon: FileText, 
      gradient: 'from-orange-500 to-red-500' 
    },
  ];

  const chartData = [
    { day: 'M', value: 65, label: 'Monday' },
    { day: 'T', value: 45, label: 'Tuesday' },
    { day: 'W', value: 80, label: 'Wednesday' },
    { day: 'T', value: 70, label: 'Thursday' },
    { day: 'F', value: 90, label: 'Friday' },
    { day: 'S', value: 55, label: 'Saturday' },
    { day: 'S', value: 40, label: 'Sunday' },
  ];

  const socialStats = [
    { platform: 'Facebook', followers: '263,650', change: '+90.2k Weekly', color: 'from-blue-600 to-blue-500' },
    { platform: 'Instagram', followers: '334,854', change: '+12% Monthly', color: 'from-pink-600 to-purple-500' },
    { platform: 'Twitter', followers: '334,854', change: '+12% Monthly', color: 'from-sky-500 to-blue-500' },
  ];

  const recentActivity = [
    { user: 'John Doe', action: 'Created new social media post', time: '2 minutes ago', type: 'post' },
    { user: 'Sarah Smith', action: 'Completed task: Update website content', time: '15 minutes ago', type: 'task' },
    { user: 'Mike Johnson', action: 'Used HR template: Employee Welcome Letter', time: '1 hour ago', type: 'template' },
    { user: 'Emily Brown', action: 'Scheduled post for LinkedIn', time: '2 hours ago', type: 'schedule' },
  ];

  const handleCreatePost = () => {
    // Scroll to top before navigation
    window.scrollTo(0, 0);
    navigate('/social-media');
  };

  const handleAddTask = () => {
    // Scroll to top before navigation
    window.scrollTo(0, 0);
    navigate('/tasks');
  };

  const handleManageUsers = () => {
    // Scroll to top before navigation
    window.scrollTo(0, 0);
    navigate('/users');
  };

  return (
    <div className="space-y-8">
      {/* Welcome Section */}
      <div className="glass-card p-8 relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-r from-blue-600/20 via-purple-600/20 to-pink-600/20"></div>
        <div className="relative z-10 flex items-center justify-between">
          <div>
            <h1 className="text-4xl font-bold text-white mb-2">
              Welcome to <span className="gradient-text">LaunchPad Suite</span>, {user?.name}!
            </h1>
            <p className="text-white/70 text-lg">Your comprehensive business management platform</p>
          </div>
          <div className="hidden lg:block">
            <TrendingUp className="h-20 w-20 text-white/20 animate-float" />
          </div>
        </div>
      </div>

      {/* Stats Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        {stats.map((stat) => (
          <div key={stat.name} className="glass-card p-6 hover:scale-105 transition-transform duration-200">
            <div className="flex items-center justify-between mb-4">
              <div className={`p-3 rounded-xl bg-gradient-to-r ${stat.gradient}`}>
                <stat.icon className="h-6 w-6 text-white" />
              </div>
              <div className={`flex items-center text-sm ${stat.trend === 'up' ? 'text-green-400' : 'text-red-400'}`}>
                {stat.trend === 'up' ? <ArrowUp className="h-4 w-4 mr-1" /> : <ArrowDown className="h-4 w-4 mr-1" />}
                {stat.change}
              </div>
            </div>
            <div>
              <p className="text-white/60 text-sm mb-1">{stat.name}</p>
              <p className="text-3xl font-bold text-white">{stat.value}</p>
            </div>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Social Media Stats */}
        <div className="glass-card p-6">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-xl font-semibold text-white">Social Media Overview</h2>
            <div className="flex items-center space-x-2 text-white/60 text-sm">
              <Calendar className="h-4 w-4" />
              <span>11-20 Jan</span>
            </div>
          </div>
          
          <div className="grid grid-cols-1 gap-4 mb-6">
            {socialStats.map((social) => (
              <div key={social.platform} className="glass-card p-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center space-x-3">
                    <div className={`w-3 h-3 rounded-full bg-gradient-to-r ${social.color}`}></div>
                    <span className="text-white font-medium">{social.platform}</span>
                  </div>
                  <TrendingUp className="h-4 w-4 text-green-400" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-white">{social.followers}</div>
                  <div className="text-sm text-white/60">{social.change}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Daily Visitors Chart */}
        <div className="glass-card p-6">
          <div className="flex items-center justify-between mb-6">
            <h2 className="text-xl font-semibold text-white">Daily Visitors</h2>
            <div className="text-white/60 text-sm">100k</div>
          </div>
          
          <div className="flex items-end justify-between h-48 mb-4">
            {chartData.map((item, index) => (
              <div key={index} className="flex flex-col items-center space-y-2">
                <div className="relative">
                  <div 
                    className="w-8 bg-gradient-to-t from-blue-600 to-blue-400 rounded-t-lg transition-all duration-500 hover:from-blue-500 hover:to-blue-300"
                    style={{ height: `${item.value * 2}px` }}
                  ></div>
                </div>
                <span className="text-white/60 text-xs">{item.day}</span>
              </div>
            ))}
          </div>
          
          <div className="grid grid-cols-3 gap-4 text-center">
            <div>
              <div className="text-xl font-bold text-white">25.20K</div>
              <div className="text-xs text-white/60">Facebook</div>
            </div>
            <div>
              <div className="text-xl font-bold text-white">38.94K</div>
              <div className="text-xs text-white/60">Instagram</div>
            </div>
            <div>
              <div className="text-xl font-bold text-white">75.05K</div>
              <div className="text-xs text-white/60">Twitter</div>
            </div>
          </div>
        </div>
      </div>

      {/* Quick Actions */}
      <div className="glass-card p-6">
        <h2 className="text-xl font-semibold text-white mb-6">{t('dashboard.quickActions')}</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <button 
            onClick={handleCreatePost}
            className="glass-card p-6 hover:bg-white/20 transition-all duration-200 group"
          >
            <div className="text-center">
              <FileText className="h-8 w-8 text-blue-400 mx-auto mb-3 group-hover:scale-110 transition-transform" />
              <p className="text-white font-medium">Create New Post</p>
              <p className="text-white/60 text-sm mt-1">Share your content</p>
            </div>
          </button>
          <button 
            onClick={handleAddTask}
            className="glass-card p-6 hover:bg-white/20 transition-all duration-200 group"
          >
            <div className="text-center">
              <CheckSquare className="h-8 w-8 text-green-400 mx-auto mb-3 group-hover:scale-110 transition-transform" />
              <p className="text-white font-medium">Add New Task</p>
              <p className="text-white/60 text-sm mt-1">Organize your work</p>
            </div>
          </button>
          <button 
            onClick={handleManageUsers}
            className="glass-card p-6 hover:bg-white/20 transition-all duration-200 group"
          >
            <div className="text-center">
              <Users className="h-8 w-8 text-purple-400 mx-auto mb-3 group-hover:scale-110 transition-transform" />
              <p className="text-white font-medium">Manage Users</p>
              <p className="text-white/60 text-sm mt-1">Control access</p>
            </div>
          </button>
        </div>
      </div>

      {/* Recent Activity */}
      <div className="glass-card p-6">
        <div className="flex items-center justify-between mb-6">
          <h2 className="text-xl font-semibold text-white">{t('dashboard.recentActivity')}</h2>
          <Activity className="h-5 w-5 text-white/40" />
        </div>
        <div className="space-y-4">
          {recentActivity.map((activity, index) => (
            <div key={index} className="flex items-start space-x-4 p-3 rounded-xl hover:bg-white/5 transition-colors">
              <div className="w-2 h-2 bg-gradient-to-r from-blue-500 to-purple-500 rounded-full mt-2 animate-pulse"></div>
              <div className="flex-1">
                <p className="text-white text-sm">
                  <span className="font-medium text-blue-400">{activity.user}</span> {activity.action}
                </p>
                <p className="text-white/60 text-xs mt-1">{activity.time}</p>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

export default Dashboard;