import React, { createContext, useContext, useState, ReactNode } from 'react';

type Language = 'en' | 'ar';

interface LanguageContextType {
  language: Language;
  setLanguage: (lang: Language) => void;
  t: (key: string) => string;
  isRTL: boolean;
}

const translations = {
  en: {
    // Navigation
    'nav.dashboard': 'Dashboard',
    'nav.socialMedia': 'Social Media',
    'nav.tasks': 'Task Manager',
    'nav.templates': 'Templates',
    'nav.users': 'User Management',
    'nav.settings': 'Settings',
    'nav.training': 'Training',
    'nav.logout': 'Logout',
    
    // Dashboard
    'dashboard.welcome': 'Welcome to BusinessHub',
    'dashboard.subtitle': 'Streamline your business operations with our comprehensive platform',
    'dashboard.activeUsers': 'Active Users',
    'dashboard.pendingTasks': 'Pending Tasks',
    'dashboard.scheduledPosts': 'Scheduled Posts',
    'dashboard.templates': 'Templates Used',
    'dashboard.recentActivity': 'Recent Activity',
    'dashboard.quickActions': 'Quick Actions',
    
    // Social Media
    'social.title': 'Social Media Management',
    'social.compose': 'Compose Post',
    'social.schedule': 'Schedule',
    'social.platforms': 'Select Platforms',
    'social.content': 'Post Content',
    'social.scheduledPosts': 'Scheduled Posts',
    
    // Tasks
    'tasks.title': 'Task Manager',
    'tasks.todo': 'To Do',
    'tasks.inProgress': 'In Progress',
    'tasks.completed': 'Completed',
    'tasks.addTask': 'Add Task',
    'tasks.newTask': 'New Task',
    
    // Templates
    'templates.title': 'Business Letter Templates',
    'templates.hr': 'HR Templates',
    'templates.sales': 'Sales Templates',
    'templates.marketing': 'Marketing Templates',
    'templates.general': 'General Templates',
    
    // Users
    'users.title': 'User Management',
    'users.addUser': 'Add User',
    'users.status': 'Status',
    'users.subscription': 'Subscription',
    'users.actions': 'Actions',
    
    // Common
    'common.save': 'Save',
    'common.cancel': 'Cancel',
    'common.edit': 'Edit',
    'common.delete': 'Delete',
    'common.view': 'View',
    'common.active': 'Active',
    'common.inactive': 'Inactive',
    'common.loading': 'Loading...',
    'common.search': 'Search...',
    
    // Login
    'login.title': 'Welcome to BusinessHub',
    'login.subtitle': 'Sign in to your account',
    'login.email': 'Email Address',
    'login.password': 'Password',
    'login.signIn': 'Sign In',
    'login.demo': 'Demo Credentials',
    'login.adminDemo': 'Admin: admin@businesshub.com / admin',
    'login.userDemo': 'User: user@company.com / user'
  },
  ar: {
    // Navigation
    'nav.dashboard': 'لوحة التحكم',
    'nav.socialMedia': 'وسائل التواصل',
    'nav.tasks': 'إدارة المهام',
    'nav.templates': 'القوالب',
    'nav.users': 'إدارة المستخدمين',
    'nav.settings': 'الإعدادات',
    'nav.training': 'التدريب',
    'nav.logout': 'تسجيل الخروج',
    
    // Dashboard
    'dashboard.welcome': 'مرحباً بك في بيزنس هب',
    'dashboard.subtitle': 'قم بتبسيط عمليات أعمالك مع منصتنا الشاملة',
    'dashboard.activeUsers': 'المستخدمون النشطون',
    'dashboard.pendingTasks': 'المهام المعلقة',
    'dashboard.scheduledPosts': 'المنشورات المجدولة',
    'dashboard.templates': 'القوالب المستخدمة',
    'dashboard.recentActivity': 'النشاط الأخير',
    'dashboard.quickActions': 'إجراءات سريعة',
    
    // Social Media
    'social.title': 'إدارة وسائل التواصل الاجتماعي',
    'social.compose': 'إنشاء منشور',
    'social.schedule': 'جدولة',
    'social.platforms': 'اختر المنصات',
    'social.content': 'محتوى المنشور',
    'social.scheduledPosts': 'المنشورات المجدولة',
    
    // Tasks
    'tasks.title': 'مدير المهام',
    'tasks.todo': 'للقيام',
    'tasks.inProgress': 'قيد التنفيذ',
    'tasks.completed': 'مكتملة',
    'tasks.addTask': 'إضافة مهمة',
    'tasks.newTask': 'مهمة جديدة',
    
    // Templates
    'templates.title': 'قوالب رسائل الأعمال',
    'templates.hr': 'قوالب الموارد البشرية',
    'templates.sales': 'قوالب المبيعات',
    'templates.marketing': 'قوالب التسويق',
    'templates.general': 'قوالب عامة',
    
    // Users
    'users.title': 'إدارة المستخدمين',
    'users.addUser': 'إضافة مستخدم',
    'users.status': 'الحالة',
    'users.subscription': 'الاشتراك',
    'users.actions': 'الإجراءات',
    
    // Common
    'common.save': 'حفظ',
    'common.cancel': 'إلغاء',
    'common.edit': 'تعديل',
    'common.delete': 'حذف',
    'common.view': 'عرض',
    'common.active': 'نشط',
    'common.inactive': 'غير نشط',
    'common.loading': 'جارٍ التحميل...',
    'common.search': 'بحث...',
    
    // Login
    'login.title': 'مرحباً بك في بيزنس هب',
    'login.subtitle': 'تسجيل الدخول إلى حسابك',
    'login.email': 'عنوان البريد الإلكتروني',
    'login.password': 'كلمة المرور',
    'login.signIn': 'تسجيل الدخول',
    'login.demo': 'بيانات التجربة',
    'login.adminDemo': 'مدير: admin@businesshub.com / admin',
    'login.userDemo': 'مستخدم: user@company.com / user'
  }
};

const LanguageContext = createContext<LanguageContextType | undefined>(undefined);

export const useLanguage = () => {
  const context = useContext(LanguageContext);
  if (context === undefined) {
    throw new Error('useLanguage must be used within a LanguageProvider');
  }
  return context;
};

export const LanguageProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [language, setLanguage] = useState<Language>('en');

  const t = (key: string): string => {
    return translations[language][key as keyof typeof translations[typeof language]] || key;
  };

  const isRTL = language === 'ar';

  return (
    <LanguageContext.Provider value={{ language, setLanguage, t, isRTL }}>
      <div dir={isRTL ? 'rtl' : 'ltr'} className={isRTL ? 'font-arabic' : ''}>
        {children}
      </div>
    </LanguageContext.Provider>
  );
};