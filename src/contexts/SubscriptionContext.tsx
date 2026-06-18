import React, { createContext, useContext, useState, ReactNode } from 'react';

interface SubscriptionStatus {
  isActive: boolean;
  plan: 'basic' | 'pro' | 'enterprise';
  expiryDate: Date;
  features: string[];
}

interface SubscriptionContextType {
  subscription: SubscriptionStatus;
  hasAccess: (feature: string) => boolean;
  updateSubscription: (status: SubscriptionStatus) => void;
}

const SubscriptionContext = createContext<SubscriptionContextType | undefined>(undefined);

export const useSubscription = () => {
  const context = useContext(SubscriptionContext);
  if (context === undefined) {
    throw new Error('useSubscription must be used within a SubscriptionProvider');
  }
  return context;
};

export const SubscriptionProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [subscription, setSubscription] = useState<SubscriptionStatus>({
    isActive: true,
    plan: 'pro',
    expiryDate: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000), // 30 days from now
    features: ['social-media', 'tasks', 'templates', 'user-management']
  });

  const hasAccess = (feature: string): boolean => {
    return subscription.isActive && subscription.features.includes(feature);
  };

  const updateSubscription = (status: SubscriptionStatus) => {
    setSubscription(status);
  };

  return (
    <SubscriptionContext.Provider value={{ subscription, hasAccess, updateSubscription }}>
      {children}
    </SubscriptionContext.Provider>
  );
};