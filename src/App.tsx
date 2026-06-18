import React from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import { AuthProvider } from './contexts/AuthContext';
import { LanguageProvider } from './contexts/LanguageContext';
import { SubscriptionProvider } from './contexts/SubscriptionContext';
import { SocialAccountsProvider } from './contexts/SocialAccountsContext';
import ProtectedRoute from './components/ProtectedRoute';
import Layout from './components/Layout';
import Login from './pages/Login';
import Dashboard from './pages/Dashboard';
import SocialMedia from './pages/SocialMedia';
import TaskManager from './pages/TaskManager';
import Templates from './pages/Templates';
import UserManagement from './pages/UserManagement';
import Training from './pages/Training';
import Settings from './pages/Settings';
import TeamManagement from './pages/TeamManagement';

function App() {
  return (
    <AuthProvider>
      <LanguageProvider>
        <SubscriptionProvider>
          <SocialAccountsProvider>
            <Router>
              <Routes>
                <Route path="/login" element={<Login />} />
                <Route path="/" element={
                  <ProtectedRoute>
                    <Layout />
                  </ProtectedRoute>
                }>
                  <Route index element={<Navigate to="/dashboard" replace />} />
                  <Route path="dashboard" element={<Dashboard />} />
                  <Route path="social-media" element={<SocialMedia />} />
                  <Route path="tasks" element={<TaskManager />} />
                  <Route path="templates" element={<Templates />} />
                  <Route path="users" element={<UserManagement />} />
                  <Route path="team" element={<TeamManagement />} />
                  <Route path="training" element={<Training />} />
                  <Route path="settings" element={<Settings />} />
                </Route>
              </Routes>
            </Router>
          </SocialAccountsProvider>
        </SubscriptionProvider>
      </LanguageProvider>
    </AuthProvider>
  );
}

export default App;