import React, { useState } from 'react';
import { PlayCircle, BookOpen, Download, Clock, Users, Star } from 'lucide-react';
import { useLanguage } from '../contexts/LanguageContext';
import jsPDF from 'jspdf';

const Training: React.FC = () => {
  const { t } = useLanguage();
  const [selectedVideo, setSelectedVideo] = useState<string | null>(null);

  const videos = [
    {
      id: '1',
      title: 'Getting Started with LaunchPad Suite',
      description: 'Learn the basics of our platform and how to set up your account',
      duration: '5:30',
      views: 1250,
      rating: 4.8,
      thumbnail: 'https://images.pexels.com/photos/3184418/pexels-photo-3184418.jpeg?auto=compress&cs=tinysrgb&w=800'
    },
    {
      id: '2',
      title: 'Social Media Management Tutorial',
      description: 'Master the art of managing multiple social media platforms',
      duration: '12:45',
      views: 890,
      rating: 4.9,
      thumbnail: 'https://images.pexels.com/photos/267350/pexels-photo-267350.jpeg?auto=compress&cs=tinysrgb&w=800'
    },
    {
      id: '3',
      title: 'Task Management with Kanban Boards',
      description: 'Organize your work efficiently using our task management system',
      duration: '8:20',
      views: 2100,
      rating: 4.7,
      thumbnail: 'https://images.pexels.com/photos/3184465/pexels-photo-3184465.jpeg?auto=compress&cs=tinysrgb&w=800'
    },
    {
      id: '4',
      title: 'Business Letter Templates Guide',
      description: 'How to use and customize our professional business templates',
      duration: '6:15',
      views: 670,
      rating: 4.6,
      thumbnail: 'https://images.pexels.com/photos/3184317/pexels-photo-3184317.jpeg?auto=compress&cs=tinysrgb&w=800'
    }
  ];

  const generateMarketingPDF = () => {
    const doc = new jsPDF();
    
    // Header
    doc.setFontSize(24);
    doc.setTextColor(37, 99, 235); // Blue color
    doc.text('LaunchPad Suite', 20, 30);
    doc.setFontSize(16);
    doc.setTextColor(0, 0, 0);
    doc.text('All-in-One Business Management Platform', 20, 40);
    
    // Separator
    doc.setLineWidth(0.5);
    doc.line(20, 45, 190, 45);
    
    // Introduction
    doc.setFontSize(14);
    doc.text('Transform Your Business Operations', 20, 60);
    doc.setFontSize(10);
    const introText = 'LaunchPad Suite is a comprehensive SaaS platform designed to streamline operations for new and growing companies. Our integrated solution combines social media management, task tracking, and professional business templates in one powerful platform.';
    doc.text(doc.splitTextToSize(introText, 170), 20, 70);
    
    // Core Features
    doc.setFontSize(14);
    doc.text('Core Features', 20, 100);
    
    const features = [
      '✓ Social Media Management - Post to multiple platforms simultaneously',
      '✓ Kanban Task Management - Organize and track project progress',
      '✓ Business Letter Templates - Professional templates for HR, Sales, and Marketing',
      '✓ User Management System - Control access and monitor subscription status',
      '✓ Multilingual Support - Available in English and Arabic',
      '✓ Training Resources - Comprehensive video tutorials and documentation'
    ];
    
    let yPosition = 110;
    features.forEach(feature => {
      doc.setFontSize(10);
      doc.text(feature, 25, yPosition);
      yPosition += 8;
    });
    
    // Benefits
    doc.setFontSize(14);
    doc.text('Key Benefits', 20, 170);
    
    const benefits = [
      '• Increase productivity by up to 40% with integrated workflow management',
      '• Save 15+ hours per week on social media management',
      '• Professional communication with ready-to-use business templates',
      '• Scalable solution that grows with your business',
      '• Affordable monthly subscription starting at $29/month'
    ];
    
    yPosition = 180;
    benefits.forEach(benefit => {
      doc.setFontSize(10);
      doc.text(benefit, 25, yPosition);
      yPosition += 8;
    });
    
    // Pricing
    doc.setFontSize(14);
    doc.text('Subscription Plans', 20, 230);
    
    doc.setFontSize(12);
    doc.text('Basic Plan - $19/month', 25, 245);
    doc.setFontSize(10);
    doc.text('• Social Media Management • Task Management • Basic Templates', 30, 252);
    
    doc.setFontSize(12);
    doc.text('Pro Plan - $29/month', 25, 265);
    doc.setFontSize(10);
    doc.text('• All Basic features • Advanced Templates • User Management • Priority Support', 30, 272);
    
    // Footer
    doc.setFontSize(8);
    doc.setTextColor(128, 128, 128);
    doc.text('For more information, visit www.launchpadsuite.com or contact sales@launchpadsuite.com', 20, 285);
    
    // Save the PDF
    doc.save('LaunchPad-Suite-Marketing-Brochure.pdf');
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold text-gray-900">{t('nav.training')}</h1>
        <button
          onClick={generateMarketingPDF}
          className="flex items-center space-x-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700"
        >
          <Download className="h-4 w-4" />
          <span>Download Marketing PDF</span>
        </button>
      </div>

      {/* Training Overview */}
      <div className="bg-gradient-to-r from-blue-600 to-purple-600 rounded-xl p-6 text-white">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-2xl font-bold mb-2">Master LaunchPad Suite</h2>
            <p className="text-blue-100">
              Comprehensive training materials to help you get the most out of our platform
            </p>
          </div>
          <BookOpen className="h-16 w-16 text-blue-200" />
        </div>
      </div>

      {/* Video Tutorials */}
      <div className="bg-white rounded-xl p-6 shadow-sm">
        <h2 className="text-lg font-semibold text-gray-900 mb-6">Video Tutorials</h2>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {videos.map((video) => (
            <div key={video.id} className="bg-gray-50 rounded-lg overflow-hidden hover:shadow-md transition-shadow">
              <div className="relative">
                <img
                  src={video.thumbnail}
                  alt={video.title}
                  className="w-full h-48 object-cover"
                />
                <button
                  onClick={() => setSelectedVideo(video.id)}
                  className="absolute inset-0 flex items-center justify-center bg-black bg-opacity-40 hover:bg-opacity-50 transition-opacity"
                >
                  <PlayCircle className="h-16 w-16 text-white" />
                </button>
                <div className="absolute bottom-2 right-2 bg-black bg-opacity-75 text-white px-2 py-1 rounded text-sm">
                  {video.duration}
                </div>
              </div>
              
              <div className="p-4">
                <h3 className="font-semibold text-gray-900 mb-2">{video.title}</h3>
                <p className="text-sm text-gray-600 mb-3">{video.description}</p>
                
                <div className="flex items-center justify-between text-sm text-gray-500">
                  <div className="flex items-center space-x-4">
                    <div className="flex items-center space-x-1">
                      <Users className="h-4 w-4" />
                      <span>{video.views.toLocaleString()}</span>
                    </div>
                    <div className="flex items-center space-x-1">
                      <Star className="h-4 w-4 text-yellow-400 fill-current" />
                      <span>{video.rating}</span>
                    </div>
                  </div>
                  <div className="flex items-center space-x-1">
                    <Clock className="h-4 w-4" />
                    <span>{video.duration}</span>
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Quick Start Guide */}
      <div className="bg-white rounded-xl p-6 shadow-sm">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">Quick Start Guide</h2>
        
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="text-center p-4 border border-gray-200 rounded-lg">
            <div className="w-12 h-12 bg-blue-100 rounded-full flex items-center justify-center mx-auto mb-3">
              <span className="text-blue-600 font-bold">1</span>
            </div>
            <h3 className="font-medium text-gray-900 mb-2">Set Up Your Account</h3>
            <p className="text-sm text-gray-600">Complete your profile and configure basic settings</p>
          </div>
          
          <div className="text-center p-4 border border-gray-200 rounded-lg">
            <div className="w-12 h-12 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-3">
              <span className="text-green-600 font-bold">2</span>
            </div>
            <h3 className="font-medium text-gray-900 mb-2">Connect Your Platforms</h3>
            <p className="text-sm text-gray-600">Link your social media accounts and integrate tools</p>
          </div>
          
          <div className="text-center p-4 border border-gray-200 rounded-lg">
            <div className="w-12 h-12 bg-purple-100 rounded-full flex items-center justify-center mx-auto mb-3">
              <span className="text-purple-600 font-bold">3</span>
            </div>
            <h3 className="font-medium text-gray-900 mb-2">Start Managing</h3>
            <p className="text-sm text-gray-600">Begin creating posts, tasks, and using templates</p>
          </div>
        </div>
      </div>

      {/* Documentation Links */}
      <div className="bg-white rounded-xl p-6 shadow-sm">
        <h2 className="text-lg font-semibold text-gray-900 mb-4">Documentation</h2>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <a href="#" className="flex items-center space-x-3 p-3 border border-gray-200 rounded-lg hover:bg-gray-50">
            <BookOpen className="h-5 w-5 text-blue-600" />
            <div>
              <h3 className="font-medium text-gray-900">User Manual</h3>
              <p className="text-sm text-gray-600">Complete guide to all features</p>
            </div>
          </a>
          
          <a href="#" className="flex items-center space-x-3 p-3 border border-gray-200 rounded-lg hover:bg-gray-50">
            <BookOpen className="h-5 w-5 text-green-600" />
            <div>
              <h3 className="font-medium text-gray-900">API Documentation</h3>
              <p className="text-sm text-gray-600">For developers and integrations</p>
            </div>
          </a>
        </div>
      </div>

      {/* Video Modal */}
      {selectedVideo && (
        <div className="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-50">
          <div className="bg-white rounded-xl max-w-4xl w-full mx-4 overflow-hidden">
            <div className="relative">
              <div className="aspect-video bg-gray-900 flex items-center justify-center">
                <div className="text-center text-white">
                  <PlayCircle className="h-24 w-24 mx-auto mb-4" />
                  <p className="text-lg">Video Player Placeholder</p>
                  <p className="text-sm opacity-75">In a real implementation, this would show the actual video</p>
                </div>
              </div>
              <button
                onClick={() => setSelectedVideo(null)}
                className="absolute top-4 right-4 text-white hover:text-gray-300"
              >
                <span className="text-2xl">×</span>
              </button>
            </div>
            
            <div className="p-6">
              {videos.find(v => v.id === selectedVideo) && (
                <div>
                  <h3 className="text-lg font-semibold text-gray-900 mb-2">
                    {videos.find(v => v.id === selectedVideo)?.title}
                  </h3>
                  <p className="text-gray-600">
                    {videos.find(v => v.id === selectedVideo)?.description}
                  </p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default Training;