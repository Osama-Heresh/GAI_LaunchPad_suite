import React, { useState, useMemo, useCallback, useEffect } from 'react';
import { Search, Download, Eye, CreditCard as Edit, Trash2, Plus, Filter } from 'lucide-react';
import { documentService } from '../services/documentService';
import TemplatePreview from '../components/TemplatePreview';
import TemplateUpload from '../components/TemplateUpload';

const Templates: React.FC = () => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [sortBy, setSortBy] = useState<'name' | 'downloads' | 'rating' | 'date'>('name');
  const [previewTemplate, setPreviewTemplate] = useState<any>(null);
  const [showUpload, setShowUpload] = useState(false);
  const [isAdmin] = useState(true); // In real app, get from auth context
  const [isLoading, setIsLoading] = useState(true);

  const categories = useMemo(() => [
    { id: 'all', name: 'All Templates', color: 'from-gray-500 to-gray-600', count: 0 },
    { id: 'hr', name: 'Human Resources (HR)', color: 'from-blue-500 to-blue-600', count: 0 },
    { id: 'finance', name: 'Finance & Accounting', color: 'from-green-500 to-green-600', count: 0 },
    { id: 'sales', name: 'Sales', color: 'from-purple-500 to-purple-600', count: 0 },
    { id: 'marketing', name: 'Marketing', color: 'from-pink-500 to-pink-600', count: 0 },
    { id: 'legal', name: 'Legal & Admin', color: 'from-red-500 to-red-600', count: 0 },
    { id: 'it', name: 'IT & Operations', color: 'from-cyan-500 to-cyan-600', count: 0 },
    { id: 'project', name: 'Client & Project Management', color: 'from-indigo-500 to-indigo-600', count: 0 },
    { id: 'strategy', name: 'Business Strategy & Planning', color: 'from-orange-500 to-orange-600', count: 0 }
  ], []);

  const templates = useMemo(() => documentService.getTemplates(), []);

  useEffect(() => {
    // Simulate loading delay and set loading to false once templates are ready
    const timer = setTimeout(() => setIsLoading(false), 100);
    return () => clearTimeout(timer);
  }, [templates]);

  // Update category counts
  const categoriesWithCounts = useMemo(() =>
    categories.map(category => ({
      ...category,
      count: category.id === 'all'
        ? templates.length
        : templates.filter(t => t.category === category.id).length
    })),
    [categories, templates]
  );

  const filteredTemplates = useMemo(() => {
    let filtered = templates;

    // Filter by category
    if (selectedCategory !== 'all') {
      filtered = filtered.filter(template => template.category === selectedCategory);
    }

    // Filter by search term
    if (searchTerm) {
      filtered = filtered.filter(template =>
        template.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        template.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
        template.tags.some(tag => tag.toLowerCase().includes(searchTerm.toLowerCase()))
      );
    }

    // Sort templates
    filtered.sort((a, b) => {
      switch (sortBy) {
        case 'downloads':
          return b.downloadCount - a.downloadCount;
        case 'rating':
          return b.rating - a.rating;
        case 'date':
          return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
        default:
          return a.title.localeCompare(b.title);
      }
    });

    return filtered;
  }, [templates, selectedCategory, searchTerm, sortBy]);

  const handleDownload = useCallback(async (template: any, variables?: Record<string, string>) => {
    try {
      await documentService.generateDocx(template, variables || {});
      // Update download count
      template.downloadCount += 1;
    } catch (error) {
      console.error('Download failed:', error);
    }
  }, []);

  const handleDelete = useCallback((templateId: string) => {
    if (window.confirm('Are you sure you want to delete this template?')) {
      // In real app, call API to delete template
      console.log('Delete template:', templateId);
    }
  }, []);

  const getCategoryInfo = useCallback((categoryId: string) => {
    return categoriesWithCounts.find(cat => cat.id === categoryId) || categoriesWithCounts[0];
  }, [categoriesWithCounts]);

  if (isLoading) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center">
        <div className="text-center">
          <div className="inline-block animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
          <p className="text-gray-600">Loading templates...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Header */}
        <div className="mb-8">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-3xl font-bold text-gray-900">Document Templates</h1>
              <p className="mt-2 text-gray-600">Professional business document templates ready for download</p>
            </div>
            {isAdmin && (
              <button
                onClick={() => setShowUpload(true)}
                className="inline-flex items-center px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              >
                <Plus className="h-4 w-4 mr-2" />
                Add Template
              </button>
            )}
          </div>
        </div>

        {/* Search and Filters */}
        <div className="mb-8 space-y-4">
          <div className="flex flex-col sm:flex-row gap-4">
            <div className="flex-1 relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 h-4 w-4" />
              <input
                type="text"
                placeholder="Search templates..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
            </div>
            <div className="flex gap-2">
              <select
                value={sortBy}
                onChange={(e) => setSortBy(e.target.value as any)}
                className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              >
                <option value="name">Sort by Name</option>
                <option value="downloads">Sort by Downloads</option>
                <option value="rating">Sort by Rating</option>
                <option value="date">Sort by Date</option>
              </select>
            </div>
          </div>
        </div>

        <div className="flex gap-8">
          {/* Categories Sidebar */}
          <div className="w-80 flex-shrink-0">
            <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
              <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                <Filter className="h-5 w-5 mr-2" />
                Categories
              </h3>
              <div className="space-y-2">
                {categoriesWithCounts.map((category) => (
                  <button
                    key={category.id}
                    onClick={() => setSelectedCategory(category.id)}
                    className={`w-full text-left px-3 py-2 rounded-lg transition-colors flex items-center justify-between ${
                      selectedCategory === category.id
                        ? 'bg-blue-50 text-blue-700 border border-blue-200'
                        : 'hover:bg-gray-50 text-gray-700'
                    }`}
                  >
                    <span className="text-base">{category.name}</span>
                    <span className={`text-xs px-2 py-1 rounded-full ${
                      selectedCategory === category.id
                        ? 'bg-blue-100 text-blue-700'
                        : 'bg-gray-100 text-gray-600'
                    }`}>
                      {category.count}
                    </span>
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Templates Grid */}
          <div className="flex-1">
            {filteredTemplates.length === 0 ? (
              <div className="text-center py-12">
                <div className="text-gray-400 mb-4">
                  <Search className="h-12 w-12 mx-auto" />
                </div>
                <h3 className="text-lg font-medium text-gray-900 mb-2">No templates found</h3>
                <p className="text-gray-600">Try adjusting your search or filter criteria</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {filteredTemplates.map((template) => {
                  const categoryInfo = getCategoryInfo(template.category);
                  return (
                    <div key={template.id} className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden hover:shadow-md transition-shadow">
                      {/* Category Badge */}
                      <div className={`h-2 bg-gradient-to-r ${categoryInfo.color}`}></div>
                      
                      <div className="p-6">
                        {/* Header */}
                        <div className="flex items-start justify-between mb-3">
                          <div className={`inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-gradient-to-r ${categoryInfo.color} text-white`}>
                            {categoryInfo.name.split(' ')[0]}
                          </div>
                          <div className="flex items-center text-xs text-yellow-500">
                            <span className="mr-1">★</span>
                            {template.rating}
                          </div>
                        </div>

                        {/* Content */}
                        <h3 className="text-lg font-semibold text-gray-900 mb-2 line-clamp-2">
                          {template.title}
                        </h3>
                        <p className="text-sm text-gray-600 mb-4 line-clamp-3">
                          {template.description}
                        </p>

                        {/* Tags */}
                        <div className="flex flex-wrap gap-1 mb-4">
                          {template.tags.slice(0, 3).map((tag, index) => (
                            <span
                              key={index}
                              className="inline-flex items-center px-2 py-1 rounded-full text-xs bg-gray-100 text-gray-700"
                            >
                              {tag}
                            </span>
                          ))}
                          {template.tags.length > 3 && (
                            <span className="text-xs text-gray-500">+{template.tags.length - 3} more</span>
                          )}
                        </div>

                        {/* Stats */}
                        <div className="flex items-center justify-between text-xs text-gray-500 mb-4">
                          <span>{template.downloadCount.toLocaleString()} downloads</span>
                          <span>{template.fileSize}</span>
                        </div>

                        {/* Actions */}
                        <div className="flex items-center justify-between space-x-1">
                          <button
                            onClick={() => setPreviewTemplate(template)}
                            className="flex items-center px-2 py-1.5 text-xs text-gray-600 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors"
                          >
                            <Eye className="h-3 w-3 mr-1" />
                            Preview
                          </button>
                          <button
                            onClick={() => handleDownload(template)}
                            className="flex items-center px-2 py-1.5 text-xs text-white bg-blue-600 hover:bg-blue-700 rounded transition-colors"
                          >
                            <Download className="h-3 w-3 mr-1" />
                            Download
                          </button>
                          {isAdmin && (
                            <>
                              <button
                                onClick={() => console.log('Edit template:', template.id)}
                                className="flex items-center px-2 py-1.5 text-xs text-orange-600 hover:text-orange-700 hover:bg-orange-50 rounded transition-colors"
                                title="Edit Template"
                              >
                                <Edit className="h-3 w-3" />
                              </button>
                              <button
                                onClick={() => handleDelete(template.id)}
                                className="flex items-center px-2 py-1.5 text-xs text-red-600 hover:text-red-700 hover:bg-red-50 rounded transition-colors"
                                title="Delete Template"
                              >
                                <Trash2 className="h-3 w-3" />
                              </button>
                            </>
                          )}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Modals */}
      {previewTemplate && (
        <TemplatePreview
          template={previewTemplate}
          onClose={() => setPreviewTemplate(null)}
          onDownload={() => handleDownload(previewTemplate)}
        />
      )}

      {showUpload && (
        <TemplateUpload
          onClose={() => setShowUpload(false)}
          onUpload={(template) => {
            console.log('Upload template:', template);
            setShowUpload(false);
          }}
        />
      )}
    </div>
  );
};

export default Templates;