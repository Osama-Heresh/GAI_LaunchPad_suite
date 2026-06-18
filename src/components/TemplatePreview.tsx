import React, { useState } from 'react';
import { X, Download, CreditCard as Edit, Eye, FileText, Calendar, Star, Users } from 'lucide-react';
import { DocumentTemplate } from '../services/documentService';

interface TemplatePreviewProps {
  template: DocumentTemplate;
  onClose: () => void;
  onDownload: (template: DocumentTemplate, variables?: Record<string, string>) => void;
  onEdit?: (template: DocumentTemplate) => void;
  isAdmin?: boolean;
}

const TemplatePreview: React.FC<TemplatePreviewProps> = ({
  template,
  onClose,
  onDownload,
  onEdit,
  isAdmin = false
}) => {
  const [variables, setVariables] = useState<Record<string, string>>({});
  const [showVariables, setShowVariables] = useState(false);

  const handleVariableChange = (variable: string, value: string) => {
    setVariables(prev => ({
      ...prev,
      [variable]: value
    }));
  };

  const handleDownload = () => {
    onDownload(template, variables);
  };

  const getCategoryColor = (category: string) => {
    const colors = {
      hr: 'bg-blue-100 text-blue-800',
      sales: 'bg-green-100 text-green-800',
      marketing: 'bg-purple-100 text-purple-800',
      legal: 'bg-red-100 text-red-800',
      general: 'bg-gray-100 text-gray-800'
    };
    return colors[category as keyof typeof colors] || colors.general;
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl max-w-4xl w-full max-h-[90vh] overflow-hidden shadow-2xl">
        {/* Header */}
        <div className="flex items-center justify-between p-6 border-b border-gray-200 bg-gradient-to-r from-blue-50 to-purple-50">
          <div className="flex items-center space-x-4">
            <div className="p-3 bg-white rounded-lg shadow-sm">
              <FileText className="h-6 w-6 text-blue-600" />
            </div>
            <div>
              <h2 className="text-xl font-semibold text-gray-900">{template.title}</h2>
              <div className="flex items-center space-x-3 mt-1">
                <span className={`px-2 py-1 text-xs rounded-full font-medium ${getCategoryColor(template.category)}`}>
                  {template.category.toUpperCase()}
                </span>
                <div className="flex items-center space-x-1 text-sm text-gray-600">
                  <Star className="h-4 w-4 text-yellow-400 fill-current" />
                  <span>{template.rating}</span>
                </div>
                <div className="flex items-center space-x-1 text-sm text-gray-600">
                  <Users className="h-4 w-4" />
                  <span>{template.downloadCount.toLocaleString()} downloads</span>
                </div>
              </div>
            </div>
          </div>
          
          <div className="flex items-center space-x-2">
            {isAdmin && onEdit && (
              <button
                onClick={() => onEdit(template)}
                className="flex items-center space-x-2 px-4 py-2 text-blue-600 border border-blue-200 rounded-lg hover:bg-blue-50 transition-colors"
              >
                <Edit className="h-4 w-4" />
                <span>Edit</span>
              </button>
            )}
            
            <button
              onClick={() => setShowVariables(!showVariables)}
              className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <Download className="h-4 w-4" />
              <span>Customize & Download</span>
            </button>
            
            <button
              onClick={onClose}
              className="p-2 text-gray-400 hover:text-gray-600 transition-colors"
            >
              <X className="h-5 w-5" />
            </button>
          </div>
        </div>

        <div className="flex h-[calc(90vh-120px)]">
          {/* Main Content */}
          <div className="flex-1 p-6 overflow-y-auto">
            {/* Description */}
            <div className="mb-6">
              <h3 className="text-lg font-medium text-gray-900 mb-2">Description</h3>
              <p className="text-gray-600 leading-relaxed">{template.description}</p>
            </div>

            {/* Tags */}
            <div className="mb-6">
              <h3 className="text-lg font-medium text-gray-900 mb-2">Tags</h3>
              <div className="flex flex-wrap gap-2">
                {template.tags.map((tag, index) => (
                  <span
                    key={index}
                    className="px-3 py-1 bg-gray-100 text-gray-700 text-sm rounded-full"
                  >
                    {tag}
                  </span>
                ))}
              </div>
            </div>

            {/* Template Info */}
            <div className="mb-6">
              <h3 className="text-lg font-medium text-gray-900 mb-3">Template Information</h3>
              <div className="grid grid-cols-2 gap-4">
                <div className="p-4 bg-gray-50 rounded-lg">
                  <div className="flex items-center space-x-2 mb-2">
                    <Calendar className="h-4 w-4 text-gray-500" />
                    <span className="text-sm font-medium text-gray-700">Last Updated</span>
                  </div>
                  <span className="text-gray-900">{new Date(template.updatedAt).toLocaleDateString()}</span>
                </div>
                
                <div className="p-4 bg-gray-50 rounded-lg">
                  <div className="flex items-center space-x-2 mb-2">
                    <FileText className="h-4 w-4 text-gray-500" />
                    <span className="text-sm font-medium text-gray-700">File Size</span>
                  </div>
                  <span className="text-gray-900">
                    {template.fileSize ? formatFileSize(template.fileSize) : 'N/A'}
                  </span>
                </div>
              </div>
            </div>

            {/* Sections Preview */}
            <div className="mb-6">
              <h3 className="text-lg font-medium text-gray-900 mb-3">Document Sections</h3>
              {template.preview?.sections && template.preview.sections.length > 0 ? (
                <div className="space-y-2">
                  {template.preview.sections.map((section, index) => (
                    <div key={index} className="flex items-center space-x-3 p-3 bg-gray-50 rounded-lg">
                      <div className="w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-sm font-medium">
                        {index + 1}
                      </div>
                      <span className="text-gray-900">{section}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="p-4 bg-gray-50 rounded-lg text-gray-600 text-sm">
                  Professional business document with customizable sections
                </div>
              )}
            </div>

            {/* Variables */}
            {template.preview?.variables && template.preview.variables.length > 0 && (
              <div>
                <h3 className="text-lg font-medium text-gray-900 mb-3">Customizable Fields</h3>
                <div className="grid grid-cols-2 gap-3">
                  {template.preview.variables.map((variable, index) => (
                    <div key={index} className="p-3 bg-blue-50 rounded-lg border border-blue-200">
                      <span className="text-blue-800 font-mono text-sm">{variable}</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Variables Panel */}
          {showVariables && (
            <div className="w-80 border-l border-gray-200 bg-gray-50 p-6 overflow-y-auto">
              <h3 className="text-lg font-medium text-gray-900 mb-4">Customize Template</h3>
              <p className="text-sm text-gray-600 mb-4">
                Fill in the fields below to customize your document before downloading.
              </p>
              
              <div className="space-y-4">
                {template.preview?.variables && template.preview.variables.length > 0 ? (
                  template.preview.variables.map((variable, index) => {
                    const cleanVariable = variable.replace(/[\[\]]/g, '');
                    return (
                      <div key={index}>
                        <label className="block text-sm font-medium text-gray-700 mb-1">
                          {cleanVariable.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase())}
                        </label>
                        <input
                          type="text"
                          value={variables[variable] || ''}
                          onChange={(e) => handleVariableChange(variable, e.target.value)}
                          placeholder={`Enter ${cleanVariable.toLowerCase()}`}
                          className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                      </div>
                    );
                  })
                ) : (
                  <p className="text-sm text-gray-600 py-4">
                    This template is ready to download as-is. No additional customization needed.
                  </p>
                )}
              </div>

              <div className="mt-6 pt-4 border-t border-gray-200">
                <button
                  onClick={handleDownload}
                  className="w-full flex items-center justify-center space-x-2 px-4 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium"
                >
                  <Download className="h-4 w-4" />
                  <span>Download DOCX</span>
                </button>
                
                <p className="text-xs text-gray-500 mt-2 text-center">
                  Document will be generated as a Microsoft Word (.docx) file
                </p>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default TemplatePreview;