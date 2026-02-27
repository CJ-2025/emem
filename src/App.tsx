/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Printer, FileSpreadsheet, Trash2, Plus, Download, Image as ImageIcon, Settings2, Move, Check, Type, Search, FileText } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import Draggable from 'react-draggable';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

interface FieldConfig {
  id: string;
  label: string;
  x: number;
  y: number;
  fontSize: number;
  lineHeight: number;
  rotation: number;
  color: string;
  fontWeight: string;
  width: number;
  textAlign: 'left' | 'center' | 'right';
  mapping?: string;
}

interface PriceTagData {
  id: string;
  selected: boolean;
  rawData: Record<string, any>;
}

const INITIAL_FIELDS: FieldConfig[] = [];

interface DraggableFieldProps {
  field: FieldConfig;
  selectedFieldId: string | null;
  setSelectedFieldId: (id: string | null) => void;
  updateField: (id: string, updates: Partial<FieldConfig>) => void;
  previewTag: PriceTagData | undefined;
}

const DraggableField: React.FC<DraggableFieldProps> = ({ 
  field, 
  selectedFieldId, 
  setSelectedFieldId, 
  updateField, 
  previewTag 
}) => {
  const nodeRef = useRef(null);
  
  const getFieldValue = () => {
    let rawValue = '';
    if (field.mapping && previewTag?.rawData) {
      rawValue = String(previewTag.rawData[field.mapping] || '');
    } else {
      return field.label;
    }

    if (field.label.toLowerCase().includes('srp')) return formatCurrency(rawValue);
    if (field.label.toLowerCase().includes('downpayment')) return formatNumber(rawValue);
    return rawValue;
  };

  return (
    <Draggable
      nodeRef={nodeRef}
      bounds="parent"
      position={{ x: field.x, y: field.y }}
      onStop={(_e, data) => updateField(field.id, { x: data.x, y: data.y })}
      handle=".drag-handle"
    >
      <div 
        ref={nodeRef}
        className={cn(
          "absolute cursor-default group/field transition-shadow",
          selectedFieldId === field.id ? "ring-2 ring-emerald-500 ring-offset-2 z-20 shadow-lg" : "hover:ring-1 hover:ring-zinc-300 z-10"
        )}
        style={{ width: field.width }}
        onClick={(e) => {
          e.stopPropagation();
          setSelectedFieldId(field.id);
        }}
      >
        <div style={{ transform: `rotate(${field.rotation}deg)`, width: '100%', height: '100%' }}>
          <div className={cn(
            "drag-handle absolute -top-7 left-0 bg-emerald-500 text-white text-[10px] px-2 py-1 rounded-t-md flex items-center gap-1.5 cursor-move transition-opacity",
            selectedFieldId === field.id ? "opacity-100" : "opacity-0 group-hover/field:opacity-100"
          )}>
            <Move size={10} /> 
            <span className="font-bold uppercase tracking-wider">{field.label}</span>
          </div>
          
          <div className={cn(
            "absolute -right-1 top-1/2 -translate-y-1/2 w-1 h-4 bg-emerald-500 rounded-full opacity-0 transition-opacity",
            selectedFieldId === field.id ? "opacity-100" : ""
          )} />

          <div 
            style={{ 
              fontSize: `${field.fontSize}px`, 
              lineHeight: field.lineHeight,
              color: field.color, 
              fontWeight: field.fontWeight,
              textAlign: field.textAlign,
            }}
            className="whitespace-pre-wrap break-words select-none p-1"
          >
            {getFieldValue()}
          </div>
        </div>
      </div>
    </Draggable>
  );
}

export default function App() {
  const [tags, setTags] = useState<PriceTagData[]>([]);
  const [selectedTags, setSelectedTags] = useState<Set<string>>(new Set());
  const [viewMode, setViewMode] = useState<'list' | 'design' | 'preview'>('list');
  const [templateImage, setTemplateImage] = useState<string | null>(null);
  const [fieldConfigs, setFieldConfigs] = useState<FieldConfig[]>(INITIAL_FIELDS);
  const [excelColumns, setExcelColumns] = useState<string[]>([]);
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [previewTagId, setPreviewTagId] = useState<string | null>(null);
  const [snapToGrid, setSnapToGrid] = useState(false);
  const [showGrid, setShowGrid] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [isDownloading, setIsDownloading] = useState(false);
  const [printLayout, setPrintLayout] = useState<'2-in-1' | '4-in-1' | '6-in-1'>('2-in-1');
  const [previewZoom, setPreviewZoom] = useState(0.6);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const editorRef = useRef<HTMLDivElement>(null);

  // Load saved design
  useEffect(() => {
    const savedConfigs = localStorage.getItem('emcor_field_configs');
    const savedTemplate = localStorage.getItem('emcor_template_image');
    if (savedConfigs) {
      try {
        setFieldConfigs(JSON.parse(savedConfigs));
      } catch (e) {
        console.error("Failed to load saved configs", e);
      }
    }
    if (savedTemplate) {
      setTemplateImage(savedTemplate);
    }
  }, []);

  // Save design whenever it changes
  useEffect(() => {
    localStorage.setItem('emcor_field_configs', JSON.stringify(fieldConfigs));
  }, [fieldConfigs]);

  useEffect(() => {
    if (templateImage) {
      localStorage.setItem('emcor_template_image', templateImage);
    } else {
      localStorage.removeItem('emcor_template_image');
    }
  }, [templateImage]);

  const previewTag = tags.find(t => selectedTags.has(t.id)) || tags[0];

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (viewMode !== 'design' || !selectedFieldId) return;
      
      const step = e.shiftKey ? 10 : 1;
      const field = fieldConfigs.find(f => f.id === selectedFieldId);
      if (!field) return;

      switch(e.key) {
        case 'ArrowUp':
          e.preventDefault();
          updateField(selectedFieldId, { y: Math.max(0, field.y - step) });
          break;
        case 'ArrowDown':
          e.preventDefault();
          updateField(selectedFieldId, { y: Math.min(561, field.y + step) });
          break;
        case 'ArrowLeft':
          e.preventDefault();
          updateField(selectedFieldId, { x: Math.max(0, field.x - step) });
          break;
        case 'ArrowRight':
          e.preventDefault();
          updateField(selectedFieldId, { x: Math.min(794 - field.width, field.x + step) });
          break;
        case 'Escape':
          setSelectedFieldId(null);
          break;
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [viewMode, selectedFieldId, fieldConfigs]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

      // Extract columns from the first row
      if (jsonData.length > 0) {
        const columns = Object.keys(jsonData[0]);
        setExcelColumns(columns);
      }

      const newTags: PriceTagData[] = jsonData.map((row) => ({
        id: crypto.randomUUID(),
        selected: true,
        rawData: row,
      }));

      setTags(newTags);
      setSelectedTags(new Set(newTags.map(t => t.id)));
      setViewMode('list');
    };
    reader.readAsArrayBuffer(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const addField = (label: string, mapping?: string) => {
    const newId = `field-${Date.now()}`;
    const newField: FieldConfig = {
      id: newId,
      label: label,
      x: 100,
      y: 100,
      fontSize: 18,
      lineHeight: 1.2,
      rotation: 0,
      color: '#18181b',
      fontWeight: '700',
      width: 200,
      textAlign: 'left',
      mapping: mapping
    };
    setFieldConfigs([...fieldConfigs, newField]);
    setSelectedFieldId(newId);
  };

  const removeField = (id: string) => {
    setFieldConfigs(fieldConfigs.filter(f => f.id !== id));
    setSelectedFieldId(null);
  };

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      setTemplateImage(event.target?.result as string);
      setViewMode('design');
    };
    reader.readAsDataURL(file);
  };

  const updateField = (id: string, updates: Partial<FieldConfig>) => {
    setFieldConfigs(prev => prev.map(f => {
      if (f.id === id) {
        let newX = updates.x !== undefined ? updates.x : f.x;
        let newY = updates.y !== undefined ? updates.y : f.y;

        if (snapToGrid) {
          newX = Math.round(newX / 10) * 10;
          newY = Math.round(newY / 10) * 10;
        }

        return { ...f, ...updates, x: newX, y: newY };
      }
      return f;
    }));
  };

  const toggleTagSelection = (id: string) => {
    const newSelected = new Set(selectedTags);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedTags(newSelected);
  };

  const selectAll = (selected: boolean) => {
    if (selected) {
      setSelectedTags(new Set(tags.map(t => t.id)));
    } else {
      setSelectedTags(new Set());
    }
  };

  const removeTag = (id: string) => {
    setTags(tags.filter(t => t.id !== id));
  };

  const clearAll = () => {
    setTags([]);
  };

  const addEmptyTag = () => {
    setTags(prev => [...prev, {
      id: crypto.randomUUID(),
      selected: true,
      rawData: {}
    }]);
  };


  const handlePrint = () => {
    window.print();
  };

  const handleDownloadPDF = async () => {
    if (selectedTags.size === 0) return;
    
    if (viewMode !== 'preview') {
      alert("Please switch to 'Preview' mode first to download PDF.");
      return;
    }

    setIsDownloading(true);
    
    try {
      // Scroll to top to ensure html2canvas captures correctly
      window.scrollTo(0, 0);
      
      // Determine PDF orientation and dimensions
      const isLandscape = printLayout === '2-in-1';
      const pdf = new jsPDF(isLandscape ? 'l' : 'p', 'mm', 'a4');
      const pageWidth = isLandscape ? 297 : 210;
      const pageHeight = isLandscape ? 210 : 297;
      
      const pageElements = document.querySelectorAll('.print-page-container');
      
      if (pageElements.length === 0) {
        throw new Error("No preview pages found.");
      }

      for (let i = 0; i < pageElements.length; i++) {
        if (i > 0) pdf.addPage();

        const element = pageElements[i] as HTMLElement;
        
        // Wait a tiny bit for any layout shifts
        await new Promise(resolve => setTimeout(resolve, 100));

        const canvas = await html2canvas(element, {
          scale: 2, // High quality
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          allowTaint: true,
          scrollX: 0,
          scrollY: 0,
        });
        
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        pdf.addImage(imgData, 'JPEG', 0, 0, pageWidth, pageHeight, undefined, 'FAST');
      }

      pdf.save(`EMCOR-Price-Tags-${new Date().getTime()}.pdf`);
    } catch (error) {
      console.error("PDF Generation Error:", error);
      alert("Failed to generate PDF. Please ensure you are in 'Preview' mode and try again.");
    } finally {
      setIsDownloading(false);
    }
  };

  const modelColumn = excelColumns.find(col => 
    col.toLowerCase() === 'item model' || 
    col.toLowerCase().includes('model')
  ) || excelColumns[0] || 'Item Model';

  const filteredTags = tags.filter(tag => {
    if (!searchQuery) return true;
    const searchLower = searchQuery.toLowerCase();
    // Search only in the model column
    const modelValue = String(tag.rawData[modelColumn] || '');
    return modelValue.toLowerCase().includes(searchLower);
  });

  const selectedField = fieldConfigs.find(f => f.id === selectedFieldId);

  return (
    <div className="min-h-screen bg-zinc-50 p-4 md:p-8 font-sans text-zinc-900">
      {/* Header */}
      <header className="max-w-6xl mx-auto mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4 print:hidden">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-zinc-900">EMCOR Price Tag Generator</h1>
          <p className="text-zinc-500 mt-1">Upload Excel data and customize your tag template.</p>
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={() => templateInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-white border border-zinc-200 rounded-lg hover:bg-zinc-50 transition-colors text-sm font-medium shadow-sm"
          >
            <ImageIcon size={16} />
            {templateImage ? 'Change Template' : 'Upload Template'}
          </button>
          {templateImage && (
            <button
              onClick={() => setTemplateImage(null)}
              className="flex items-center gap-2 px-4 py-2 bg-white border border-red-100 text-red-500 rounded-lg hover:bg-red-50 transition-colors text-sm font-medium shadow-sm"
              title="Remove Template"
            >
              <Trash2 size={16} />
            </button>
          )}
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors text-sm font-medium shadow-sm"
          >
            <Upload size={16} />
            Upload Excel
          </button>
          <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".xlsx, .xls, .csv" className="hidden" />
          <input type="file" ref={templateInputRef} onChange={handleTemplateUpload} accept="image/*" className="hidden" />
        </div>
      </header>

      {/* Controls */}
      <div className="max-w-6xl mx-auto mb-6 flex flex-col gap-4 print:hidden">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            {tags.length > 0 && (
              <button
                onClick={clearAll}
                className="flex items-center gap-2 px-3 py-1.5 text-red-600 hover:bg-red-50 rounded-md text-xs font-medium transition-colors"
              >
                <Trash2 size={14} />
                Clear All
              </button>
            )}
          </div>
          
          <div className="flex items-center bg-white border border-zinc-200 rounded-lg p-1 shadow-sm">
            <button
              onClick={() => setViewMode('list')}
              className={cn(
                "px-4 py-1.5 text-xs font-medium rounded-md transition-all",
                viewMode === 'list' ? "bg-zinc-900 text-white shadow-sm" : "text-zinc-500 hover:text-zinc-900"
              )}
            >
              Item List
            </button>
            <button
              onClick={() => setViewMode('design')}
              className={cn(
                "px-4 py-1.5 text-xs font-medium rounded-md transition-all",
                viewMode === 'design' ? "bg-zinc-900 text-white shadow-sm" : "text-zinc-500 hover:text-zinc-900"
              )}
            >
              Design Template
            </button>
            <button
              onClick={() => setViewMode('preview')}
              className={cn(
                "px-4 py-1.5 text-xs font-medium rounded-md transition-all",
                viewMode === 'preview' ? "bg-zinc-900 text-white shadow-sm" : "text-zinc-500 hover:text-zinc-900"
              )}
            >
              Preview
            </button>
          </div>

          {selectedTags.size > 0 && (
            <div className="flex items-center gap-2">
              <button
                onClick={handleDownloadPDF}
                disabled={isDownloading}
                className={cn(
                  "flex items-center gap-2 px-6 py-2.5 bg-zinc-900 text-white rounded-full hover:bg-zinc-800 transition-all shadow-lg hover:scale-105 active:scale-95 disabled:opacity-50 disabled:scale-100",
                  isDownloading && "animate-pulse"
                )}
              >
                {isDownloading ? (
                  <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                ) : (
                  <Download size={18} />
                )}
                {isDownloading ? 'Generating PDF...' : `Download PDF (${selectedTags.size})`}
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Main Content Area */}
      <main className="max-w-6xl mx-auto">
        {viewMode === 'list' ? (
          tags.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-20 border-2 border-dashed border-zinc-200 rounded-2xl bg-white print:hidden">
              <FileSpreadsheet size={48} className="text-zinc-300 mb-4" />
              <p className="text-zinc-500 font-medium">No tags to display. Upload an Excel file to get started.</p>
            </div>
          ) : (
            <div className="space-y-4 print:hidden">
              <div className="flex items-center justify-between gap-4">
                <div className="relative flex-1 max-w-md">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
                  <input
                    type="text"
                    placeholder="Search items..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-10 pr-4 py-2.5 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition-all shadow-sm"
                  />
                </div>
                <div className="text-xs font-bold text-zinc-400 uppercase tracking-widest">
                  Showing {filteredTags.length} of {tags.length} items
                </div>
              </div>

              <div className="bg-white border border-zinc-200 rounded-xl overflow-hidden shadow-sm">
                <table className="w-full text-left text-sm">
                  <thead className="bg-zinc-50 border-b border-zinc-200">
                    <tr>
                      <th className="px-6 py-3 font-semibold text-zinc-600 w-12 text-center">
                        <input 
                          type="checkbox" 
                          checked={filteredTags.length > 0 && filteredTags.every(t => selectedTags.has(t.id))}
                          onChange={(e) => {
                            const newSelected = new Set(selectedTags);
                            filteredTags.forEach(t => {
                              if (e.target.checked) newSelected.add(t.id);
                              else newSelected.delete(t.id);
                            });
                            setSelectedTags(newSelected);
                          }}
                          className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
                        />
                      </th>
                      <th className="px-6 py-3 font-semibold text-zinc-600">{modelColumn}</th>
                      <th className="px-6 py-3 font-semibold text-zinc-600 text-right">Actions</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-zinc-100">
                    {filteredTags.map((tag) => (
                      <tr key={tag.id} className={cn("hover:bg-zinc-50 transition-colors", !selectedTags.has(tag.id) && "opacity-60")}>
                        <td className="px-6 py-4 text-center">
                          <input 
                            type="checkbox" 
                            checked={selectedTags.has(tag.id)}
                            onChange={() => toggleTagSelection(tag.id)}
                            className="rounded border-zinc-300 text-emerald-600 focus:ring-emerald-500"
                          />
                        </td>
                        <td className="px-6 py-4 text-zinc-600 font-medium">
                          {String(tag.rawData[modelColumn] || '')}
                        </td>
                        <td className="px-6 py-4 text-right">
                          <button onClick={() => removeTag(tag.id)} className="p-1.5 text-zinc-400 hover:text-red-600 hover:bg-red-50 rounded-md transition-all">
                            <Trash2 size={16} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )
        ) : viewMode === 'design' ? (
          <div className="flex-1 flex flex-col items-center justify-center p-8 overflow-auto bg-zinc-50/50 rounded-3xl border border-zinc-200 shadow-inner relative group/canvas">
            {/* Toolbar Overlay */}
            <div className="absolute top-8 left-1/2 -translate-x-1/2 flex items-center gap-2 bg-white/90 backdrop-blur-md p-2 rounded-2xl border border-zinc-200 shadow-xl z-30 opacity-0 group-hover/canvas:opacity-100 transition-opacity duration-300">
              <div className="px-3 py-1 text-[10px] font-bold text-zinc-500 border-r border-zinc-200 mr-1">
                561 x 794 px (A5 Portrait)
              </div>
              
              {/* Insert Merge Field Dropdown */}
              <div className="relative group/dropdown">
                <button className="flex items-center gap-1.5 px-3 py-2 hover:bg-zinc-100 rounded-xl text-zinc-600 transition-colors text-xs font-bold">
                  <Plus size={16} />
                  Insert Merge Field
                </button>
                <div className="absolute top-full left-0 mt-2 w-56 bg-white rounded-2xl border border-zinc-200 shadow-2xl opacity-0 invisible group-hover/dropdown:opacity-100 group-hover/dropdown:visible transition-all duration-200 z-50 p-2">
                  {excelColumns.length > 0 ? (
                    <div className="max-h-64 overflow-y-auto">
                      <div className="px-3 py-2 text-[9px] font-black text-zinc-400 uppercase tracking-widest border-b border-zinc-50 mb-1">Excel Columns</div>
                      {excelColumns.map(col => (
                        <button
                          key={col}
                          onClick={() => addField(col, col)}
                          className="w-full text-left px-3 py-2 hover:bg-zinc-50 rounded-lg text-xs font-bold text-zinc-600 transition-colors"
                        >
                          {col}
                        </button>
                      ))}
                    </div>
                  ) : (
                    <div className="px-3 py-4 text-center">
                      <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Upload Excel first</p>
                    </div>
                  )}
                </div>
              </div>

              <div className="w-px h-4 bg-zinc-200 mx-1" />

              <button 
                onClick={() => templateInputRef.current?.click()}
                className="p-2 hover:bg-zinc-100 rounded-xl text-zinc-600 transition-colors"
                title="Change Template"
              >
                <ImageIcon size={18} />
              </button>
              {templateImage && (
                <button 
                  onClick={() => setTemplateImage(null)}
                  className="p-2 hover:bg-red-50 rounded-xl text-red-500 transition-colors"
                  title="Remove Template"
                >
                  <Trash2 size={18} />
                </button>
              )}
              <button 
                onClick={() => setShowGrid(!showGrid)}
                className={cn(
                  "p-2 rounded-xl transition-colors",
                  showGrid ? "bg-emerald-50 text-emerald-600" : "hover:bg-zinc-100 text-zinc-600"
                )}
                title="Toggle Grid"
              >
                <Settings2 size={18} />
              </button>
              <button 
                onClick={() => setSnapToGrid(!snapToGrid)}
                className={cn(
                  "p-2 rounded-xl transition-colors",
                  snapToGrid ? "bg-emerald-50 text-emerald-600" : "hover:bg-zinc-100 text-zinc-600"
                )}
                title="Snap to Grid"
              >
                <Check size={18} />
              </button>
            </div>

              <div 
                ref={editorRef}
                className="relative bg-white shadow-[0_20px_50px_rgba(0,0,0,0.15)] overflow-hidden rounded-sm transition-transform duration-300"
                style={{ width: '561px', height: '794px' }} 
                onClick={() => setSelectedFieldId(null)}
              >
                {/* Visual Grid */}
                {showGrid && (
                  <div 
                    className="absolute inset-0 pointer-events-none z-0"
                    style={{ 
                      backgroundImage: `radial-gradient(circle, #e5e7eb 1px, transparent 1px)`,
                      backgroundSize: '10px 10px'
                    }}
                  />
                )}

                {templateImage ? (
                  <img src={templateImage} className="absolute inset-0 w-full h-full object-cover pointer-events-none select-none" alt="Template" />
                ) : (
                  <div className="absolute inset-0 flex flex-col items-center justify-center text-zinc-300 border-4 border-dashed border-zinc-100 m-8 rounded-xl">
                    <ImageIcon size={64} className="mb-4 opacity-20" />
                    <p className="font-medium">No template image uploaded</p>
                    <button 
                      onClick={() => templateInputRef.current?.click()}
                      className="mt-4 text-emerald-600 font-bold text-sm hover:underline"
                    >
                      Upload Background Image
                    </button>
                  </div>
                )}

                {fieldConfigs.map((field: FieldConfig) => (
                  <DraggableField
                    key={field.id}
                    field={field}
                    selectedFieldId={selectedFieldId}
                    setSelectedFieldId={setSelectedFieldId}
                    updateField={updateField}
                    previewTag={previewTag}
                  />
                ))}
              </div>
            {/* Floating Field Settings */}
            {selectedField && (
              <div className="absolute right-12 top-1/2 -translate-y-1/2 w-72 bg-white/95 backdrop-blur-xl rounded-3xl border border-zinc-200 shadow-2xl p-6 space-y-6 animate-in fade-in slide-in-from-right-4 duration-300 z-40">
                <div className="flex items-center justify-between border-b border-zinc-100 pb-3">
                  <div className="flex items-center gap-2">
                    <div className="w-2 h-2 rounded-full bg-emerald-500" />
                    <span className="text-[10px] font-black text-zinc-900 uppercase tracking-widest">{selectedField.label} Settings</span>
                  </div>
                  <button 
                    onClick={() => setSelectedFieldId(null)}
                    className="text-[10px] font-bold text-zinc-400 hover:text-zinc-600 transition-colors"
                  >
                    Close
                  </button>
                </div>

                <div className="space-y-5">
                  <div className="space-y-1.5">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Excel Mapping</label>
                    <select 
                      value={selectedField.mapping || ''}
                      onChange={(e) => updateField(selectedField.id, { mapping: e.target.value })}
                      className="w-full px-3 py-2 bg-zinc-50 border border-zinc-100 rounded-xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                    >
                      <option value="">Default</option>
                      {excelColumns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>

                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Font Size</label>
                      <span className="text-[10px] font-bold text-zinc-900">{selectedField.fontSize}px</span>
                    </div>
                    <input 
                      type="range" min="8" max="150" value={selectedField.fontSize}
                      onChange={(e) => updateField(selectedField.id, { fontSize: parseInt(e.target.value) })}
                      className="w-full h-1 bg-zinc-100 rounded-lg appearance-none cursor-pointer accent-emerald-500"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Box Width</label>
                      <span className="text-[10px] font-bold text-zinc-900">{selectedField.width}px</span>
                    </div>
                    <input 
                      type="range" min="50" max="561" value={selectedField.width}
                      onChange={(e) => updateField(selectedField.id, { width: parseInt(e.target.value) })}
                      className="w-full h-1 bg-zinc-100 rounded-lg appearance-none cursor-pointer accent-emerald-500"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Line Height</label>
                      <span className="text-[10px] font-bold text-zinc-900">{selectedField.lineHeight}</span>
                    </div>
                    <input 
                      type="range" min="0.8" max="2.5" step="0.1" value={selectedField.lineHeight}
                      onChange={(e) => updateField(selectedField.id, { lineHeight: parseFloat(e.target.value) })}
                      className="w-full h-1 bg-zinc-100 rounded-lg appearance-none cursor-pointer accent-emerald-500"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Rotation</label>
                      <span className="text-[10px] font-bold text-zinc-900">{selectedField.rotation}°</span>
                    </div>
                    <input 
                      type="range" min="-180" max="180" value={selectedField.rotation}
                      onChange={(e) => updateField(selectedField.id, { rotation: parseInt(e.target.value) })}
                      className="w-full h-1 bg-zinc-100 rounded-lg appearance-none cursor-pointer accent-emerald-500"
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1.5">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Color</label>
                      <div className="flex items-center gap-2">
                        <input 
                          type="color" value={selectedField.color}
                          onChange={(e) => updateField(selectedField.id, { color: e.target.value })}
                          className="w-8 h-8 rounded-lg cursor-pointer border-none p-0 bg-transparent"
                        />
                        <span className="text-[9px] font-mono text-zinc-400">{selectedField.color}</span>
                      </div>
                    </div>
                    <div className="space-y-1.5">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Weight</label>
                      <select 
                        value={selectedField.fontWeight}
                        onChange={(e) => updateField(selectedField.id, { fontWeight: e.target.value })}
                        className="w-full h-8 px-2 bg-zinc-50 border border-zinc-100 rounded-lg text-[10px] font-bold"
                      >
                        <option value="400">Regular</option>
                        <option value="600">Semi</option>
                        <option value="700">Bold</option>
                        <option value="900">Black</option>
                      </select>
                    </div>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Alignment</label>
                    <div className="flex bg-zinc-50 p-1 rounded-xl border border-zinc-100">
                      {(['left', 'center', 'right'] as const).map(align => (
                        <button
                          key={align}
                          onClick={() => updateField(selectedField.id, { textAlign: align })}
                          className={cn(
                            "flex-1 py-1.5 text-[9px] font-black uppercase rounded-lg transition-all",
                            selectedField.textAlign === align ? "bg-white text-emerald-600 shadow-sm" : "text-zinc-400 hover:text-zinc-600"
                          )}
                        >
                          {align}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="pt-4 border-t border-zinc-100">
                    <button 
                      onClick={() => removeField(selectedField.id)}
                      className="w-full py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-[10px] font-black uppercase tracking-widest transition-colors flex items-center justify-center gap-2"
                    >
                      <Trash2 size={14} />
                      Remove Field
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : viewMode === 'preview' ? (
          <div className="space-y-6">
            <div className="flex flex-col md:flex-row md:items-center justify-between bg-white p-4 rounded-xl border border-zinc-200 shadow-sm print:hidden gap-4">
              <div className="flex items-center gap-4">
                <span className="text-sm font-bold text-zinc-500 uppercase tracking-widest">Print Layout:</span>
                <div className="flex bg-zinc-50 p-1 rounded-lg border border-zinc-100">
                  {(['2-in-1', '4-in-1', '6-in-1'] as const).map(layout => (
                    <button
                      key={layout}
                      onClick={() => setPrintLayout(layout)}
                      className={cn(
                        "px-4 py-1.5 text-xs font-bold rounded-md transition-all",
                        printLayout === layout ? "bg-white text-emerald-600 shadow-sm" : "text-zinc-400 hover:text-zinc-600"
                      )}
                    >
                      {layout}
                    </button>
                  ))}
                </div>
              </div>
              
              <div className="flex items-center gap-4">
                <span className="text-sm font-bold text-zinc-500 uppercase tracking-widest">Zoom:</span>
                <input 
                  type="range" 
                  min="0.2" 
                  max="1" 
                  step="0.1" 
                  value={previewZoom} 
                  onChange={(e) => setPreviewZoom(parseFloat(e.target.value))}
                  className="w-32 accent-emerald-600"
                />
                <span className="text-xs font-bold text-zinc-400 w-8">{Math.round(previewZoom * 100)}%</span>
              </div>

              <div className="text-[10px] text-zinc-400 font-medium italic">
                {printLayout === '2-in-1' ? 'A4 Landscape (2 tags)' : printLayout === '4-in-1' ? 'A4 Portrait (4 tags)' : 'A4 Portrait (6 tags)'}
              </div>
            </div>

            <div className="flex flex-col items-center gap-8 overflow-auto pb-20">
              {tags.filter(t => selectedTags.has(t.id)).length === 0 ? (
                <div className="w-full flex flex-col items-center justify-center py-20 border-2 border-dashed border-zinc-200 rounded-2xl bg-white print:hidden">
                  <p className="text-zinc-500 font-medium">No items selected for preview. Please go to Item List and select items.</p>
                  <button onClick={() => setViewMode('list')} className="mt-4 text-emerald-600 font-semibold hover:underline">Go to Item List</button>
                </div>
              ) : (
                (() => {
                  const selectedItems = tags.filter(t => selectedTags.has(t.id));
                  const itemsPerPage = printLayout === '2-in-1' ? 2 : printLayout === '4-in-1' ? 4 : 6;
                  const pages = Math.ceil(selectedItems.length / itemsPerPage);
                  
                  return Array.from({ length: pages }).map((_, pageIndex) => (
                    <div 
                      key={pageIndex} 
                      className={cn(
                        "print-page-container print:break-after-page print:mx-auto print:bg-white bg-white shadow-2xl rounded-sm overflow-hidden print:mb-0 print:shadow-none print:rounded-none flex-shrink-0",
                        printLayout === '2-in-1' ? "print:h-[210mm] print:w-[297mm] w-[297mm] h-[210mm]" : "print:h-[297mm] print:w-[210mm] w-[210mm] h-[297mm]"
                      )}
                      style={{ 
                        transform: `scale(${previewZoom})`, 
                        transformOrigin: 'top center',
                        marginBottom: `calc(${printLayout === '2-in-1' ? '210mm' : '297mm'} * ${previewZoom - 1} + 2rem)`
                      }}
                    >
                      <div className={cn(
                        "grid h-full w-full",
                        printLayout === '2-in-1' ? "grid-cols-2" : "grid-cols-2 grid-rows-2",
                        printLayout === '6-in-1' && "grid-rows-3"
                      )}>
                        {Array.from({ length: itemsPerPage }).map((_, offset) => {
                          const tagIndex = pageIndex * itemsPerPage + offset;
                          const tag = selectedItems[tagIndex];
                          
                          if (!tag) return <div key={offset} className="border border-zinc-50 print:border-none" />;
                          
                          const scale = printLayout === '2-in-1' ? 1 : printLayout === '4-in-1' ? 0.707 : 0.47;
                          
                          return (
                            <div 
                              key={tag.id} 
                              className="relative border border-zinc-50 print:border-none overflow-hidden flex items-center justify-center bg-white"
                            >
                              <div style={{ transform: `scale(${scale})`, transformOrigin: 'center' }}>
                                {templateImage ? (
                                  <CustomPriceTag tag={tag} templateImage={templateImage} fieldConfigs={fieldConfigs} />
                                ) : (
                                  <DefaultPriceTag tag={tag} />
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ));
                })()
              )}
            </div>
          </div>
        ) : null}
      </main>

      <style dangerouslySetInnerHTML={{ __html: `
        @media print {
          body { background: white !important; margin: 0 !important; padding: 0 !important; }
          @page { size: A4; margin: 0; }
          .print-hidden { display: none !important; }
        }
      `}} />
    </div>
  );
}

function CustomPriceTag({ tag, templateImage, fieldConfigs }: { tag: PriceTagData, templateImage: string, fieldConfigs: FieldConfig[] }) {
  return (
    <div className="relative overflow-hidden bg-white shadow-sm" style={{ width: '148.5mm', height: '210mm' }}>
      <img src={templateImage} className="absolute inset-0 w-full h-full object-cover" alt="Template" />
      {fieldConfigs.map(field => {
        let value = '';
        let rawValue = '';
        
        if (field.mapping && tag.rawData) {
          rawValue = String(tag.rawData[field.mapping] || '');
        }

        if (field.label.toLowerCase().includes('srp')) value = formatCurrency(rawValue);
        else if (field.label.toLowerCase().includes('downpayment')) value = formatNumber(rawValue);
        else value = rawValue;
        
        return (
          <div 
            key={field.id}
            className="absolute leading-none whitespace-pre-wrap break-words"
            style={{ 
              left: `${field.x}px`, 
              top: `${field.y}px`, 
              width: `${field.width}px`,
              fontSize: `${field.fontSize}px`, 
              color: field.color, 
              fontWeight: field.fontWeight,
              textAlign: field.textAlign,
              lineHeight: field.lineHeight,
              transform: `rotate(${field.rotation}deg)`,
            }}
          >
            {value}
          </div>
        );
      })}
    </div>
  );
}

function DefaultPriceTag({ tag }: { tag: PriceTagData }) {
  return (
    <div className="w-full h-full flex flex-col bg-white overflow-hidden relative p-4 print:p-0" style={{ minHeight: '210mm' }}>
      <div className="flex-1 flex flex-col m-2 relative z-0 items-center justify-center border-2 border-dashed rounded-xl" style={{ borderColor: '#f4f4f5' }}>
        <p className="font-bold text-xs uppercase tracking-widest" style={{ color: '#a1a1aa' }}>Custom Template Required</p>
        <p className="text-[10px] mt-1" style={{ color: '#d4d4d8' }}>Design your template to see data here</p>
      </div>
    </div>
  );
}

function formatCurrency(val: string | number) {
  const num = typeof val === 'string' ? parseFloat(val.replace(/,/g, '')) : val;
  if (isNaN(num as number)) return String(val);
  return new Intl.NumberFormat('en-PH', { style: 'decimal', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(num as number);
}

function formatNumber(val: string | number) {
  const num = typeof val === 'string' ? parseFloat(val.replace(/,/g, '')) : val;
  if (isNaN(num as number)) return String(val);
  return new Intl.NumberFormat('en-PH', { style: 'decimal', minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(num as number);
}
