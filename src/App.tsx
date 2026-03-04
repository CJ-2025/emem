/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Printer, FileSpreadsheet, Trash2, Plus, Download, Image as ImageIcon, Settings2, Move, Check, Type, Search, FileText, AlignLeft, AlignCenter, AlignRight, X, Minus, Square, Maximize2 } from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import Draggable from 'react-draggable';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { Document, Packer, Paragraph, ImageRun, SectionType, PageOrientation } from 'docx';
import { saveAs } from 'file-saver';

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
  vOffset?: number;
  formatAsNumber?: boolean;
}

interface PriceTagData {
  id: string;
  selected: boolean;
  rawData: Record<string, any>;
}

const INITIAL_FIELDS: FieldConfig[] = [];

const PAPER_DIMENSIONS = {
  'A4': { width: 794, height: 1122, label: 'A4 (210 x 297mm)' },
  'Letter': { width: 816, height: 1056, label: 'Letter (8.5 x 11in)' },
  'Legal': { width: 816, height: 1344, label: 'Legal (8.5 x 14in)' }
};

interface DraggableFieldProps {
  field: FieldConfig;
  selectedFieldId: string | null;
  setSelectedFieldId: (id: string | null) => void;
  updateField: (id: string, updates: Partial<FieldConfig>) => void;
  previewTag: PriceTagData | undefined;
  updateTagData: (tagId: string, key: string, value: string) => void;
}

const DraggableField: React.FC<DraggableFieldProps> = ({ 
  field, 
  selectedFieldId, 
  setSelectedFieldId, 
  updateField, 
  previewTag,
  updateTagData
}) => {
  const nodeRef = useRef(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const [isEditing, setIsEditing] = useState(false);
  
  const getFieldValue = () => {
    let rawValue = '';
    if (field.mapping && previewTag?.rawData) {
      rawValue = String(previewTag.rawData[field.mapping] || '');
    } else {
      return field.label;
    }

    const labelLower = field.label.toLowerCase();
    const shouldFormat = field.formatAsNumber || 
                        labelLower.includes('srp') || 
                        labelLower.includes('down') || 
                        labelLower.includes('price') || 
                        labelLower.includes('amount') ||
                        labelLower.includes('cash') ||
                        labelLower.includes('total') ||
                        labelLower.includes('monthly');

    if (shouldFormat) {
      if (labelLower.includes('srp')) return formatCurrency(rawValue);
      return formatNumber(rawValue);
    }
    
    return rawValue;
  };

  const currentRawValue = field.mapping && previewTag?.rawData 
    ? String(previewTag.rawData[field.mapping] || '') 
    : field.label;

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [isEditing]);

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
        <div style={{ transform: `rotate(${field.rotation}deg)`, transformOrigin: '0 0', width: '100%' }}>
          <div className={cn(
            "drag-handle absolute -top-7 left-0 bg-emerald-500 text-white text-[10px] px-2 py-1 rounded-t-md flex items-center gap-1.5 cursor-move transition-opacity",
            selectedFieldId === field.id ? "opacity-100" : "opacity-0 group-hover/field:opacity-100"
          )}>
            <Move size={10} /> 
            <span className="font-bold uppercase tracking-wider">{field.label}</span>
            {selectedFieldId === field.id && <span className="ml-2 opacity-70 font-normal normal-case italic">Double-click text to edit data</span>}
          </div>
          
          <div className={cn(
            "absolute -right-1 top-1/2 -translate-y-1/2 w-1 h-4 bg-emerald-500 rounded-full opacity-0 transition-opacity",
            selectedFieldId === field.id ? "opacity-100" : ""
          )} />

          <div 
            style={{ 
              fontSize: `${field.fontSize}px`, 
              lineHeight: field.lineHeight || 1,
              color: field.color, 
              fontWeight: field.fontWeight,
              textAlign: field.textAlign,
              transform: `translateY(${field.vOffset || 0}px)`,
              transformOrigin: '0 0',
              display: 'block'
            }}
            className={cn(
              "whitespace-pre-wrap break-words p-0 transition-all",
              isEditing ? "select-text" : "select-none cursor-text"
            )}
            onDoubleClick={(e) => {
              e.stopPropagation();
              setIsEditing(true);
            }}
          >
            {isEditing ? (
              <textarea
                ref={inputRef}
                value={currentRawValue}
                onChange={(e) => {
                  if (field.mapping && previewTag) {
                    updateTagData(previewTag.id, field.mapping, e.target.value);
                  } else {
                    updateField(field.id, { label: e.target.value });
                  }
                }}
                onBlur={() => setIsEditing(false)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    setIsEditing(false);
                  }
                  if (e.key === 'Escape') {
                    setIsEditing(false);
                  }
                }}
                className="w-full bg-white/95 border-2 border-emerald-500 rounded-lg p-1 outline-none text-zinc-900 shadow-2xl z-50 relative"
                style={{ 
                  fontSize: 'inherit', 
                  fontWeight: 'inherit', 
                  textAlign: 'inherit',
                  lineHeight: 'inherit',
                  minHeight: '1.2em',
                  color: '#000'
                }}
              />
            ) : (
              getFieldValue()
            )}
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
  const [paperSize, setPaperSize] = useState<'A4' | 'Letter' | 'Legal'>('A4');
  const [previewZoom, setPreviewZoom] = useState(0.6);
  const [searchColumn, setSearchColumn] = useState<string>('');
  const [fillMode, setFillMode] = useState(true);
  const [isDownloadMenuOpen, setIsDownloadMenuOpen] = useState(false);
  const [isDownloadingWord, setIsDownloadingWord] = useState(false);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const editorRef = useRef<HTMLDivElement>(null);
  const downloadMenuRef = useRef<HTMLDivElement>(null);

  // Close download menu on click outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (downloadMenuRef.current && !downloadMenuRef.current.contains(event.target as Node)) {
        setIsDownloadMenuOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

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

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const allNewTags: PriceTagData[] = [];
    const allColumns = new Set<string>(excelColumns);

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

      if (jsonData.length > 0) {
        Object.keys(jsonData[0]).forEach(col => allColumns.add(col));
      }

      const fileTags: PriceTagData[] = jsonData.map((row) => ({
        id: crypto.randomUUID(),
        selected: true,
        rawData: row,
      }));
      allNewTags.push(...fileTags);
    }

    const updatedColumns = Array.from(allColumns);
    setExcelColumns(updatedColumns);
    
    // Set default search column if not already set
    if (!searchColumn && updatedColumns.length > 0) {
      const defaultCol = updatedColumns.find(col => 
        col.toLowerCase() === 'item model' || 
        col.toLowerCase().includes('model')
      ) || updatedColumns[0] || '';
      setSearchColumn(defaultCol);
    }

    const finalTags = [...tags, ...allNewTags];
    setTags(finalTags);
    
    // Update selected tags to include new ones
    const newSelectedTags = new Set(selectedTags);
    allNewTags.forEach(t => newSelectedTags.add(t.id));
    setSelectedTags(newSelectedTags);
    
    setViewMode('list');
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
      lineHeight: 1,
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

  const updateTagData = (tagId: string, key: string, value: string) => {
    setTags(prev => prev.map(tag => 
      tag.id === tagId 
        ? { ...tag, rawData: { ...tag.rawData, [key]: value } }
        : tag
    ));
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
    const selectedItems = tags.filter(t => selectedTags.has(t.id));
    if (selectedItems.length === 0) {
      alert("No items selected for PDF generation. Please select items from the list.");
      return;
    }
    
    // Automatically switch to preview mode if not already there
    const originalViewMode = viewMode;
    if (viewMode !== 'preview') {
      setViewMode('preview');
      // Wait longer if we had to switch modes to ensure everything renders
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    setIsDownloading(true);
    
    // Store original zoom to restore later
    const originalZoom = previewZoom;
    // Reset zoom to 1 for accurate capture
    setPreviewZoom(1);
    
    // Wait for React to re-render and for any layout shifts to settle
    await new Promise(resolve => setTimeout(resolve, 800));
    
    try {
      // Scroll to top to ensure html2canvas captures correctly
      window.scrollTo(0, 0);
      
      // Determine PDF orientation and dimensions
      const isLandscape = printLayout === '2-in-1';
      const pdfSize = paperSize.toLowerCase() as any;
      const pdf = new jsPDF(isLandscape ? 'l' : 'p', 'mm', pdfSize);
      
      const mmDimensions = {
        'A4': { w: 210, h: 297 },
        'Letter': { w: 215.9, h: 279.4 },
        'Legal': { w: 215.9, h: 355.6 }
      };
      
      const currentMM = mmDimensions[paperSize];
      const pageWidth = isLandscape ? currentMM.h : currentMM.w;
      const pageHeight = isLandscape ? currentMM.w : currentMM.h;
      
      const pageElements = document.querySelectorAll('.print-page-container');
      
      if (pageElements.length === 0) {
        throw new Error("No preview pages found in the document. Please ensure you are in 'Preview' mode.");
      }

      const dimensions = PAPER_DIMENSIONS[paperSize];
      const canvasWidth = isLandscape ? dimensions.height : dimensions.width;
      const canvasHeight = isLandscape ? dimensions.width : dimensions.height;

      for (let i = 0; i < pageElements.length; i++) {
        if (i > 0) pdf.addPage(pdfSize, isLandscape ? 'l' : 'p');

        const element = pageElements[i] as HTMLElement;
        
        const canvas = await html2canvas(element, {
          scale: 2, // Slightly lower scale for better performance/reliability
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          allowTaint: true,
          scrollX: 0,
          scrollY: 0,
          windowWidth: canvasWidth,
          windowHeight: canvasHeight
        });
        
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        pdf.addImage(imgData, 'JPEG', 0, 0, pageWidth, pageHeight, undefined, 'FAST');
      }

      pdf.save(`EMCOR-Price-Tags-${new Date().getTime()}.pdf`);
    } catch (error) {
      console.error("PDF Generation Error:", error);
      alert(`Failed to generate PDF: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setPreviewZoom(originalZoom);
      setIsDownloading(false);
      // We stay in preview mode as it's where the user can see what was generated
    }
  };

  const handleDownloadWord = async () => {
    const selectedItems = tags.filter(t => selectedTags.has(t.id));
    if (selectedItems.length === 0) {
      alert("No items selected for Word generation. Please select items from the list.");
      return;
    }
    
    const originalViewMode = viewMode;
    if (viewMode !== 'preview') {
      setViewMode('preview');
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    setIsDownloadingWord(true);
    const originalZoom = previewZoom;
    setPreviewZoom(1);
    await new Promise(resolve => setTimeout(resolve, 800));
    
    try {
      window.scrollTo(0, 0);
      const pageElements = document.querySelectorAll('.print-page-container');
      
      if (pageElements.length === 0) {
        throw new Error("No preview pages found.");
      }

      const isLandscape = printLayout === '2-in-1';
      const dimensions = PAPER_DIMENSIONS[paperSize];
      const canvasWidth = isLandscape ? dimensions.height : dimensions.width;
      const canvasHeight = isLandscape ? dimensions.width : dimensions.height;

      const sections = [];

      for (let i = 0; i < pageElements.length; i++) {
        const element = pageElements[i] as HTMLElement;
        const canvas = await html2canvas(element, {
          scale: 2,
          useCORS: true,
          logging: false,
          backgroundColor: '#ffffff',
          allowTaint: true,
          scrollX: 0,
          scrollY: 0,
          windowWidth: canvasWidth,
          windowHeight: canvasHeight
        });
        
        const imgData = canvas.toDataURL('image/png');
        const base64Data = imgData.split(',')[1];
        const binaryData = atob(base64Data);
        const arrayBuffer = new ArrayBuffer(binaryData.length);
        const uint8Array = new Uint8Array(arrayBuffer);
        for (let j = 0; j < binaryData.length; j++) {
          uint8Array[j] = binaryData.charCodeAt(j);
        }

        // Word page size in twips (1/1440 of an inch)
        // A4: 11906 x 16838
        const mmToTwips = (mm: number) => Math.round((mm / 25.4) * 1440);
        
        const mmDimensions = {
          'A4': { w: 210, h: 297 },
          'Letter': { w: 215.9, h: 279.4 },
          'Legal': { w: 215.9, h: 355.6 }
        };
        const currentMM = mmDimensions[paperSize];
        const pageWidthTwips = mmToTwips(isLandscape ? currentMM.h : currentMM.w);
        const pageHeightTwips = mmToTwips(isLandscape ? currentMM.w : currentMM.h);

        sections.push({
          properties: {
            page: {
              size: {
                width: pageWidthTwips,
                height: pageHeightTwips,
              },
              orientation: isLandscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
              margin: {
                top: 0,
                right: 0,
                bottom: 0,
                left: 0,
              },
            },
            type: SectionType.NEXT_PAGE,
          },
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: uint8Array,
                  transformation: {
                    width: isLandscape ? currentMM.h * 3.78 : currentMM.w * 3.78, // mm to pixels approx
                    height: isLandscape ? currentMM.w * 3.78 : currentMM.h * 3.78,
                  },
                } as any),
              ],
            }),
          ],
        });
      }

      const doc = new Document({
        sections: sections,
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `EMCOR-Price-Tags-${new Date().getTime()}.docx`);
    } catch (error) {
      console.error("Word Generation Error:", error);
      alert(`Failed to generate Word document: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setPreviewZoom(originalZoom);
      setIsDownloadingWord(false);
      setIsDownloadMenuOpen(false);
    }
  };

  const filteredTags = tags.filter(tag => {
    if (!searchQuery) return true;
    const searchLower = searchQuery.toLowerCase();
    // Search only in the selected column
    const modelValue = String(tag.rawData[searchColumn] || '');
    return modelValue.toLowerCase().includes(searchLower);
  });

  const selectedField = fieldConfigs.find(f => f.id === selectedFieldId);

  return (
    <div className="min-h-screen bg-zinc-50 p-4 md:p-8 font-sans text-zinc-900 print:bg-white print:p-0">
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
          <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".xlsx, .xls, .csv" multiple className="hidden" />
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
            <div className="flex items-center gap-2 relative" ref={downloadMenuRef}>
              <div className="relative">
                <button
                  onClick={() => setIsDownloadMenuOpen(!isDownloadMenuOpen)}
                  disabled={isDownloading || isDownloadingWord}
                  className={cn(
                    "flex items-center gap-2 px-6 py-2.5 bg-zinc-900 text-white rounded-full hover:bg-zinc-800 transition-all shadow-lg hover:scale-105 active:scale-95 disabled:opacity-50 disabled:scale-100",
                    (isDownloading || isDownloadingWord) && "animate-pulse"
                  )}
                >
                  {isDownloading || isDownloadingWord ? (
                    <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                  ) : (
                    <Download size={18} />
                  )}
                  {isDownloading ? 'Generating PDF...' : isDownloadingWord ? 'Generating Word...' : `Download (${selectedTags.size})`}
                  <Plus size={14} className={cn("transition-transform", isDownloadMenuOpen ? "rotate-45" : "")} />
                </button>

                {isDownloadMenuOpen && (
                  <div className="absolute top-full right-0 mt-2 w-48 bg-white rounded-2xl border border-zinc-200 shadow-2xl z-[100] p-2 animate-in fade-in slide-in-from-top-2 duration-200">
                    <button
                      onClick={() => {
                        handleDownloadPDF();
                        setIsDownloadMenuOpen(false);
                      }}
                      className="w-full flex items-center gap-3 px-4 py-3 hover:bg-zinc-50 rounded-xl text-sm font-bold text-zinc-700 transition-colors"
                    >
                      <FileText size={18} className="text-red-500" />
                      Download as PDF
                    </button>
                    <button
                      onClick={() => {
                        handleDownloadWord();
                        setIsDownloadMenuOpen(false);
                      }}
                      className="w-full flex items-center gap-3 px-4 py-3 hover:bg-zinc-50 rounded-xl text-sm font-bold text-zinc-700 transition-colors"
                    >
                      <FileText size={18} className="text-blue-500" />
                      Download as Word
                    </button>
                  </div>
                )}
              </div>
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
              <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
                <div className="flex flex-1 items-center gap-3 max-w-2xl">
                  <div className="relative flex-1">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
                    <input
                      type="text"
                      placeholder={`Search by ${searchColumn}...`}
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-10 pr-4 py-2.5 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none transition-all shadow-sm"
                    />
                  </div>
                  <div className="flex items-center gap-2 min-w-[200px]">
                    <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest whitespace-nowrap">Search In:</span>
                    <select 
                      value={searchColumn}
                      onChange={(e) => setSearchColumn(e.target.value)}
                      className="flex-1 bg-white border border-zinc-200 rounded-xl px-3 py-2.5 text-sm font-medium text-zinc-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all cursor-pointer shadow-sm"
                    >
                      {excelColumns.map(col => (
                        <option key={col} value={col}>{col}</option>
                      ))}
                    </select>
                  </div>
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
                      <th className="px-6 py-3 font-semibold text-zinc-600">{searchColumn}</th>
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
                          {String(tag.rawData[searchColumn] || '')}
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
          <div className="flex flex-col lg:flex-row gap-8 items-start justify-center">
            <div className="flex-1 w-full flex flex-col items-center justify-center p-8 overflow-auto bg-zinc-50/50 rounded-3xl border border-zinc-200 shadow-inner relative group/canvas min-h-[850px]">
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
                    updateTagData={updateTagData}
                  />
                ))}
              </div>
            </div>

            {/* Sidebar Field Settings */}
            {selectedField && (
              <div className="w-full lg:w-80 bg-white rounded-3xl border border-zinc-200 shadow-xl p-6 space-y-6 animate-in fade-in slide-in-from-right-4 duration-300 z-40 sticky top-8 max-h-[calc(100vh-120px)] overflow-y-auto custom-scrollbar">
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
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Field Label / Static Text</label>
                    <input 
                      type="text"
                      value={selectedField.label}
                      onChange={(e) => updateField(selectedField.id, { label: e.target.value })}
                      className="w-full px-3 py-2 bg-zinc-50 border border-zinc-100 rounded-xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                      placeholder="Enter field text..."
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Excel Mapping</label>
                    <div className="flex gap-2">
                      <select 
                        value={selectedField.mapping || ''}
                        onChange={(e) => updateField(selectedField.id, { mapping: e.target.value })}
                        className="flex-1 px-3 py-2 bg-zinc-50 border border-zinc-100 rounded-xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                      >
                        <option value="">Default</option>
                        {excelColumns.map(col => (
                          <option key={col} value={col}>{col}</option>
                        ))}
                      </select>
                      <button
                        onClick={() => updateField(selectedField.id, { formatAsNumber: !selectedField.formatAsNumber })}
                        className={cn(
                          "px-3 py-2 rounded-xl border transition-all flex items-center gap-2 text-[10px] font-bold",
                          selectedField.formatAsNumber 
                            ? "bg-emerald-500 border-emerald-600 text-white shadow-sm" 
                            : "bg-zinc-50 border-zinc-100 text-zinc-400 hover:text-zinc-600"
                        )}
                        title="Format with commas"
                      >
                        <span className="text-xs">,</span>
                        {selectedField.formatAsNumber ? "ON" : "OFF"}
                      </button>
                    </div>
                  </div>

                  {selectedField.mapping && previewTag && (
                    <div className="space-y-2 p-4 bg-emerald-50 rounded-2xl border border-emerald-100 shadow-sm">
                      <div className="flex items-center justify-between">
                        <label className="text-[9px] font-black text-emerald-600 uppercase tracking-widest">Edit Data for this Tag</label>
                        <span className="text-[8px] px-1.5 py-0.5 bg-emerald-100 text-emerald-700 rounded font-bold uppercase">Excel Value</span>
                      </div>
                      <textarea 
                        value={previewTag.rawData[selectedField.mapping] || ''}
                        onChange={(e) => updateTagData(previewTag.id, selectedField.mapping!, e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-emerald-200 rounded-xl text-xs font-bold text-emerald-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 min-h-[60px] resize-none"
                        placeholder="Edit excel value..."
                      />
                      <p className="text-[9px] text-emerald-600/70 font-medium leading-tight">
                        Changes here update the actual data for this specific item in your list.
                      </p>
                    </div>
                  )}

                  <div className="space-y-1.5">
                    <div className="flex justify-between items-center">
                      <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest">Font Size</label>
                      <span className="text-[10px] font-bold text-zinc-900">{selectedField.fontSize}px</span>
                    </div>
                    <input 
                      type="range" min="8" max="300" value={selectedField.fontSize}
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
                      {[
                        { id: 'left', icon: AlignLeft },
                        { id: 'center', icon: AlignCenter },
                        { id: 'right', icon: AlignRight },
                      ].map(align => (
                        <button
                          key={align.id}
                          onClick={() => updateField(selectedField.id, { textAlign: align.id as any })}
                          className={cn(
                            "flex-1 py-2 flex justify-center items-center rounded-lg transition-all",
                            selectedField.textAlign === align.id ? "bg-white text-emerald-600 shadow-sm" : "text-zinc-400 hover:text-zinc-600"
                          )}
                          title={`Align ${align.id}`}
                        >
                          <align.icon size={14} />
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
              <div className="flex items-center gap-3">
                <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">Paper Size</span>
                <select 
                  value={paperSize}
                  onChange={(e) => setPaperSize(e.target.value as any)}
                  className="bg-zinc-50 border border-zinc-200 rounded-lg px-3 py-1.5 text-xs font-bold text-zinc-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all cursor-pointer"
                >
                  <option value="A4">A4 (210 x 297mm)</option>
                  <option value="Letter">Letter (8.5 x 11in)</option>
                  <option value="Legal">Legal (8.5 x 14in)</option>
                </select>
              </div>

              <div className="flex items-center gap-3">
                <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">Print Layout</span>
                <select 
                  value={printLayout}
                  onChange={(e) => setPrintLayout(e.target.value as any)}
                  className="bg-zinc-50 border border-zinc-200 rounded-lg px-3 py-1.5 text-xs font-bold text-zinc-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all cursor-pointer"
                >
                  <option value="2-in-1">2-in-1 (Landscape)</option>
                  <option value="4-in-1">4-in-1 (Portrait)</option>
                  <option value="6-in-1">6-in-1 (Portrait)</option>
                </select>
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

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setFillMode(!fillMode)}
                  className={cn(
                    "flex items-center gap-2 px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all",
                    fillMode 
                      ? "bg-emerald-600 text-white shadow-md" 
                      : "bg-zinc-100 text-zinc-400 hover:bg-zinc-200"
                  )}
                >
                  {fillMode ? <Maximize2 size={12} /> : <Square size={12} />}
                  {fillMode ? 'Fill Mode: ON' : 'Fill Mode: OFF'}
                </button>
              </div>

              <div className="text-[10px] text-zinc-400 font-medium italic">
                {paperSize} {printLayout === '2-in-1' ? 'Landscape (2 tags)' : printLayout === '4-in-1' ? 'Portrait (4 tags)' : 'Portrait (6 tags)'}
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
                  
                  const dimensions = PAPER_DIMENSIONS[paperSize];
                  const isLandscape = printLayout === '2-in-1';
                  const pageWidth = isLandscape ? dimensions.height : dimensions.width;
                  const pageHeight = isLandscape ? dimensions.width : dimensions.height;

                  return Array.from({ length: pages }).map((_, pageIndex) => (
                    <div 
                      key={pageIndex} 
                      className={cn(
                        "print-page-container print:break-after-page print:mx-auto print:bg-white bg-white shadow-2xl rounded-sm overflow-hidden print:mb-0 print:shadow-none print:rounded-none flex-shrink-0",
                      )}
                      style={{ 
                        width: `${pageWidth}px`,
                        height: `${pageHeight}px`,
                        transform: `scale(${previewZoom})`, 
                        transformOrigin: 'top center',
                        marginBottom: `calc(${pageHeight}px * ${previewZoom - 1} + 2rem)`
                      }}
                    >
                      <div className={cn(
                        "grid h-full w-full gap-0",
                        printLayout === '2-in-1' ? "grid-cols-2" : "grid-cols-2 grid-rows-2",
                        printLayout === '6-in-1' && "grid-rows-3"
                      )}>
                        {Array.from({ length: itemsPerPage }).map((_, offset) => {
                          const tagIndex = pageIndex * itemsPerPage + offset;
                          const tag = selectedItems[tagIndex];
                          
                          if (!tag) return <div key={offset} className="print:border-none" />;
                          
                          let scaleX = 1;
                          let scaleY = 1;
                          const cols = 2;
                          const rows = printLayout === '2-in-1' ? 1 : (printLayout === '4-in-1' ? 2 : 3);
                          const cellWidth = pageWidth / cols;
                          const cellHeight = pageHeight / rows;

                          if (fillMode) {
                            scaleX = cellWidth / 561;
                            scaleY = cellHeight / 794;
                          } else {
                            const scale = Math.min(cellWidth / 561, cellHeight / 794);
                            scaleX = scale;
                            scaleY = scale;
                            if (printLayout === '2-in-1') {
                              scaleX = Math.min(1, scaleX);
                              scaleY = Math.min(1, scaleY);
                            }
                          }
                          
                          return (
                            <div 
                              key={tag.id} 
                              className="relative print:border-none overflow-hidden flex items-center justify-center bg-white"
                            >
                              <div style={{ transform: `scale(${scaleX}, ${scaleY})`, transformOrigin: 'center' }}>
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
          @page { size: ${paperSize} ${printLayout === '2-in-1' ? 'landscape' : 'portrait'}; margin: 0; }
          .print-hidden { display: none !important; }
        }
      `}} />
    </div>
  );
}

function CustomPriceTag({ tag, templateImage, fieldConfigs }: { tag: PriceTagData, templateImage: string, fieldConfigs: FieldConfig[] }) {
  return (
    <div className="relative overflow-hidden bg-white shadow-sm" style={{ width: '561px', height: '794px' }}>
      <img src={templateImage} className="absolute inset-0 w-full h-full object-cover" alt="Template" />
      {fieldConfigs.map(field => {
        let value = '';
        let rawValue = '';
        
        if (field.mapping && tag.rawData) {
          rawValue = String(tag.rawData[field.mapping] || '');
        }

        const labelLower = field.label.toLowerCase();
        const shouldFormat = field.formatAsNumber || 
                            labelLower.includes('srp') || 
                            labelLower.includes('down') || 
                            labelLower.includes('price') || 
                            labelLower.includes('amount') ||
                            labelLower.includes('cash') ||
                            labelLower.includes('total') ||
                            labelLower.includes('monthly');

        if (shouldFormat) {
          if (labelLower.includes('srp')) value = formatCurrency(rawValue);
          else value = formatNumber(rawValue);
        } else {
          value = rawValue;
        }
        
        return (
          <div 
            key={field.id}
            className="absolute"
            style={{ 
              left: `${field.x}px`, 
              top: `${field.y - 28}px`, 
              width: `${field.width}px`,
            }}
          >
            <div style={{ transform: `rotate(${field.rotation}deg)`, transformOrigin: '0 0' }}>
              <div 
                style={{ 
                  fontSize: `${field.fontSize}px`, 
                  lineHeight: field.lineHeight || 1,
                  color: field.color, 
                  fontWeight: field.fontWeight,
                  textAlign: field.textAlign,
                  transform: `translateY(${field.vOffset || 0}px)`,
                  transformOrigin: '0 0',
                  display: 'block'
                }}
                className="whitespace-pre-wrap break-words p-0"
              >
                {value}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function DefaultPriceTag({ tag }: { tag: PriceTagData }) {
  return (
    <div className="flex flex-col bg-white overflow-hidden relative p-4 print:p-0" style={{ width: '561px', height: '794px' }}>
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
