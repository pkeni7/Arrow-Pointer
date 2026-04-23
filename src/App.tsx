/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Upload, ArrowRight, Image as ImageIcon, Loader2, CheckCircle2, AlertCircle, Maximize2, List, Download, FileText, ChevronLeft, ChevronRight, Package, Tag, User, FileSpreadsheet } from 'lucide-react';
import { detectArrowsAndTargets, DetectionResult, BoundingBox } from './services/geminiService';
import * as pdfjs from 'pdfjs-dist';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import ExcelJS from 'exceljs';

// Set up PDF.js worker
pdfjs.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjs.version}/build/pdf.worker.min.mjs`;

interface CroppedResult extends DetectionResult {
  croppedUrl: string;
  pageNumber?: number;
}

export default function App() {
  const [image, setImage] = useState<string | null>(null);
  const [mimeType, setMimeType] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [results, setResults] = useState<CroppedResult[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isExcelDownloading, setIsExcelDownloading] = useState(false);
  const [isZipping, setIsZipping] = useState(false);
  const [allResults, setAllResults] = useState<CroppedResult[]>([]);
  const [isProcessingAll, setIsProcessingAll] = useState(false);
  const [processingProgress, setProcessingProgress] = useState({ current: 0, total: 0 });
  const [whiteBackground, setWhiteBackground] = useState(true);
  
  // PDF related state
  const [pdfDoc, setPdfDoc] = useState<pdfjs.PDFDocumentProxy | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [isPdfLoading, setIsPdfLoading] = useState(false);
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const imageRef = useRef<HTMLImageElement>(null);

  const downloadAsExcel = async () => {
    const targetResults = allResults.length > 0 ? allResults : results;
    if (targetResults.length === 0) return;
    setIsExcelDownloading(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Detected Items');

      // Set columns
      worksheet.columns = [
        { header: 'Label', key: 'label', width: 10 },
        { header: 'Image', key: 'image', width: 25 },
        { header: 'SAP Code', key: 'sap_code', width: 20 },
        { header: 'Name', key: 'name', width: 30 },
        { header: 'Description', key: 'description', width: 50 },
        { header: 'Page', key: 'page', width: 8 },
        { header: 'X (min)', key: 'xmin', width: 10 },
        { header: 'Y (min)', key: 'ymin', width: 10 },
      ];

      // Format headers
      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' }
      };

      for (let i = 0; i < targetResults.length; i++) {
        const res = targetResults[i];
        const rowIndex = i + 2;
        const row = worksheet.addRow({
          label: res.label,
          sap_code: res.sap_code,
          name: res.name,
          description: res.description,
          page: res.pageNumber || 1,
          xmin: Math.round(res.target_box.xmin),
          ymin: Math.round(res.target_box.ymin),
        });

        // Set row height to accommodate image
        row.height = 120;

        // Add image
        try {
          const imageBase64 = res.croppedUrl.split(',')[1];
          const imageId = workbook.addImage({
            base64: imageBase64,
            extension: 'png',
          });

          worksheet.addImage(imageId, {
            tl: { col: 1, row: rowIndex - 1 },
            ext: { width: 150, height: 150 },
            editAs: 'oneCell'
          });
        } catch (imgErr) {
          console.error("Failed to add image to Excel row", i, imgErr);
        }
      }

      // Vertical alignment for all cells
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          row.eachCell((cell) => {
            cell.alignment = { vertical: 'middle', wrapText: true };
          });
        }
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `arrowpointer-export-${Date.now()}.xlsx`);
    } catch (err) {
      console.error("Excel export failed", err);
      setError("Failed to create Excel file. Please try again.");
    } finally {
      setIsExcelDownloading(false);
    }
  };

  const downloadAsSheet = async () => {
    if (results.length === 0) return;
    setIsDownloading(true);

    try {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      const padding = 40;
      const labelHeight = 60;
      const columns = results.length > 4 ? 3 : 2;
      const itemWidth = 400;
      const itemHeight = 400 + labelHeight;
      
      const rows = Math.ceil(results.length / columns);
      
      canvas.width = columns * itemWidth + (columns + 1) * padding;
      canvas.height = rows * itemHeight + (rows + 1) * padding + 100; // Extra space for title

      // Background
      if (whiteBackground) {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      } else {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }

      // Title
      ctx.fillStyle = '#1A1A1A';
      ctx.font = 'bold 32px Inter, sans-serif';
      ctx.fillText('ArrowPointer: Detected Regions', padding, padding + 30);
      ctx.font = '16px Inter, sans-serif';
      ctx.fillStyle = '#6B7280';
      ctx.fillText(`Generated on ${new Date().toLocaleString()}`, padding, padding + 60);

      const startY = padding + 100;

      for (let i = 0; i < results.length; i++) {
        const res = results[i];
        const col = i % columns;
        const row = Math.floor(i / columns);
        
        const x = padding + col * (itemWidth + padding);
        const y = startY + row * (itemHeight + padding);

        // Load image
        const img = new Image();
        img.src = res.croppedUrl;
        await new Promise((resolve) => (img.onload = resolve));

        // Draw image (contained)
        const scale = Math.min(itemWidth / img.width, 400 / img.height);
        const drawW = img.width * scale;
        const drawH = img.height * scale;
        const offsetX = (itemWidth - drawW) / 2;
        const offsetY = (400 - drawH) / 2;

        if (!whiteBackground) {
          ctx.fillStyle = '#F3F4F6';
          ctx.fillRect(x, y, itemWidth, 400);
        }
        ctx.drawImage(img, x + offsetX, y + offsetY, drawW, drawH);

        // Label
        ctx.fillStyle = '#3b82f6';
        ctx.fillRect(x, y + 400, 40, 40);
        ctx.fillStyle = '#FFFFFF';
        ctx.font = 'bold 20px Inter, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(`${res.label}`, x + 20, y + 428);
        
        ctx.textAlign = 'left';
        ctx.fillStyle = '#1A1A1A';
        ctx.font = 'bold 16px Inter, sans-serif';
        ctx.fillText(`Region ${res.label}`, x + 50, y + 420);
        
        ctx.font = '12px Inter, sans-serif';
        ctx.fillStyle = '#3b82f6';
        ctx.fillText(`SAP: ${res.sap_code}`, x + 50, y + 438);
        ctx.fillText(`NAME: ${res.name}`, x + 50, y + 454);

        ctx.font = '14px Inter, sans-serif';
        ctx.fillStyle = '#4B5563';
        
        // Wrap text for description
        const words = res.description.split(' ');
        let line = '';
        let lineCount = 0;
        for (let n = 0; n < words.length; n++) {
          const testLine = line + words[n] + ' ';
          const metrics = ctx.measureText(testLine);
          if (metrics.width > itemWidth - 50 && n > 0) {
            ctx.fillText(line, x + 50, y + 474 + lineCount * 20);
            line = words[n] + ' ';
            lineCount++;
          } else {
            line = testLine;
          }
        }
        ctx.fillText(line, x + 50, y + 474 + lineCount * 20);
      }

      const link = document.createElement('a');
      link.download = `arrowpointer-results-${Date.now()}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
    } catch (err) {
      console.error("Download failed", err);
    } finally {
      setIsDownloading(false);
    }
  };

  const cleanBackgroundToWhite = (canvas: HTMLCanvasElement) => {
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = imageData.data;

    // Sample all four corners to find the most likely background color
    const corners = [
      { r: data[0], g: data[1], b: data[2] }, // Top-left
      { r: data[(canvas.width - 1) * 4], g: data[(canvas.width - 1) * 4 + 1], b: data[(canvas.width - 1) * 4 + 2] }, // Top-right
      { r: data[(data.length - canvas.width * 4)], g: data[(data.length - canvas.width * 4) + 1], b: data[(data.length - canvas.width * 4) + 2] }, // Bottom-left
      { r: data[data.length - 4], g: data[data.length - 3], b: data[data.length - 2] } // Bottom-right
    ];

    // Use the brightest corner as the reference background
    let ref = corners[0];
    let maxBrightness = (ref.r + ref.g + ref.b) / 3;
    for (const corner of corners) {
      const b = (corner.r + corner.g + corner.b) / 3;
      if (b > maxBrightness) {
        maxBrightness = b;
        ref = corner;
      }
    }

    // Tolerance for color matching
    const tolerance = 60;
    // Minimum brightness to consider as background
    const minBrightness = 150;

    for (let i = 0; i < data.length; i += 4) {
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      
      // Calculate distance from reference color
      const dist = Math.sqrt(
        Math.pow(r - ref.r, 2) + 
        Math.pow(g - ref.g, 2) + 
        Math.pow(b - ref.b, 2)
      );

      // Calculate brightness
      const brightness = (r + g + b) / 3;

      // If it's close to the reference color AND it's relatively bright
      // OR if it's very close to pure white (fallback)
      if ((dist < tolerance && brightness > minBrightness) || (r > 220 && g > 220 && b > 220)) {
        data[i] = 255;
        data[i + 1] = 255;
        data[i + 2] = 255;
        data[i + 3] = 255; // Pure White
      }
    }

    ctx.putImageData(imageData, 0, 0);
  };

  const downloadAllSeparate = async () => {
    const targetResults = allResults.length > 0 ? allResults : results;
    if (targetResults.length === 0) return;
    setIsZipping(true);

    try {
      const zip = new JSZip();
      
      for (let i = 0; i < targetResults.length; i++) {
        const res = targetResults[i];
        const base64Data = res.croppedUrl.split(',')[1];
        // Include page number in filename if available to avoid collisions
        const pageSuffix = res.pageNumber ? `_P${res.pageNumber}` : '';
        const fileName = `${res.sap_code !== 'Unknown' ? res.sap_code : `Item_${res.label}${pageSuffix}`}.png`;
        zip.file(fileName, base64Data, { base64: true });
      }

      const content = await zip.generateAsync({ type: 'blob' });
      saveAs(content, `arrowpointer-all-files-${Date.now()}.zip`);
    } catch (err) {
      console.error("ZIP creation failed", err);
      setError("Failed to create ZIP file for separate downloads.");
    } finally {
      setIsZipping(false);
    }
  };

  const processAllPages = async (autoExport = false) => {
    if (!pdfDoc) return;
    setIsProcessingAll(true);
    setAllResults([]);
    setProcessingProgress({ current: 0, total: pdfDoc.numPages });

    try {
      const accumulatedResults: CroppedResult[] = [];
      
      for (let p = 1; p <= pdfDoc.numPages; p++) {
        setProcessingProgress(prev => ({ ...prev, current: p }));
        
        // Render page to data URL
        const page = await pdfDoc.getPage(p);
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        if (!context) continue;
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        // @ts-ignore
        await page.render({ canvasContext: context, viewport }).promise;
        const pageImage = canvas.toDataURL('image/png');

        // Detect on this page
        const detections = await detectArrowsAndTargets(pageImage, 'image/png');
        
        // Process crops for this page
        const img = new Image();
        img.src = pageImage;
        await new Promise((resolve) => (img.onload = resolve));

        const cropCanvas = document.createElement('canvas');
        const cropCtx = cropCanvas.getContext('2d');
        if (!cropCtx) continue;

        for (const det of detections) {
          const { target_box } = det;
          const x = (target_box.xmin / 1000) * img.naturalWidth;
          const y = (target_box.ymin / 1000) * img.naturalHeight;
          const w = ((target_box.xmax - target_box.xmin) / 1000) * img.naturalWidth;
          const h = ((target_box.ymax - target_box.ymin) / 1000) * img.naturalHeight;
          const padding = Math.min(w, h) * 0.1;
          const cropX = Math.max(0, x - padding);
          const cropY = Math.max(0, y - padding);
          const cropW = Math.min(img.naturalWidth - cropX, w + padding * 2);
          const cropH = Math.min(img.naturalHeight - cropY, h + padding * 2);

          cropCanvas.width = cropW;
          cropCanvas.height = cropH;
          cropCtx.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH);
          
          if (whiteBackground) {
            cleanBackgroundToWhite(cropCanvas);
          }
          
          accumulatedResults.push({
            ...det,
            croppedUrl: cropCanvas.toDataURL('image/png'),
            pageNumber: p
          });
        }
      }
      
      setAllResults(accumulatedResults);
      // After processing all, show results for current page
      const currentPageDetections = accumulatedResults.filter(r => r.pageNumber === currentPage);
      setResults(currentPageDetections);
      
      // Auto export if requested
      if (autoExport && accumulatedResults.length > 0) {
        // We'll export to Excel by default as it seems the most requested format recently
        // But we wait a bit for state to settle
        setTimeout(() => {
          downloadAsExcel();
        }, 500);
      }
      
    } catch (err) {
      console.error(err);
      setError("An error occurred while processing all pages.");
    } finally {
      setIsProcessingAll(false);
    }
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setError(null);
      setResults([]);
      setAllResults([]);
      setPreviewUrl(null);
      setPdfDoc(null);
      setImage(null);
      
      const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
      const isImage = file.type.startsWith('image/');
      const isPpt = file.name.toLowerCase().endsWith('.pptx') || file.name.toLowerCase().endsWith('.ppt');

      if (isPdf) {
        setIsPdfLoading(true);
        try {
          const arrayBuffer = await file.arrayBuffer();
          const loadingTask = pdfjs.getDocument({ data: arrayBuffer });
          const pdf = await loadingTask.promise;
          setPdfDoc(pdf);
          setCurrentPage(1);
          await renderPdfPage(pdf, 1);
        } catch (err) {
          console.error(err);
          setError("Failed to load PDF file. Please ensure it's a valid PDF.");
        } finally {
          setIsPdfLoading(false);
        }
      } else if (isImage) {
        const reader = new FileReader();
        reader.onload = (event) => {
          setImage(event.target?.result as string);
          setMimeType(file.type || 'image/png');
        };
        reader.readAsDataURL(file);
      } else if (isPpt) {
        setError("PowerPoint files (.pptx) cannot be read directly by the browser. Please save your PPT as a PDF (File > Export > PDF) and upload the PDF here!");
      } else {
        setError("Unsupported file type. Please upload an image (JPG, PNG) or a PDF.");
      }
    }
  };

  const renderPdfPage = async (pdf: pdfjs.PDFDocumentProxy, pageNum: number) => {
    try {
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      if (!context) return;

      canvas.height = viewport.height;
      canvas.width = viewport.width;

      // @ts-ignore - pdfjs-dist types can be tricky
      await page.render({ canvasContext: context, viewport }).promise;
      const dataUrl = canvas.toDataURL('image/png');
      setImage(dataUrl);
      setMimeType('image/png');
    } catch (err) {
      console.error(err);
      setError("Failed to render PDF page.");
    }
  };

  const handlePageChange = async (delta: number) => {
    if (!pdfDoc) return;
    const newPage = currentPage + delta;
    if (newPage >= 1 && newPage <= pdfDoc.numPages) {
      setCurrentPage(newPage);
      setPreviewUrl(null);
      
      // If we have allResults, show them for this page
      if (allResults.length > 0) {
        const pageResults = allResults.filter(r => r.pageNumber === newPage);
        setResults(pageResults);
      } else {
        setResults([]);
      }
      
      await renderPdfPage(pdfDoc, newPage);
    }
  };

  const processImage = async () => {
    if (!image) return;

    setIsProcessing(true);
    setError(null);
    try {
      const detections = await detectArrowsAndTargets(image, mimeType);
      
      if (detections.length === 0) {
        setError("No arrows detected in the image.");
        setIsProcessing(false);
        return;
      }

      // Create crops and preview
      const img = new Image();
      img.src = image;
      await new Promise((resolve) => (img.onload = resolve));

      const croppedResults: CroppedResult[] = [];
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');

      if (!ctx) throw new Error("Could not get canvas context");

      // Sort detections by their detected label if numeric, otherwise spatial
      const sortedDetections = [...detections].sort((a, b) => {
        const aLabel = parseInt(String(a.label));
        const bLabel = parseInt(String(b.label));
        
        if (!isNaN(aLabel) && !isNaN(bLabel)) {
          return aLabel - bLabel;
        }
        
        const yDiff = a.target_box.ymin - b.target_box.ymin;
        if (Math.abs(yDiff) > 50) return yDiff;
        return a.target_box.xmin - b.target_box.xmin;
      });

      for (const detection of sortedDetections) {
        const { target_box } = detection;
        
        // Convert normalized to pixel coordinates
        const x = (target_box.xmin / 1000) * img.naturalWidth;
        const y = (target_box.ymin / 1000) * img.naturalHeight;
        const w = ((target_box.xmax - target_box.xmin) / 1000) * img.naturalWidth;
        const h = ((target_box.ymax - target_box.ymin) / 1000) * img.naturalHeight;

        // Add some padding to the crop
        const padding = Math.min(w, h) * 0.1;
        const cropX = Math.max(0, x - padding);
        const cropY = Math.max(0, y - padding);
        const cropW = Math.min(img.naturalWidth - cropX, w + padding * 2);
        const cropH = Math.min(img.naturalHeight - cropY, h + padding * 2);

        canvas.width = cropW;
        canvas.height = cropH;
        ctx.drawImage(img, cropX, cropY, cropW, cropH, 0, 0, cropW, cropH);
        
        if (whiteBackground) {
          cleanBackgroundToWhite(canvas);
        }
        
        croppedResults.push({
          ...detection,
          croppedUrl: canvas.toDataURL('image/png')
        });
      }

      // Create preview with bounding boxes
      const previewCanvas = document.createElement('canvas');
      previewCanvas.width = img.naturalWidth;
      previewCanvas.height = img.naturalHeight;
      const pCtx = previewCanvas.getContext('2d');
      if (pCtx) {
        pCtx.drawImage(img, 0, 0);
        pCtx.lineWidth = Math.max(2, img.naturalWidth / 300);
        pCtx.font = `bold ${Math.max(16, img.naturalWidth / 40)}px Inter, sans-serif`;
        
        croppedResults.forEach((res, index) => {
          const { target_box, arrow_box } = res;
          
          // Draw target box
          const tx = (target_box.xmin / 1000) * img.naturalWidth;
          const ty = (target_box.ymin / 1000) * img.naturalHeight;
          const tw = ((target_box.xmax - target_box.xmin) / 1000) * img.naturalWidth;
          const th = ((target_box.ymax - target_box.ymin) / 1000) * img.naturalHeight;
          
          pCtx.strokeStyle = '#3b82f6'; // Blue
          pCtx.strokeRect(tx, ty, tw, th);
          
          // Draw arrow box (optional but helpful)
          const ax = (arrow_box.xmin / 1000) * img.naturalWidth;
          const ay = (arrow_box.ymin / 1000) * img.naturalHeight;
          const aw = ((arrow_box.xmax - arrow_box.xmin) / 1000) * img.naturalWidth;
          const ah = ((arrow_box.ymax - arrow_box.ymin) / 1000) * img.naturalHeight;
          
          pCtx.strokeStyle = '#ef4444'; // Red
          pCtx.setLineDash([5, 5]);
          pCtx.strokeRect(ax, ay, aw, ah);
          pCtx.setLineDash([]);

          // Draw label
          pCtx.fillStyle = '#3b82f6';
          const labelText = `${res.label}`;
          const textMetrics = pCtx.measureText(labelText);
          const labelPadding = 8;
          pCtx.fillRect(tx, ty - textMetrics.actualBoundingBoxAscent - labelPadding * 2, textMetrics.width + labelPadding * 2, textMetrics.actualBoundingBoxAscent + labelPadding * 2);
          
          pCtx.fillStyle = 'white';
          pCtx.fillText(labelText, tx + labelPadding, ty - labelPadding);
        });
        
        setPreviewUrl(previewCanvas.toDataURL(mimeType));
      }

      setResults(croppedResults);
    } catch (err) {
      console.error(err);
      setError("An error occurred while processing the image.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#F8F9FA] text-[#1A1A1A] font-sans selection:bg-blue-100">
      {/* Header */}
      <header className="border-b border-gray-200 bg-white sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
              <ArrowRight className="text-white w-5 h-5 -rotate-45" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">ArrowPointer</h1>
          </div>
          <div className="hidden sm:flex items-center gap-6 text-sm font-medium text-gray-500">
            <span className="flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-green-500" /> Detect</span>
            <span className="flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-green-500" /> Map</span>
            <span className="flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-green-500" /> Crop</span>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          
          {/* Left Column: Upload & Main View */}
          <div className="lg:col-span-7 space-y-6">
            <section className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden">
              <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                <h2 className="text-sm font-semibold uppercase tracking-wider text-gray-500 flex items-center gap-2">
                  <ImageIcon className="w-4 h-4" /> Source Image
                </h2>
                {image && !isProcessing && !isProcessingAll && (
                  <div className="flex items-center gap-4">
                    <label className="flex items-center gap-2 cursor-pointer group">
                      <div className={`w-10 h-5 rounded-full transition-all relative ${whiteBackground ? 'bg-blue-600' : 'bg-gray-300'}`}>
                        <input 
                          type="checkbox" 
                          className="sr-only" 
                          checked={whiteBackground}
                          onChange={(e) => setWhiteBackground(e.target.checked)}
                        />
                        <div className={`absolute top-1 w-3 h-3 bg-white rounded-full transition-all ${whiteBackground ? 'left-6' : 'left-1'}`} />
                      </div>
                      <span className="text-xs font-semibold text-gray-600 group-hover:text-blue-600 transition-colors">White Background</span>
                    </label>

                    <div className="h-6 w-px bg-gray-200" />

                    <div className="flex items-center gap-2">
                      {pdfDoc && (
                        <div className="flex flex-col items-end gap-1">
                          <button 
                            onClick={() => processAllPages(true)}
                            className="bg-green-600 hover:bg-green-700 text-white px-6 py-2 rounded-full text-sm font-bold transition-all flex items-center gap-2 shadow-lg shadow-green-200 animate-pulse hover:animate-none"
                          >
                            <FileSpreadsheet className="w-5 h-5" /> Analyze All & Export Excel (One-Click)
                          </button>
                          <span className="text-[10px] text-gray-400 font-medium px-2">Best for catalogs with multiple pages</span>
                        </div>
                      )}
                      <div className="h-8 w-px bg-gray-200 mx-2" />
                      <button 
                        onClick={processImage}
                        className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-full text-sm font-medium transition-all flex items-center gap-2 shadow-lg shadow-blue-200"
                      >
                        Analyze Current Page <ArrowRight className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                )}
              </div>
              
              <div className="p-6">
                {!image ? (
                  <div 
                    onClick={() => fileInputRef.current?.click()}
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={async (e) => {
                      e.preventDefault();
                      const file = e.dataTransfer.files[0];
                      if (file) {
                        // Trigger the same logic as handleFileChange
                        const mockEvent = { target: { files: [file] } } as any;
                        handleFileChange(mockEvent);
                      }
                    }}
                    className="border-2 border-dashed border-gray-200 rounded-xl p-12 flex flex-col items-center justify-center gap-4 cursor-pointer hover:border-blue-400 hover:bg-blue-50/30 transition-all group"
                  >
                    <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center group-hover:scale-110 transition-transform">
                      <Upload className="text-gray-400 group-hover:text-blue-500 w-8 h-8" />
                    </div>
                    <div className="text-center">
                      <p className="text-lg font-medium text-gray-900">Drop your image or PDF here</p>
                      <p className="text-sm text-gray-500 mb-4">or click to browse. Export PPT as PDF for best results!</p>
                    </div>
                    <input 
                      type="file" 
                      ref={fileInputRef} 
                      onChange={handleFileChange} 
                      className="hidden" 
                      accept="image/*,.pdf,.pptx,.ppt"
                    />
                  </div>
                ) : (
                  <div className="relative rounded-lg overflow-hidden bg-gray-900 flex flex-col items-center justify-center min-h-[400px]">
                    <img 
                      ref={imageRef}
                      src={previewUrl || image} 
                      alt="Source" 
                      className="max-w-full h-auto object-contain"
                      referrerPolicy="no-referrer"
                    />
                    
                    {pdfDoc && (
                      <div className="absolute bottom-4 left-1/2 -translate-x-1/2 bg-black/60 backdrop-blur-md px-4 py-2 rounded-full flex items-center gap-4 text-white border border-white/10 shadow-xl">
                        <button 
                          onClick={() => handlePageChange(-1)}
                          disabled={currentPage === 1 || isProcessing}
                          className="p-1 hover:bg-white/20 rounded-full disabled:opacity-30 transition-colors"
                        >
                          <ChevronLeft className="w-5 h-5" />
                        </button>
                        <span className="text-sm font-medium tabular-nums">
                          Page {currentPage} of {pdfDoc.numPages}
                        </span>
                        <button 
                          onClick={() => handlePageChange(1)}
                          disabled={currentPage === pdfDoc.numPages || isProcessing}
                          className="p-1 hover:bg-white/20 rounded-full disabled:opacity-30 transition-colors"
                        >
                          <ChevronRight className="w-5 h-5" />
                        </button>
                      </div>
                    )}

                    {(isProcessing || isPdfLoading || isProcessingAll) && (
                      <div className="absolute inset-0 bg-black/40 backdrop-blur-sm flex flex-col items-center justify-center text-white gap-4">
                        <Loader2 className="w-10 h-10 animate-spin text-blue-400" />
                        <p className="font-medium animate-pulse">
                          {isPdfLoading ? "Rendering PDF page..." : 
                           isProcessingAll ? `Processing page ${processingProgress.current} of ${processingProgress.total}...` :
                           "Analyzing visual cues..."}
                        </p>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </section>

            {error && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-red-50 border border-red-100 p-4 rounded-xl space-y-3 text-red-700"
              >
                <div className="flex items-center gap-3">
                  <AlertCircle className="w-5 h-5 shrink-0" />
                  <p className="text-sm font-medium">{error}</p>
                </div>
                {error.includes("PowerPoint") && (
                  <div className="bg-white/50 p-3 rounded-lg border border-red-200 text-xs text-red-600 space-y-2">
                    <p className="font-bold uppercase tracking-wider">How to convert PPT to PDF:</p>
                    <ol className="list-decimal list-inside space-y-1">
                      <li>Open your PowerPoint file</li>
                      <li>Go to <b>File {'>'} Export</b> (or Save As)</li>
                      <li>Select <b>PDF</b> as the file format</li>
                      <li>Upload that PDF file here!</li>
                    </ol>
                  </div>
                )}
              </motion.div>
            )}
          </div>

          {/* Right Column: Results */}
          <div className="lg:col-span-5 space-y-6">
            <section className="bg-white rounded-2xl border border-gray-200 shadow-sm flex flex-col h-full min-h-[600px]">
              <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                <h2 className="text-sm font-semibold uppercase tracking-wider text-gray-500 flex items-center gap-2">
                  <List className="w-4 h-4" /> Detected Regions
                </h2>
                <div className="flex items-center gap-2">
                  {(results.length > 0 || allResults.length > 0) && (
                    <>
                      <button 
                        onClick={downloadAllSeparate}
                        disabled={isZipping}
                        className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all flex items-center gap-1.5 text-xs font-medium"
                        title={allResults.length > 0 ? "Download all pages as separate files (ZIP)" : "Download current page as separate files (ZIP)"}
                      >
                        {isZipping ? <Loader2 className="w-4 h-4 animate-spin" /> : <Package className="w-4 h-4" />}
                        <span className="hidden sm:inline">{allResults.length > 0 ? 'Download All Pages' : 'Download Page'}</span>
                      </button>
                      <button 
                        onClick={downloadAsSheet}
                        disabled={isDownloading}
                        className="p-1.5 text-gray-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all flex items-center gap-1.5 text-xs font-medium"
                        title="Download results as a single sheet"
                      >
                        {isDownloading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Download className="w-4 h-4" />}
                        <span className="hidden sm:inline">Export Sheet</span>
                      </button>
                      <button 
                        onClick={downloadAsExcel}
                        disabled={isExcelDownloading}
                        className="p-1.5 text-gray-400 hover:text-green-600 hover:bg-green-50 rounded-lg transition-all flex items-center gap-1.5 text-xs font-medium"
                        title="Export all data and images to Excel"
                      >
                        {isExcelDownloading ? <Loader2 className="w-4 h-4 animate-spin" /> : <FileSpreadsheet className="w-4 h-4" />}
                        <span className="hidden sm:inline">Export Excel</span>
                      </button>
                    </>
                  )}
                <span className="text-xs font-mono bg-gray-200 px-2 py-0.5 rounded text-gray-600">
                  {allResults.length > 0 ? `${allResults.length} TOTAL` : `${results.length} FOUND`}
                </span>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
                <AnimatePresence mode="popLayout">
                  {results.length === 0 && !isProcessing && (
                    <motion.div 
                      key="empty"
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      className="h-full flex flex-col items-center justify-center text-gray-400 py-20"
                    >
                      <Maximize2 className="w-12 h-12 mb-4 opacity-20" />
                      <p className="text-sm">Processed regions will appear here</p>
                    </motion.div>
                  )}

                  {results.map((res, idx) => (
                    <motion.div
                      key={idx}
                      initial={{ opacity: 0, x: 20 }}
                      animate={{ opacity: 1, x: 0 }}
                      transition={{ delay: idx * 0.1 }}
                      className="group bg-white border border-gray-100 rounded-xl overflow-hidden hover:border-blue-200 hover:shadow-md transition-all"
                    >
                      <div className="flex gap-4 p-3">
                        <div className="relative w-32 h-32 shrink-0 bg-white rounded-lg overflow-hidden border border-gray-100 group-hover:border-blue-200 transition-all">
                          <img 
                            src={res.croppedUrl} 
                            alt={`Crop ${idx + 1}`} 
                            className="w-full h-full object-contain relative z-10"
                            referrerPolicy="no-referrer"
                          />
                          <div className="absolute top-1 left-1 bg-blue-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded shadow-sm">
                            #{res.label}
                          </div>
                        </div>
                        <div className="flex flex-col justify-center py-1 flex-1">
                          <h3 className="font-bold text-gray-900 mb-1 flex items-center gap-2">
                            Region {res.label}
                          </h3>
                          
                          <div className="flex flex-wrap gap-2 mb-2">
                            <div className="flex items-center gap-1 bg-blue-50 text-blue-700 px-2 py-0.5 rounded text-[10px] font-bold">
                              <Tag className="w-3 h-3" /> SAP: {res.sap_code}
                            </div>
                            <div className="flex items-center gap-1 bg-gray-50 text-gray-700 px-2 py-0.5 rounded text-[10px] font-bold">
                              <User className="w-3 h-3" /> NAME: {res.name}
                            </div>
                          </div>

                          <p className="text-sm text-gray-600 leading-relaxed line-clamp-2">
                            {res.description}
                          </p>
                          <div className="mt-2 flex items-center gap-3">
                            <span className="text-[10px] font-mono text-gray-400 uppercase tracking-tighter">
                              COORDS: {Math.round(res.target_box.xmin)},{Math.round(res.target_box.ymin)}
                            </span>
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  ))}
                </AnimatePresence>
              </div>

              {results.length > 0 && (
                <div className="p-4 border-t border-gray-100 bg-gray-50/50">
                  <button 
                    onClick={() => setImage(null)}
                    className="w-full py-2 text-sm font-medium text-gray-500 hover:text-gray-900 transition-colors"
                  >
                    Clear and Start Over
                  </button>
                </div>
              )}
            </section>
          </div>

        </div>
      </main>

      <style dangerouslySetInnerHTML={{ __html: `
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #E5E7EB;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #D1D5DB;
        }
      `}} />
    </div>
  );
}
