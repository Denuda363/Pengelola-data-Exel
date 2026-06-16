/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import { Upload, FileUp, FileSpreadsheet, Trash2, HelpCircle } from 'lucide-react';
import { ExcelWorkbookData } from '../types';
import { parseExcelFile, generateSampleData } from '../utils/excelHelper';

interface ExcelImporterProps {
  onDataLoaded: (data: ExcelWorkbookData | null) => void;
  currentData: ExcelWorkbookData | null;
}

export default function ExcelImporter({ onDataLoaded, currentData }: ExcelImporterProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  const processFile = (file: File) => {
    setError(null);
    const suffix = file.name.split('.').pop()?.toLowerCase();
    
    if (suffix !== 'xlsx' && suffix !== 'xls' && suffix !== 'csv') {
      setError('Format file tidak didukung. Silakan unggah file dengan format .xlsx, .xls, atau .csv');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (data instanceof ArrayBuffer) {
          const workbookData = parseExcelFile(data, file.name, file.size);
          onDataLoaded(workbookData);
        } else {
          setError('Terjadi kesalahan saat membaca file.');
        }
      } catch (err: any) {
        console.error(err);
        setError(`Gagal memproses file Excel: ${err.message || 'File mungkin korup atau tidak valid'}`);
      }
    };
    reader.onerror = () => {
      setError('Gagal membaca file.');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  const loadDemo = () => {
    setError(null);
    const sample = generateSampleData();
    onDataLoaded(sample);
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  const removeFile = () => {
    onDataLoaded(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  return (
    <div className="w-full">
      {!currentData ? (
        <div
          id="dropzone"
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={triggerFileSelect}
          className={`relative border-2 border-dashed rounded-xl p-8 flex flex-col items-center justify-center cursor-pointer transition-all duration-250 ${
            isDragging
              ? 'border-emerald-500 bg-emerald-500/5 scale-[1.01]'
              : 'border-slate-800 bg-slate-900/40 hover:bg-slate-900/60 hover:border-slate-700'
          }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".xlsx, .xls, .csv"
            className="hidden"
          />

          <div className="p-4 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-full mb-4">
            <Upload className="w-7 h-7" />
          </div>

          <p className="text-base font-semibold text-slate-200 text-center mb-1 font-sans">
            Tarik &amp; Lepas File Excel Anda Di Sini
          </p>
          <p className="text-sm text-slate-400 text-center mb-4">
            atau <span className="text-emerald-400 font-medium hover:underline hover:text-emerald-300">pilih file dari perangkat</span> (.xlsx, .xls, .csv)
          </p>

          <p className="text-xs text-slate-500 text-center max-w-md leading-relaxed font-sans">
            Keamanan Terjamin: Berkas diproses sepenuhnya secara lokal pada peramban Anda dan tidak pernah dikirimkan ke server.
          </p>

          {error && (
            <div className="mt-4 p-3 bg-red-950/40 text-xs text-red-400 rounded-lg max-w-xl text-center border border-red-800/40">
              {error}
            </div>
          )}

          <div className="mt-6 flex flex-wrap gap-3 justify-center items-center">
            <div className="text-xs text-slate-500 font-medium mr-1 font-sans">Belum memiliki file excel?</div>
            <button
              id="btn-demo-data"
              type="button"
              onClick={(e) => {
                e.stopPropagation();
                loadDemo();
              }}
              className="px-4 py-2 bg-slate-800 hover:bg-slate-700 hover:text-white border border-slate-705/80 text-slate-300 font-semibold text-xs rounded transition-all flex items-center gap-1.5"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
              Gunakan Data Sampel Penjualan
            </button>
          </div>
        </div>
      ) : (
        <div id="file-info-header" className="bg-slate-900 border border-slate-800 rounded-xl p-4 md:p-5 shadow-lg flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-start gap-4">
            <div className="p-3 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-xl">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-serif italic text-white line-clamp-1">{currentData.fileName}</h3>
              <div className="flex flex-wrap items-center gap-x-3 gap-y-1 mt-1 text-[11px] text-slate-400">
                <span>Ukuran: <strong className="font-semibold text-slate-200 font-mono text-xs">{currentData.fileSize}</strong></span>
                <span className="text-slate-700">•</span>
                <span>Lembar Kerja: <strong className="font-semibold text-slate-200 font-mono text-xs">{currentData.sheets.length}</strong></span>
                <span className="text-slate-700">•</span>
                <span>Aktif: <strong className="font-semibold text-emerald-400 font-mono text-xs">{currentData.sheets[currentData.activeSheetIndex]?.name}</strong></span>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 self-end md:self-center">
            <button
              id="btn-ganti-file"
              onClick={triggerFileSelect}
              className="px-3.5 py-1.5 border border-slate-700 hover:border-slate-600 bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-300 rounded transition-all flex items-center gap-1.5 cursor-pointer hover:text-white"
            >
              <FileUp className="w-3.5 h-3.5" />
              Ganti Berkas
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              accept=".xlsx, .xls, .csv"
              className="hidden"
            />
            
            <button
              id="btn-hapus-file"
              onClick={removeFile}
              className="p-2 bg-red-950/30 hover:bg-red-950/60 text-red-400 hover:text-red-300 border border-red-900/30 rounded transition-all cursor-pointer"
              title="Hapus file"
            >
              <Trash2 className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
