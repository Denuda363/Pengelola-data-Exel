/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect, useRef } from 'react';
import { FileSpreadsheet, Printer, Table, Library, ChevronRight, HelpCircle, Columns } from 'lucide-react';
import { ExcelWorkbookData, FilterRule, SortConfig } from './types';
import ExcelImporter from './components/ExcelImporter';
import ExcelFilters from './components/ExcelFilters';
import ExcelTable from './components/ExcelTable';
import ExcelPrintTemplate from './components/ExcelPrintTemplate';

export default function App() {
  const [workbook, setWorkbook] = useState<ExcelWorkbookData | null>(null);
  const [filterRules, setFilterRules] = useState<FilterRule[]>([]);
  const [quickSearch, setQuickSearch] = useState<string>('');
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [printedColumnKeys, setPrintedColumnKeys] = useState<string[]>([]);
  const [showColDropdown, setShowColDropdown] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Extract current sheet data
  const currentSheet = useMemo(() => {
    if (!workbook || workbook.sheets.length === 0) return null;
    return workbook.sheets[workbook.activeSheetIndex] || workbook.sheets[0];
  }, [workbook]);

  // Sync printedColumnKeys with currentSheet columns when sheet/data changes
  useEffect(() => {
    if (currentSheet) {
      setPrintedColumnKeys(currentSheet.columns.map(c => c.key));
    } else {
      setPrintedColumnKeys([]);
    }
  }, [currentSheet]);

  // Click outside to close column dropdown picker
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowColDropdown(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // Handle new file loaded or reset
  const handleDataLoaded = (data: ExcelWorkbookData | null) => {
    setWorkbook(data);
    setFilterRules([]);
    setQuickSearch('');
    setSortConfig(null);
  };

  // Change active sheet of current workbook
  const handleSheetChange = (sheetIndex: number) => {
    if (!workbook) return;
    setWorkbook({
      ...workbook,
      activeSheetIndex: sheetIndex,
    });
    setFilterRules([]);
    setSortConfig(null);
  };

  // Execute actual filtering logic
  const filteredRows = useMemo(() => {
    if (!currentSheet) return [];
    let rows = [...currentSheet.rows];

    // 1. Process column filter rules
    if (filterRules.length > 0) {
      rows = rows.filter(row => {
        return filterRules.every(rule => {
          const colValue = row[rule.columnKey];
          const colType = currentSheet.columns.find(c => c.key === rule.columnKey)?.type || 'string';
          const ruleValue = rule.value;

          // Operators "empty" and "not_empty" do not require rule values
          if (rule.operator === 'empty') {
            return colValue === null || colValue === undefined || colValue === '';
          }
          if (rule.operator === 'not_empty') {
            return colValue !== null && colValue !== undefined && colValue !== '';
          }

          if (colValue === null || colValue === undefined) return false;

          const strVal = String(colValue).toLowerCase().trim();
          const strRuleVal = String(ruleValue).toLowerCase().trim();

          switch (rule.operator) {
            case 'contains':
              return strVal.includes(strRuleVal);
            case 'equals':
              if (colType === 'number') {
                return Number(colValue) === Number(ruleValue);
              }
              if (colType === 'boolean') {
                const isTrueVal = colValue === true || String(colValue).toUpperCase() === 'TRUE' || colValue === 1;
                const isTrueRule = ruleValue === 'true';
                return isTrueVal === isTrueRule;
              }
              return strVal === strRuleVal;
            case 'not_equals':
              if (colType === 'number') {
                return Number(colValue) !== Number(ruleValue);
              }
              return strVal !== strRuleVal;
            case 'starts_with':
              return strVal.startsWith(strRuleVal);
            case 'ends_with':
              return strVal.endsWith(strRuleVal);
            case 'greater_than':
              if (colType === 'number') {
                return Number(colValue) > Number(ruleValue);
              }
              if (colType === 'date') {
                return new Date(colValue).getTime() > new Date(ruleValue).getTime();
              }
              return colValue > ruleValue;
            case 'less_than':
              if (colType === 'number') {
                return Number(colValue) < Number(ruleValue);
              }
              if (colType === 'date') {
                return new Date(colValue).getTime() < new Date(ruleValue).getTime();
              }
              return colValue < ruleValue;
            case 'one_of': {
              try {
                const list = JSON.parse(ruleValue) as (string | number)[];
                if (!Array.isArray(list) || list.length === 0) return false;
                const cellValStr = (colValue === null || colValue === undefined || colValue === '') ? '—' : String(colValue).trim();
                return list.some(item => String(item).toLowerCase().trim() === cellValStr.toLowerCase());
              } catch (e) {
                return false;
              }
            }
            case 'not_one_of': {
              try {
                const list = JSON.parse(ruleValue) as (string | number)[];
                if (!Array.isArray(list) || list.length === 0) return true;
                const cellValStr = (colValue === null || colValue === undefined || colValue === '') ? '—' : String(colValue).trim();
                return !list.some(item => String(item).toLowerCase().trim() === cellValStr.toLowerCase());
              } catch (e) {
                return true;
              }
            }
            default:
              return true;
          }
        });
      });
    }

    // 2. Process global quick search
    if (quickSearch) {
      const searchTerms = quickSearch.toLowerCase().trim().split(' ');
      rows = rows.filter(row => {
        // Must contain all terms in any of the cellular fields
        return searchTerms.every(term => {
          return Object.values(row).some(cellVal => {
            if (cellVal === null || cellVal === undefined) return false;
            return String(cellVal).toLowerCase().includes(term);
          });
        });
      });
    }

    // 3. Process column-specific sorting
    if (sortConfig) {
      const { columnKey, direction } = sortConfig;
      const colType = currentSheet.columns.find(c => c.key === columnKey)?.type || 'string';

      rows.sort((a, b) => {
        const valA = a[columnKey];
        const valB = b[columnKey];

        // Handle empty values priority
        if (valA === '' || valA === null || valA === undefined) return direction === 'asc' ? 1 : -1;
        if (valB === '' || valB === null || valB === undefined) return direction === 'asc' ? -1 : 1;

        if (colType === 'number') {
          return direction === 'asc' ? Number(valA) - Number(valB) : Number(valB) - Number(valA);
        }

        if (colType === 'date') {
          const dateA = new Date(valA).getTime();
          const dateB = new Date(valB).getTime();
          if (!isNaN(dateA) && !isNaN(dateB)) {
            return direction === 'asc' ? dateA - dateB : dateB - dateA;
          }
        }

        // Default string fallback sorting
        const strA = String(valA).toLowerCase();
        const strB = String(valB).toLowerCase();
        return direction === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
      });
    }

    return rows;
  }, [currentSheet, filterRules, quickSearch, sortConfig]);

  // Triggers browser native print dialog
  const printFilteredData = () => {
    window.print();
  };

  return (
    <div id="app-root-container" className="min-h-screen bg-slate-950 py-6 px-4 sm:px-6 lg:px-8 text-slate-200 antialiased selection:bg-emerald-500/25 selection:text-emerald-300">
      
      {/* Printable template - ONLY shown on print @media */}
      {workbook && currentSheet && (
        <ExcelPrintTemplate
          fileName={workbook.fileName}
          sheetName={currentSheet.name}
          columns={currentSheet.columns.filter(c => printedColumnKeys.includes(c.key))}
          rows={filteredRows}
          activeRules={filterRules}
          quickSearchQuery={quickSearch}
        />
      )}

      {/* Screen Interface - hidden on printed pages */}
      <div className="max-w-7xl mx-auto space-y-6 print:hidden">
        
        {/* Header Block */}
        <header id="main-header" className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800 pb-5">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-emerald-500/20 border border-emerald-500/50 rounded flex items-center justify-center text-emerald-400">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-xl font-serif italic tracking-tight text-white">
                Pengelola File Excel
              </h1>
              <p className="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-bold">
                Aplikasi Pengolahan &amp; Filter Data Cerdas
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <div className="text-right hidden sm:block">
              <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest font-mono">Status Sistem</p>
              <div className="flex items-center gap-1.5 justify-end mt-0.5">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                <span className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">Sistem Privat / Siap</span>
              </div>
            </div>
          </div>
        </header>

        {/* File Importer Widget */}
        <section id="module-importer" className="transition-all duration-300">
          <ExcelImporter onDataLoaded={handleDataLoaded} currentData={workbook} />
        </section>

        {workbook && currentSheet ? (
          <>
            {/* Control Panel: Sheet switching & Global Statistics & PRINT BUTTON */}
            <div id="module-dashboard-actions" className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-lg flex flex-col md:flex-row md:items-center justify-between gap-4">
              
              {/* Sheet Navigation Tabs */}
              <div className="flex items-center gap-1.5 overflow-x-auto pb-2 md:pb-0 scrollbar-none">
                <div className="flex items-center gap-1 text-slate-500 mr-2 text-[10px] font-bold uppercase tracking-wider flex-shrink-0">
                  <Library className="w-3.5 h-3.5" />
                  Lembar Kerja:
                </div>
                {workbook.sheets.map((sheet, index) => {
                  const isActive = workbook.activeSheetIndex === index;
                  return (
                    <button
                      key={index}
                      id={`btn-sheet-tab-${index}`}
                      onClick={() => handleSheetChange(index)}
                      className={`px-3.5 py-1.5 text-xs font-semibold rounded transition-all flex-shrink-0 flex items-center gap-1.5 cursor-pointer ${
                        isActive
                          ? 'bg-emerald-600 text-white shadow-lg shadow-emerald-950/40'
                          : 'bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-705/50 hover:text-white'
                      }`}
                    >
                      <Table className="w-3.5 h-3.5" />
                      {sheet.name}
                      <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono font-bold ${isActive ? 'bg-emerald-500 text-white' : 'bg-slate-950 text-slate-500'}`}>
                        {sheet.rows.length}
                      </span>
                    </button>
                  );
                })}
              </div>

              {/* Print Data Control & Statistics */}
              <div className="flex items-center justify-between md:justify-end gap-3 border-t md:border-t-0 border-slate-800 pt-3 md:pt-0">
                <div className="text-left font-sans mr-2">
                  <span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider">Hasil Filter</span>
                  <span id="filtered-matches-tag" className="text-xs text-slate-400 font-medium">
                    Lolos: <strong className="text-emerald-400 font-bold font-mono">{filteredRows.length}</strong> / <span className="font-mono">{currentSheet.rows.length}</span> baris
                  </span>
                </div>

                {/* Print Column Selector */}
                <div className="relative" ref={dropdownRef}>
                  <button
                    id="btn-pilih-kolom-cetak"
                    type="button"
                    onClick={() => setShowColDropdown(prev => !prev)}
                    className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700/80 rounded flex items-center gap-1.5 transition-colors cursor-pointer text-xs font-semibold"
                  >
                    <Columns className="w-3.5 h-3.5 text-emerald-400" />
                    Kolom Cetak ({printedColumnKeys.length}/{currentSheet.columns.length})
                  </button>

                  {showColDropdown && (
                    <div className="absolute right-0 mt-2 w-64 bg-slate-900 border border-slate-800 rounded-lg shadow-2xl p-3 z-50 animate-fadeIn text-slate-100 flex flex-col gap-2">
                      <div className="flex items-center justify-between border-b border-slate-850 pb-2">
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider font-mono">PILIH KOLOM</span>
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={() => setPrintedColumnKeys(currentSheet.columns.map(c => c.key))}
                            className="text-[9px] text-emerald-400 hover:text-emerald-300 font-bold uppercase tracking-wider bg-transparent p-0 cursor-pointer"
                          >
                            Semua
                          </button>
                          <span className="text-slate-800 text-[10px]">|</span>
                          <button
                            type="button"
                            onClick={() => setPrintedColumnKeys([])}
                            className="text-[9px] text-red-400 hover:text-red-300 font-bold uppercase tracking-wider bg-transparent p-0 cursor-pointer"
                          >
                            Kosong
                          </button>
                        </div>
                      </div>

                      <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1 scrollbar-thin">
                        {currentSheet.columns.map((col) => {
                          const isChecked = printedColumnKeys.includes(col.key);
                          return (
                            <label
                              key={col.key}
                              className={`flex items-center gap-2.5 px-2 py-1.5 rounded hover:bg-slate-855 transition-colors cursor-pointer select-none text-[11px] border ${
                                isChecked ? 'bg-emerald-500/5 border-emerald-500/20 text-emerald-300' : 'border-transparent text-slate-400 hover:text-slate-200'
                              }`}
                            >
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => {
                                  if (isChecked) {
                                    setPrintedColumnKeys(prev => prev.filter(k => k !== col.key));
                                  } else {
                                    setPrintedColumnKeys(prev => [...prev, col.key]);
                                  }
                                }}
                                className="rounded border-slate-700 bg-slate-950 text-emerald-600 focus:ring-emerald-500/30 focus:ring-opacity-50 h-3.5 w-3.5 cursor-pointer"
                              />
                              <span className="flex-1 truncate">{col.name}</span>
                            </label>
                          );
                        })}
                      </div>

                      <div className="border-t border-slate-850 pt-2 flex items-center justify-between text-[9px] text-slate-500 italic">
                        <span>Kolom yang dicetak</span>
                        <span className="font-mono text-emerald-500/70 font-semibold">{printedColumnKeys.length} Aktif</span>
                      </div>
                    </div>
                  )}
                </div>

                <button
                  id="btn-cetak-hasil"
                  type="button"
                  onClick={printFilteredData}
                  disabled={printedColumnKeys.length === 0}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:bg-slate-800/80 disabled:text-slate-500 disabled:cursor-not-allowed text-white text-xs font-semibold rounded shadow-lg shadow-emerald-950/20 flex items-center gap-2 transition-all cursor-pointer group"
                >
                  <Printer className="w-3.5 h-3.5 text-white group-hover:rotate-6 transition-transform" />
                  Cetak Hasil Filter
                </button>
              </div>

            </div>

            {/* Filter Conditions Builder Panel */}
            <section id="module-filters">
              <ExcelFilters
                columns={currentSheet.columns}
                rows={currentSheet.rows}
                rules={filterRules}
                onRulesChange={setFilterRules}
                onClearFilters={() => setFilterRules([])}
                onQuickSearch={setQuickSearch}
                quickSearchQuery={quickSearch}
              />
            </section>

            {/* Main Interactive Excel Grid */}
            <section id="module-table">
              <ExcelTable
                columns={currentSheet.columns}
                rows={filteredRows}
                sortConfig={sortConfig}
                onSortChange={setSortConfig}
              />
            </section>
          </>
        ) : (
          /* Empty / Instructions Panel */
          <div id="instructions-card" className="bg-slate-900 border border-slate-800 rounded-xl p-6 md:p-8 shadow-xl max-w-3xl mx-auto mt-6">
            <div className="flex gap-4 items-start">
              <div className="p-3 bg-emerald-500/15 text-emerald-400 border border-emerald-500/20 rounded-xl hidden sm:block">
                <HelpCircle className="w-5 h-5" />
              </div>
              <div className="space-y-4">
                <h2 className="text-base font-serif italic text-white">Petunjuk Penggunaan &amp; Fitur Sistem:</h2>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                  <div className="space-y-1.5 p-4 bg-slate-950/40 rounded-lg border border-slate-800">
                    <h4 className="font-semibold text-slate-200 flex items-center gap-1.5">
                      <ChevronRight className="w-3.5 h-3.5 text-emerald-400" />
                      1. Unggah Excel atau CSV
                    </h4>
                    <p className="text-slate-400 leading-relaxed pl-5">
                      Unggah berkas spreadsheet (.xlsx, .xls, .csv). Format kolom dipindai otomatis untuk mengidentifikasi tipe data (Teks, Angka, Tanggal, atau Boolean).
                    </p>
                  </div>

                  <div className="space-y-1.5 p-4 bg-slate-950/40 rounded-lg border border-slate-800">
                    <h4 className="font-semibold text-slate-200 flex items-center gap-1.5">
                      <ChevronRight className="w-3.5 h-3.5 text-emerald-400" />
                      2. Filter Header Responsif
                    </h4>
                    <p className="text-slate-400 leading-relaxed pl-5">
                      Tambahkan kriteria penyaringan berseri. Anda dapat memfilter angka (e.g. &gt; 5 juta), rujukan tanggal, kesamaan boolean, atau teks secara instan.
                    </p>
                  </div>

                  <div className="space-y-1.5 p-4 bg-slate-950/40 rounded-lg border border-slate-800">
                    <h4 className="font-semibold text-slate-200 flex items-center gap-1.5">
                      <ChevronRight className="w-3.5 h-3.5 text-emerald-400" />
                      3. Kalkulasi &amp; Urut Otomatis
                    </h4>
                    <p className="text-slate-400 leading-relaxed pl-5">
                      Aplikasi otomatis menampilkan indikator rangkuman statistik data (Total Sum, Rata-rata, Nilai Maks/Min) dinamis mengikuti hasil saringan baris aktif.
                    </p>
                  </div>

                  <div className="space-y-1.5 p-4 bg-slate-950/40 rounded-lg border border-slate-800">
                    <h4 className="font-semibold text-slate-200 flex items-center gap-1.5">
                      <ChevronRight className="w-3.5 h-3.5 text-emerald-400" />
                      4. Cetak Berkas Saringan
                    </h4>
                    <p className="text-slate-400 leading-relaxed pl-5">
                      Gunakan tombol <strong>Cetak Hasil Filter</strong> untuk melihat print-view bersih berprioritas cetak tinggi dengan semua ringkasan parameter dan data!
                    </p>
                  </div>
                </div>

                <div className="pt-2 text-[11px] text-slate-500 text-center italic border-t border-slate-800">
                  Diproses secara privat di peramban Anda dengan jaminan keamanan Sandbox internal.
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
