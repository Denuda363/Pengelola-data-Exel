/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { ChevronUp, ChevronDown, ChevronRight, ChevronsUpDown, Info, Percent, Sigma } from 'lucide-react';
import { ExcelColumn, ExcelRow, SortConfig } from '../types';

interface ExcelTableProps {
  columns: ExcelColumn[];
  rows: ExcelRow[];
  sortConfig: SortConfig | null;
  onSortChange: (config: SortConfig | null) => void;
  groupByColumn: string;
  setGroupByColumn: (col: string) => void;
}

// Format Helper
function formatCellValue(value: any, type: 'string' | 'number' | 'boolean' | 'date', columnKey: string): React.ReactNode {
  if (value === null || value === undefined || value === '') {
    return <span className="text-slate-600">—</span>;
  }

  if (type === 'boolean') {
    const isTrue = value === true || String(value).toUpperCase() === 'TRUE' || value === 1 || String(value) === '1' || String(value).toLowerCase() === 'ya';
    return (
      <span
        className={`inline-flex items-center px-2 py-0.5 rounded text-[10px] font-mono font-bold border ${
          isTrue ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20' : 'bg-rose-500/10 text-rose-400 border-rose-500/20'
        }`}
      >
        {isTrue ? 'YA' : 'TIDAK'}
      </span>
    );
  }

  if (type === 'date') {
    try {
      const d = new Date(value);
      if (!isNaN(d.getTime())) {
        return (
          <span className="font-mono text-xs text-slate-400">
            {d.toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' })}
          </span>
        );
      }
    } catch (_) {}
  }

  if (type === 'number') {
    const num = Number(value);
    
    // Check if it looks like Indonesian Rupiah / large financial data
    const keyLower = columnKey.toLowerCase();
    const isCurrency = keyLower.includes('harga') || keyLower.includes('total') || keyLower.includes('pendapatan') || keyLower.includes('omset') || keyLower.includes('gaji') || keyLower.includes('nominal') || keyLower.includes('revenue') || keyLower.includes('biaya') || keyLower.includes('price');
    
    if (isCurrency && !isNaN(num)) {
      return (
        <span className="font-mono font-medium text-white">
          {new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(num)}
        </span>
      );
    }
    
    if (!isNaN(num)) {
      return <span className="font-mono text-slate-300">{new Intl.NumberFormat('id-ID').format(num)}</span>;
    }
  }

  // Format special status strings nicely
  const strVal = String(value);
  if (strVal === 'Lunas') {
    return <span className="inline-flex items-center px-2.5 py-0.5 rounded text-[10px] font-mono font-semibold bg-emerald-500/15 text-emerald-400 border border-emerald-500/25">Lunas</span>;
  }
  if (strVal === 'Menunggu Pembayaran' || strVal === 'Pending') {
    return <span className="inline-flex items-center px-2.5 py-0.5 rounded text-[10px] font-mono font-semibold bg-amber-500/15 text-amber-400 border border-amber-500/25">Menunggu</span>;
  }
  if (strVal === 'Dicicil') {
    return <span className="inline-flex items-center px-2.5 py-0.5 rounded text-[10px] font-mono font-semibold bg-blue-500/15 text-blue-400 border border-blue-500/25">Dicicil</span>;
  }
  if (strVal === 'Dibatalkan' || strVal === 'Gagal') {
    return <span className="inline-flex items-center px-2.5 py-0.5 rounded text-[10px] font-mono font-semibold bg-rose-500/15 text-rose-400 border border-rose-500/25">Dibatalkan</span>;
  }

  return <span className="text-slate-300 font-medium">{strVal}</span>;
}

export default function ExcelTable({ columns, rows, sortConfig, onSortChange, groupByColumn, setGroupByColumn }: ExcelTableProps) {
  const [currentPage, setCurrentPage] = useState(1);
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [selectedSumCol, setSelectedSumCol] = useState<string>('');
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({});

  // Reset collapsed state when grouping changes
  React.useEffect(() => {
    setCollapsedGroups({});
  }, [groupByColumn]);

  // Reset pagination when rows length changes
  React.useEffect(() => {
    setCurrentPage(1);
  }, [rows.length]);

  // Handle Sort Toggle
  const handleSort = (columnKey: string) => {
    if (sortConfig?.columnKey === columnKey) {
      if (sortConfig.direction === 'asc') {
        onSortChange({ columnKey, direction: 'desc' });
      } else {
        onSortChange(null); // Clear sort
      }
    } else {
      onSortChange({ columnKey, direction: 'asc' });
    }
  };

  // Find numeric columns for calculation
  const numericColumns = useMemo(() => {
    return columns.filter(c => c.type === 'number');
  }, [columns]);

  // Set default sum column on load
  React.useEffect(() => {
    if (numericColumns.length > 0 && !selectedSumCol) {
      // Prefer totals or income
      const bestChoice = numericColumns.find(c => 
        c.key.toLowerCase().includes('total') || 
        c.key.toLowerCase().includes('pendapatan') || 
        c.key.toLowerCase().includes('harga')
      ) || numericColumns[0];
      setSelectedSumCol(bestChoice.key);
    }
  }, [numericColumns, selectedSumCol]);

  // Calculate statistics for the SELECTED column
  const stats = useMemo(() => {
    if (!selectedSumCol || rows.length === 0) return null;
    
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    let validCount = 0;

    rows.forEach(r => {
      const val = Number(r[selectedSumCol]);
      if (!isNaN(val) && r[selectedSumCol] !== '') {
        sum += val;
        if (val < min) min = val;
        if (val > max) max = val;
        validCount++;
      }
    });

    const average = validCount > 0 ? sum / validCount : 0;

    return {
      sum: validCount > 0 ? sum : 0,
      average: validCount > 0 ? average : 0,
      min: validCount > 0 ? min : 0,
      max: validCount > 0 ? max : 0,
      count: validCount,
    };
  }, [selectedSumCol, rows]);

  // Calculate Paginated Rows
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return rows.slice(start, start + rowsPerPage);
  }, [rows, currentPage, rowsPerPage]);

  const totalPages = Math.ceil(rows.length / rowsPerPage);

  // Group rows if group-by column is selected
  const groupedData = useMemo(() => {
    if (!groupByColumn) return null;
    
    const groups: Record<string, ExcelRow[]> = {};
    rows.forEach(row => {
      const rawVal = row[groupByColumn];
      const groupVal = (rawVal === null || rawVal === undefined || rawVal === '') ? 'Tidak Ada Nilai' : String(rawVal);
      if (!groups[groupVal]) {
        groups[groupVal] = [];
      }
      groups[groupVal].push(row);
    });
    
    return Object.entries(groups).map(([groupKey, groupRows]) => {
      let groupSum = 0;
      let hasNumeric = false;
      if (selectedSumCol) {
        groupRows.forEach(r => {
          const val = Number(r[selectedSumCol]);
          if (!isNaN(val) && r[selectedSumCol] !== '') {
            groupSum += val;
            hasNumeric = true;
          }
        });
      }
      
      return {
        groupValue: groupKey,
        rows: groupRows,
        subtotal: hasNumeric ? groupSum : null,
      };
    });
  }, [rows, groupByColumn, selectedSumCol]);

  const toggleGroup = (groupKey: string) => {
    setCollapsedGroups(prev => ({
      ...prev,
      [groupKey]: !prev[groupKey]
    }));
  };

  const expandAllGroups = () => {
    setCollapsedGroups({});
  };

  const collapseAllGroups = () => {
    if (!groupedData) return;
    const collapsed: Record<string, boolean> = {};
    groupedData.forEach(g => {
      collapsed[g.groupValue] = true;
    });
    setCollapsedGroups(collapsed);
  };

  const formatSummaryValue = (val: number, columnKey: string) => {
    const keyLower = columnKey.toLowerCase();
    const isCurrency = keyLower.includes('harga') || keyLower.includes('total') || keyLower.includes('pendapatan') || keyLower.includes('omset') || keyLower.includes('gaji') || keyLower.includes('nominal') || keyLower.includes('revenue') || keyLower.includes('biaya');
    if (isCurrency) {
      return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(val);
    }
    return new Intl.NumberFormat('id-ID').format(val);
  };

  return (
    <div className="space-y-4">
      {/* Table Statistics Summary Card */}
      {stats && (
        <div id="quick-stats-panel" className="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 shadow-lg">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-lg">
              <Sigma className="w-5 h-5 animate-pulse" />
            </div>
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs font-semibold text-white uppercase tracking-wider font-sans">Kalkulator Rangkuman</span>
                <select
                  id="select-kalkulasi-kolom"
                  value={selectedSumCol}
                  onChange={(e) => setSelectedSumCol(e.target.value)}
                  className="bg-slate-950 border border-slate-800 text-slate-300 font-semibold text-xs py-0.5 px-2 rounded focus:outline-none focus:ring-1 focus:ring-emerald-500 outline-none cursor-pointer"
                >
                  {numericColumns.map(nc => (
                    <option key={nc.key} value={nc.key}>{nc.name}</option>
                  ))}
                </select>
              </div>
              <p className="text-[11px] text-slate-500 mt-0.5">Statistik dinamis untuk baris data yang lolos filter aktif</p>
            </div>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-x-6 gap-y-2 text-left bg-slate-950 px-4 py-2.5 rounded border border-slate-800/85 flex-1 max-w-2xl">
            <div>
              <span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider font-mono">Total Jumlah</span>
              <span id="stat-sum-val" className="text-xs md:text-sm font-bold text-white font-mono">
                {formatSummaryValue(stats.sum, selectedSumCol)}
              </span>
            </div>
            <div>
              <span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider font-mono">Rata-rata</span>
              <span id="stat-avg-val" className="text-xs md:text-sm font-semibold text-emerald-400 font-mono">
                {formatSummaryValue(stats.average, selectedSumCol)}
              </span>
            </div>
            <div>
              <span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider font-mono">Tertinggi</span>
              <span id="stat-max-val" className="text-xs md:text-sm font-bold text-emerald-400 font-mono">
                {formatSummaryValue(stats.max, selectedSumCol)}
              </span>
            </div>
            <div>
              <span className="block text-[10px] text-slate-500 font-bold uppercase tracking-wider font-mono">Terendah</span>
              <span id="stat-min-val" className="text-xs md:text-sm font-bold text-amber-500 font-mono">
                {formatSummaryValue(stats.min, selectedSumCol)}
              </span>
            </div>
          </div>
        </div>
      )}

      {/* Main Table Container */}
      <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-lg">
        {groupByColumn && groupedData && (
          <div className="px-4 py-3 bg-slate-950/60 border-b border-slate-800/85 flex flex-col sm:flex-row gap-3 justify-between items-center text-xs text-slate-300">
            <span className="font-semibold text-slate-400">
              Pengelompokan aktif berdasarkan <span className="text-emerald-400">"{columns.find(c => c.key === groupByColumn)?.name}"</span> &bull; Terbentuk <span className="text-emerald-400 font-mono font-bold">{groupedData.length}</span> Kelompok
            </span>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={expandAllGroups}
                className="text-[10px] bg-slate-800 hover:bg-slate-750 text-slate-200 px-2.5 py-1 rounded transition-colors font-bold uppercase tracking-wider cursor-pointer"
              >
                Buka Semua
              </button>
              <button
                type="button"
                onClick={collapseAllGroups}
                className="text-[10px] bg-slate-800 hover:bg-slate-750 text-slate-200 px-2.5 py-1 rounded transition-colors font-bold uppercase tracking-wider cursor-pointer"
              >
                Tutup Semua
              </button>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table id="excel-data-table" className="w-full text-left border-collapse min-w-[700px]">
            <thead>
              <tr className="bg-slate-950 border-b border-slate-800 text-xs font-semibold text-slate-400">
                <th className="py-3 px-3 w-12 text-center text-slate-500 border-r border-slate-800">#</th>
                {columns.map((col) => {
                  const isSorted = sortConfig?.columnKey === col.key;
                  return (
                    <th
                      key={col.key}
                      onClick={() => handleSort(col.key)}
                      className={`py-3 px-4 font-semibold select-none cursor-pointer transition-colors border-r border-slate-800/70 last:border-0 ${
                        isSorted ? 'bg-emerald-500/5 text-emerald-400' : 'hover:bg-slate-800 text-slate-300'
                      }`}
                    >
                      <div className="flex items-center gap-1.5">
                        <span className="line-clamp-1">{col.name}</span>
                        {isSorted ? (
                          sortConfig.direction === 'asc' ? (
                            <ChevronUp className="w-3.5 h-3.5 text-emerald-400 flex-shrink-0" />
                          ) : (
                            <ChevronDown className="w-3.5 h-3.5 text-emerald-400 flex-shrink-0" />
                          )
                        ) : (
                          <ChevronsUpDown className="w-3 h-3 text-slate-600 hover:text-slate-400 flex-shrink-0" />
                        )}
                      </div>
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800 text-xs">
              {groupByColumn ? (
                groupedData && groupedData.length === 0 ? (
                  <tr>
                    <td colSpan={columns.length + 1} className="py-12 text-center text-slate-500 italic">
                      Tidak ada baris yang cocok dengan filter aktif
                    </td>
                  </tr>
                ) : groupedData ? (
                  groupedData.map(({ groupValue, rows: groupRows, subtotal }) => {
                    const isCollapsed = collapsedGroups[groupValue];
                    return (
                      <React.Fragment key={groupValue}>
                        {/* Group Header Row */}
                        <tr
                          onClick={() => toggleGroup(groupValue)}
                          className="bg-slate-800/40 hover:bg-slate-800/70 transition-colors cursor-pointer border-y border-slate-800 select-none font-bold"
                        >
                          <td className="py-3 px-3 text-center border-r border-slate-800">
                            {isCollapsed ? (
                              <ChevronRight className="w-3.5 h-3.5 mx-auto text-emerald-400" />
                            ) : (
                              <ChevronDown className="w-3.5 h-3.5 mx-auto text-emerald-400" />
                            )}
                          </td>
                          <td colSpan={columns.length} className="py-3 px-4 text-xs text-slate-200">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                              <div className="flex items-center gap-2">
                                <span className="text-slate-500 text-[10px] font-bold uppercase tracking-wider">Kelompok:</span>
                                <span className="text-emerald-400 font-extrabold">{groupValue}</span>
                                <span className="bg-slate-950 px-2 py-0.5 rounded text-[10px] text-slate-400 border border-slate-850 font-normal">
                                  {groupRows.length} Baris
                                </span>
                              </div>
                              {subtotal !== null && (
                                <div className="flex items-center gap-1.5 bg-emerald-500/5 border border-emerald-500/10 px-2.5 py-1 rounded text-[10px]">
                                  <span className="text-slate-500 font-bold uppercase text-[9px]">Subtotal:</span>
                                  <span className="text-emerald-400 font-bold font-mono">
                                    {formatSummaryValue(subtotal, selectedSumCol)}
                                  </span>
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>

                        {/* Group Rows (only if not collapsed) */}
                        {!isCollapsed && groupRows.map((row, idx) => {
                          const globalIndex = rows.indexOf(row) + 1;
                          return (
                            <tr
                              key={`${groupValue}-row-${idx}`}
                              className="bg-slate-900/30 hover:bg-slate-850/50 transition-colors"
                            >
                              <td className="py-2.5 px-3 text-center text-slate-500 border-r border-slate-800 bg-slate-950/10 font-mono text-[11px] font-medium border-l-[3px] border-l-emerald-500/30">
                                {globalIndex}
                              </td>
                              {columns.map((col) => (
                                <td
                                  key={col.key}
                                  className="py-2.5 px-4 text-slate-300 border-r border-slate-800/60 last:border-0"
                                >
                                  {formatCellValue(row[col.key], col.type, col.key)}
                                </td>
                              ))}
                            </tr>
                          );
                        })}
                      </React.Fragment>
                    );
                  })
                ) : null
              ) : paginatedRows.length === 0 ? (
                <tr>
                  <td colSpan={columns.length + 1} className="py-12 text-center text-slate-500 italic">
                    {rows.length === 0 ? 'Tidak ada data file excel yang tersedia' : 'Tidak ada baris yang cocok dengan filter aktif'}
                  </td>
                </tr>
              ) : (
                paginatedRows.map((row, idx) => {
                  const absoluteIndex = (currentPage - 1) * rowsPerPage + idx + 1;
                  return (
                    <tr
                      key={idx}
                      className="hover:bg-slate-850/50 transition-colors"
                    >
                      <td className="py-2.5 px-3 text-center text-slate-500 border-r border-slate-800 bg-slate-950/40 font-mono text-[11px] font-medium">
                        {absoluteIndex}
                      </td>
                      {columns.map((col) => (
                        <td
                          key={col.key}
                          className="py-2.5 px-4 text-slate-300 border-r border-slate-800/60 last:border-0"
                        >
                          {formatCellValue(row[col.key], col.type, col.key)}
                        </td>
                      ))}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Dynamic Pagination Controls */}
        {!groupByColumn && totalPages > 0 && (
          <div className="bg-slate-950 px-4 py-3.5 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-4 text-slate-400">
            <div className="flex items-center gap-4 text-xs">
              <span className="font-medium text-slate-400">
                Menampilkan <strong className="font-semibold text-slate-200 font-mono text-xs">{Math.min(rows.length, (currentPage - 1) * rowsPerPage + 1)}-{Math.min(rows.length, currentPage * rowsPerPage)}</strong> dari <strong className="font-bold text-emerald-400 font-mono text-xs">{rows.length}</strong> baris data
              </span>

              <div className="flex items-center gap-1.5 border-l border-slate-800 pl-4">
                <span className="text-slate-500 text-[11px]">Baris:</span>
                <select
                  id="select-rows-per-page"
                  value={rowsPerPage}
                  onChange={(e) => {
                    setRowsPerPage(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="bg-slate-900 border border-slate-800 text-slate-300 text-[11px] font-semibold py-0.5 px-1.5 rounded cursor-pointer outline-none focus:border-emerald-500"
                >
                  <option value={10}>10</option>
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                </select>
              </div>
            </div>

            <div className="flex items-center gap-1">
              <button
                id="btn-pagination-first"
                onClick={() => setCurrentPage(1)}
                disabled={currentPage === 1}
                className="px-2.5 py-1 text-xs border border-slate-850 bg-slate-900 hover:bg-slate-800 hover:text-white text-slate-300 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-all font-semibold font-sans cursor-pointer"
              >
                Pertama
              </button>
              <button
                id="btn-pagination-prev"
                onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                disabled={currentPage === 1}
                className="px-2.5 py-1 text-xs border border-slate-850 bg-slate-900 hover:bg-slate-800 hover:text-white text-slate-300 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-all font-semibold font-sans cursor-pointer"
              >
                Sebelumnya
              </button>

              <span className="px-3.5 text-xs font-semibold text-slate-450 font-mono">
                {currentPage} / {totalPages}
              </span>

              <button
                id="btn-pagination-next"
                onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                disabled={currentPage === totalPages}
                className="px-2.5 py-1 text-xs border border-slate-850 bg-slate-900 hover:bg-slate-800 hover:text-white text-slate-300 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-all font-semibold font-sans cursor-pointer"
              >
                Selanjutnya
              </button>
              <button
                id="btn-pagination-last"
                onClick={() => setCurrentPage(totalPages)}
                disabled={currentPage === totalPages}
                className="px-2.5 py-1 text-xs border border-slate-850 bg-slate-900 hover:bg-slate-800 hover:text-white text-slate-300 rounded disabled:opacity-30 disabled:cursor-not-allowed transition-all font-semibold font-sans cursor-pointer"
              >
                Terakhir
              </button>
            </div>
          </div>
        )}

        {groupByColumn && (
          <div className="bg-slate-950 px-4 py-3.5 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-3 text-slate-400 text-xs mt-auto">
            <span>
              Menampilkan <strong className="text-emerald-400 font-mono text-xs">{groupedData?.length || 0}</strong> kelompok dengan total <strong className="font-bold text-emerald-400 font-mono text-xs">{rows.length}</strong> baris data.
            </span>
            <span className="text-[11px] text-slate-500 italic">
              Klik pada baris kelompok untuk melipat/membuka isi baris data.
            </span>
          </div>
        )}
      </div>
    </div>
  );
}
