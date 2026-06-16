/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { ExcelColumn, ExcelRow, FilterRule } from '../types';

interface ExcelPrintTemplateProps {
  fileName: string;
  sheetName: string;
  columns: ExcelColumn[];
  rows: ExcelRow[];
  activeRules: FilterRule[];
  quickSearchQuery: string;
  paperSize: string;
}

export default function ExcelPrintTemplate({
  fileName,
  sheetName,
  columns,
  rows,
  activeRules,
  quickSearchQuery,
  paperSize,
}: ExcelPrintTemplateProps) {
  const currentDateStr = new Date().toLocaleString('id-ID', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  // Calculate sum of numeric columns to display in print totals row
  const columnSums = React.useMemo(() => {
    const sums: Record<string, number> = {};
    columns.forEach(col => {
      if (col.type === 'number') {
        let colSum = 0;
        let hasValue = false;
        rows.forEach(r => {
          const val = Number(r[col.key]);
          if (!isNaN(val) && r[col.key] !== '') {
            colSum += val;
            hasValue = true;
          }
        });
        if (hasValue) {
          sums[col.key] = colSum;
        }
      }
    });
    return sums;
  }, [columns, rows]);

  const formatPrintCellValue = (value: any, type: 'string' | 'number' | 'boolean' | 'date', columnKey: string): string => {
    if (value === null || value === undefined || value === '') return '—';
    if (type === 'boolean') {
      const isTrue = value === true || String(value).toUpperCase() === 'TRUE' || value === 1 || String(value) === '1' || String(value).toLowerCase() === 'ya';
      return isTrue ? 'YA' : 'TIDAK';
    }
    if (type === 'date') {
      try {
        const d = new Date(value);
        if (!isNaN(d.getTime())) {
          return d.toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' });
        }
      } catch (_) {}
    }
    if (type === 'number') {
      const num = Number(value);
      const isCurrency = columnKey.toLowerCase().includes('harga') || columnKey.toLowerCase().includes('total') || columnKey.toLowerCase().includes('pendapatan') || columnKey.toLowerCase().includes('omset') || columnKey.toLowerCase().includes('gaji') || columnKey.toLowerCase().includes('nominal') || columnKey.toLowerCase().includes('revenue') || columnKey.toLowerCase().includes('biaya');
      if (isCurrency && !isNaN(num)) {
        return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(num);
      }
      if (!isNaN(num)) {
        return new Intl.NumberFormat('id-ID').format(num);
      }
    }
    return String(value);
  };

  const pageStyle = React.useMemo(() => {
    let sizeDecl = paperSize;
    if (paperSize === 'F4') {
      sizeDecl = '215.9mm 330.2mm'; // Commonly used for F4 / Folio
    }
    return `
      @page {
        size: ${sizeDecl};
        margin: 1.5cm;
      }
    `;
  }, [paperSize]);

  return (
    <div className="hidden print:block p-8 bg-white text-slate-950 text-xs leading-relaxed font-sans w-full">
      <style dangerouslySetInnerHTML={{ __html: pageStyle }} />
      {/* Report Header */}
      <div className="border-b-2 border-slate-900 pb-4 mb-6 flex justify-between items-end">
        <div>
          <h1 className="text-xl font-bold uppercase tracking-tight text-slate-900">
            Laporan Analisis Data Excel
          </h1>
          <p className="text-slate-500 text-[10px] mt-1 font-mono">
            Sumber Berkas: {fileName} &bull; Lembar Kerja: {sheetName}
          </p>
        </div>
        <div className="text-right">
          <p className="text-[10px] text-slate-500">Dicetak Pada:</p>
          <p className="text-[10px] font-semibold text-slate-800">{currentDateStr}</p>
        </div>
      </div>

      {/* Applied Filters Metadata Box */}
      <div className="bg-slate-50 border border-slate-200 rounded p-3 mb-6">
        <h3 className="font-bold text-[10px] uppercase text-slate-600 tracking-wider mb-1.5">
          Parameter Filter Aktif
        </h3>
        <div className="grid grid-cols-2 gap-y-1 text-[10px] text-slate-700">
          <div>
            <span>Filter Kriteria: </span>
            <strong className="font-semibold">
              {activeRules.length === 0
                ? 'Semua data (Tanpa filter)'
                : `${activeRules.length} kriteria filter aktif`}
            </strong>
          </div>
          <div>
            <span>Pencarian Global: </span>
            <strong className="font-semibold">
              {quickSearchQuery ? `"${quickSearchQuery}"` : 'Tidak ada'}
            </strong>
          </div>
          {activeRules.length > 0 && (
            <div className="col-span-2 mt-1.5 pt-1.5 border-t border-slate-200">
              <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Rincian Aturan:</span>
              <ul className="list-disc list-inside space-y-0.5 pl-1">
                {activeRules.map((rule, idx) => {
                  const ruleColName = columns.find(c => c.key === rule.columnKey)?.name || rule.columnKey;
                  return (
                    <li key={idx} className="text-[9px] text-slate-600 font-mono">
                      {ruleColName} &rarr; {rule.operator.replace('_', ' ')} &rarr; {rule.value || '(Kosong)'}
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
        </div>
      </div>

      {/* Rows matched statistics */}
      <p className="text-[10px] font-bold text-slate-800 mb-3 bg-slate-100 px-2 py-1 inline-block rounded">
        Daftar Baris Hasil Filter ({rows.length} Baris Ditemukan)
      </p>

      {/* High-contrast printable grid table */}
      <table className="w-full text-left border-collapse border border-slate-400">
        <thead>
          <tr className="bg-slate-100 text-[10px] font-bold text-slate-900 border-b border-slate-400 font-mono">
            <th className="py-2 px-2 border-r border-slate-400 w-10 text-center">No</th>
            {columns.map(col => (
              <th key={col.key} className="py-2 px-2 border-r border-slate-400 last:border-0 uppercase tracking-wider">
                {col.name}
              </th>
            ))}
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-300 text-[9px]">
          {rows.length === 0 ? (
            <tr>
              <td colSpan={columns.length + 1} className="py-8 text-center text-slate-500 italic">
                Tidak ada baris yang cocok dengan filter aktif
              </td>
            </tr>
          ) : (
            rows.map((row, idx) => (
              <tr key={idx} className="page-break-inside-avoid">
                <td className="py-1.5 px-2 text-center border-r border-slate-300 font-mono font-medium">
                  {idx + 1}
                </td>
                {columns.map(col => (
                  <td key={col.key} className="py-1.5 px-2 border-r border-slate-300 last:border-0 max-w-[150px] truncate">
                    {formatPrintCellValue(row[col.key], col.type, col.key)}
                  </td>
                ))}
              </tr>
            ))
          )}

          {/* Sum footer row */}
          {rows.length > 0 && Object.keys(columnSums).length > 0 && (
            <tr className="bg-slate-50 border-t-2 border-slate-400 font-semibold font-mono text-[9px]">
              <td className="py-2 px-2 text-center border-r border-slate-400 text-slate-700">RINGKASAN</td>
              {columns.map(col => {
                const sum = columnSums[col.key];
                return (
                  <td key={col.key} className="py-2 px-2 border-r border-slate-400 last:border-0">
                    {sum !== undefined ? (
                      <div>
                        <span className="block text-[8px] text-slate-400 font-bold uppercase">SUM</span>
                        {formatPrintCellValue(sum, 'number', col.key)}
                      </div>
                    ) : (
                      '—'
                    )}
                  </td>
                );
              })}
              </tr>
            )}
        </tbody>
      </table>

      {/* Professional Footer Statement */}
      <div className="mt-12 text-center border-t border-slate-300 pt-4 page-break-inside-avoid">
        <p className="text-[9px] text-slate-400 uppercase tracking-widest font-mono">
          - Dokumen ini dibuat otomatis menggunakan aplikasi Pengelola File Excel -
        </p>
      </div>
    </div>
  );
}
