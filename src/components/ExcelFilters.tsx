/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { Plus, Trash2, Filter, RotateCcw, ChevronDown } from 'lucide-react';
import { ExcelColumn, FilterRule, FilterOperator, ExcelRow } from '../types';

interface ExcelFiltersProps {
  columns: ExcelColumn[];
  rows: ExcelRow[];
  rules: FilterRule[];
  onRulesChange: (rules: FilterRule[]) => void;
  onClearFilters: () => void;
  onQuickSearch: (query: string) => void;
  quickSearchQuery: string;
}

const OPERATORS_MAP: { value: FilterOperator; label: string; types: string[] }[] = [
  { value: 'one_of', label: 'Pilih dari data unik (Pilihan)', types: ['string', 'number'] },
  { value: 'not_one_of', label: 'Kecualikan dari data unik', types: ['string', 'number'] },
  { value: 'contains', label: 'Mengandung kata', types: ['string'] },
  { value: 'equals', label: 'Sama dengan (=)', types: ['string', 'number', 'boolean', 'date'] },
  { value: 'not_equals', label: 'Tidak sama dengan (≠)', types: ['string', 'number', 'boolean', 'date'] },
  { value: 'starts_with', label: 'Diawali dengan', types: ['string'] },
  { value: 'ends_with', label: 'Diakhiri dengan', types: ['string'] },
  { value: 'greater_than', label: 'Lebih besar dari (>)', types: ['number', 'date'] },
  { value: 'less_than', label: 'Lebih kecil dari (<)', types: ['number', 'date'] },
  { value: 'empty', label: 'Kosong (Blangko)', types: ['string', 'number', 'boolean', 'date'] },
  { value: 'not_empty', label: 'Tidak Kosong', types: ['string', 'number', 'boolean', 'date'] },
];

interface UniqueValuesChecklistProps {
  rows: ExcelRow[];
  columnKey: string;
  value: string;
  onChange: (newValue: string) => void;
}

function UniqueValuesChecklist({ rows, columnKey, value, onChange }: UniqueValuesChecklistProps) {
  const [searchTerm, setSearchTerm] = React.useState('');

  const uniqueVals = React.useMemo(() => {
    const counts: Record<string, number> = {};
    rows.forEach(r => {
      const val = r[columnKey];
      const key = (val === null || val === undefined || val === '') ? '—' : String(val).trim();
      counts[key] = (counts[key] || 0) + 1;
    });
    
    return Object.entries(counts)
      .map(([val, count]) => ({ value: val, count }))
      .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value, undefined, { numeric: true }));
  }, [rows, columnKey]);

  const selectedValues = React.useMemo(() => {
    try {
      return (JSON.parse(value || '[]') as string[]);
    } catch (e) {
      return [];
    }
  }, [value]);

  const handleToggle = (val: string) => {
    const next = selectedValues.includes(val)
      ? selectedValues.filter(v => v !== val)
      : [...selectedValues, val];
    onChange(JSON.stringify(next));
  };

  const handleSelectAll = () => {
    onChange(JSON.stringify(uniqueVals.map(u => u.value)));
  };

  const handleClearAll = () => {
    onChange(JSON.stringify([]));
  };

  const filtered = uniqueVals.filter(uv => 
    uv.value.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="mt-2.5 p-3 bg-slate-950 border border-slate-800/80 rounded space-y-2.5 text-xs w-full animate-fadeIn">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800/60 pb-2">
        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider font-mono">
          Daftar Nilai Unik Terdeteksi ({selectedValues.length} dipilih dari {uniqueVals.length})
        </span>
        <div className="flex gap-2.5 items-center">
          <button
            type="button"
            onClick={handleSelectAll}
            className="text-[10px] text-emerald-400 hover:text-emerald-300 font-bold uppercase tracking-wider p-0 bg-transparent cursor-pointer transition-colors"
          >
            Pilih Semua
          </button>
          <span className="text-slate-800">|</span>
          <button
            type="button"
            onClick={handleClearAll}
            className="text-[10px] text-red-400 hover:text-red-300 font-bold uppercase tracking-wider p-0 bg-transparent cursor-pointer transition-colors"
          >
            Hapus Semua
          </button>
        </div>
      </div>

      <input
        type="text"
        placeholder="Cari kata kunci dalam nilai unik ini..."
        value={searchTerm}
        onChange={(e) => setSearchTerm(e.target.value)}
        className="w-full px-2.5 py-1.5 bg-slate-900 border border-slate-800 rounded text-xs text-slate-200 placeholder-slate-600 focus:outline-none focus:border-emerald-500/50 transition-colors"
      />

      <div className="max-h-36 overflow-y-auto space-y-1.5 pr-1 scrollbar-thin text-slate-300">
        {filtered.length === 0 ? (
          <div className="text-[11px] text-slate-500 italic py-3 text-center">
            Tidak ada nilai data yang cocok dengan pencarian kata kunci.
          </div>
        ) : (
          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-x-4 gap-y-1">
            {filtered.map(({ value: val, count }) => {
              const isChecked = selectedValues.includes(val);
              return (
                <label
                  key={val}
                  className={`flex items-center gap-2 px-2 py-1.5 rounded hover:bg-slate-900/60 transition-colors cursor-pointer select-none border ${
                    isChecked ? 'bg-emerald-950/10 border-emerald-900/35 text-emerald-300' : 'border-transparent text-slate-400'
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={isChecked}
                    onChange={() => handleToggle(val)}
                    className="rounded border-slate-800 text-emerald-600 bg-slate-900 focus:ring-emerald-500/30 focus:ring-opacity-50 h-3.5 w-3.5 text-emerald-500 cursor-pointer"
                  />
                  <span className="flex-1 text-[11px] truncate leading-none">
                    {val}
                  </span>
                  <span className="font-mono text-[9px] px-1.5 py-0.2 bg-slate-900 text-slate-500 rounded border border-slate-800/40 font-bold shrink-0">
                    {count}x
                  </span>
                </label>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

export default function ExcelFilters({
  columns,
  rows,
  rules,
  onRulesChange,
  onClearFilters,
  onQuickSearch,
  quickSearchQuery,
}: ExcelFiltersProps) {

  const addRule = () => {
    if (columns.length === 0) return;
    const defaultColumn = columns[0];
    const operatorsForType = OPERATORS_MAP.filter(op => op.types.includes(defaultColumn.type));
    // Default to 'one_of' if it exists to showcase the awesome feature!
    const defaultOperator = operatorsForType.some(op => op.value === 'one_of') ? 'one_of' : (operatorsForType[0]?.value || 'contains');

    const newRule: FilterRule = {
      id: crypto.randomUUID(),
      columnKey: defaultColumn.key,
      operator: defaultOperator as FilterOperator,
      value: '[]',
    };
    onRulesChange([...rules, newRule]);
  };

  const removeRule = (id: string) => {
    onRulesChange(rules.filter(r => r.id !== id));
  };

  const updateRule = (id: string, updates: Partial<FilterRule>) => {
    onRulesChange(
      rules.map(r => {
        if (r.id !== id) return r;
        const updated = { ...r, ...updates };

        // Handle operator change initialization
        if (updates.operator) {
          if (updates.operator === 'one_of' || updates.operator === 'not_one_of') {
            updated.value = '[]';
          } else if (updates.operator === 'empty' || updates.operator === 'not_empty') {
            updated.value = '';
          } else {
            // Pick a default if changing away
            const targetCol = columns.find(c => c.key === updated.columnKey);
            updated.value = targetCol?.type === 'boolean' ? 'true' : '';
          }
        }

        // If column changed, reset operator and value if current isn't supported by the new column's type
        if (updates.columnKey) {
          const targetCol = columns.find(c => c.key === updates.columnKey);
          if (targetCol) {
            const allowedOps = OPERATORS_MAP.filter(op => op.types.includes(targetCol.type)).map(o => o.value);
            if (!allowedOps.includes(updated.operator)) {
              updated.operator = (allowedOps.includes('one_of') ? 'one_of' : allowedOps[0]) as FilterOperator;
            }
            
            if (updated.operator === 'one_of' || updated.operator === 'not_one_of') {
              updated.value = '[]';
            } else if (targetCol.type === 'boolean') {
              updated.value = 'true';
            } else if (updated.operator === 'empty' || updated.operator === 'not_empty') {
              updated.value = '';
            } else {
              updated.value = '';
            }
          }
        }

        return updated;
      })
    );
  };

  return (
    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 md:p-6 shadow-lg">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-slate-800 pb-4 mb-4">
        <div className="flex items-center gap-3">
          <div className="p-2 bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 rounded-lg">
            <Filter className="w-5 h-5" />
          </div>
          <div>
            <h4 className="text-sm font-semibold text-white font-sans">Filter Kolom Excel</h4>
            <p className="text-xs text-slate-400">Saring baris data berdasarkan kriteria nilai kolom header</p>
          </div>
        </div>

        {/* Global/Quick Search */}
        <div className="w-full md:w-72 flex gap-2">
          <input
            id="input-pencarian-cepat"
            type="text"
            value={quickSearchQuery}
            onChange={(e) => onQuickSearch(e.target.value)}
            placeholder="Cari kata kunci global..."
            className="w-full px-3 py-1.5 bg-slate-950 focus:bg-slate-950 text-xs text-slate-200 placeholder-slate-500 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 rounded outline-none transition-all"
          />
          {quickSearchQuery && (
            <button
              onClick={() => onQuickSearch('')}
              className="px-2.5 bg-slate-800 hover:bg-slate-700 text-xs text-slate-300 border border-slate-700 rounded font-medium transition-all cursor-pointer"
            >
              Reset
            </button>
          )}
        </div>
      </div>

      {/* Rules Builder */}
      {rules.length === 0 ? (
        <div className="flex flex-col items-center justify-center py-6 text-slate-500">
          <p className="text-xs text-center max-w-sm mb-3">
            Belum ada filter aktif. Seluruh data spreadsheet ditampilkan. Tambahkan aturan filter untuk menyaring baris.
          </p>
          <button
            id="btn-tambah-filter-awal"
            type="button"
            onClick={addRule}
            className="px-3.5 py-1.5 border border-dashed border-emerald-500/30 hover:border-emerald-400 text-emerald-400 hover:text-emerald-300 font-semibold text-xs rounded transition-all flex items-center gap-1.5 bg-emerald-500/5 cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            Tambah Filter Kolom
          </button>
        </div>
      ) : (
        <div className="space-y-3">
          <div className="max-h-[380px] overflow-y-auto space-y-3 pr-1 scrollbar-thin">
            {rules.map((rule, idx) => {
              const ruleCol = columns.find(c => c.key === rule.columnKey);
              const allowedOperators = OPERATORS_MAP.filter(op => op.types.includes(ruleCol?.type || 'string'));
              const isOneOfOperator = rule.operator === 'one_of' || rule.operator === 'not_one_of';

              return (
                <div
                  key={rule.id}
                  id={`filter-rule-${idx}`}
                  className="p-3 bg-slate-950 border border-slate-800/80 rounded flex flex-col gap-2.5 transition-all"
                >
                  <div className="flex flex-wrap items-center gap-2.5 w-full">
                    {/* Selector Header Kolom */}
                    <div className="flex-1 min-w-[150px]">
                      <span className="block text-[10px] font-bold text-slate-500 mb-1 tracking-wider uppercase font-mono">PILIH HEADER</span>
                      <select
                        id={`pilih-kolom-${idx}`}
                        value={rule.columnKey}
                        onChange={(e) => updateRule(rule.id, { columnKey: e.target.value })}
                        className="w-full text-xs text-slate-300 bg-slate-900 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 p-2 rounded outline-none cursor-pointer"
                      >
                        {columns.map(col => (
                          <option key={col.key} value={col.key}>
                            {col.name} ({col.type === 'number' ? 'Angka' : col.type === 'date' ? 'Tanggal' : col.type === 'boolean' ? 'Boolean' : 'Teks'})
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* Operator */}
                    <div className="flex-1 min-w-[135px]">
                      <span className="block text-[10px] font-bold text-slate-500 mb-1 tracking-wider uppercase font-mono">OPERATOR</span>
                      <select
                        id={`pilih-operator-${idx}`}
                        value={rule.operator}
                        onChange={(e) => updateRule(rule.id, { operator: e.target.value as FilterOperator })}
                        className="w-full text-xs text-slate-300 bg-slate-900 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 p-2 rounded outline-none cursor-pointer"
                      >
                        {allowedOperators.map(op => (
                          <option key={op.value} value={op.value}>
                            {op.label}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* Input Perbandingan Slider/Text/Date/Check */}
                    {rule.operator !== 'empty' && rule.operator !== 'not_empty' && (
                      <div className="flex-[1.5] min-w-[180px]">
                        <span className="block text-[10px] font-bold text-slate-500 mb-1 tracking-wider uppercase font-mono">NILAI ACUAN</span>
                        {isOneOfOperator ? (
                          <div className="text-xs text-emerald-400 font-semibold italic bg-emerald-500/5 px-2.5 py-2 border border-emerald-500/10 rounded flex items-center justify-between">
                            <span>Centang data pilihan di bawah</span>
                            <ChevronDown className="w-3 h-3 text-emerald-400" />
                          </div>
                        ) : ruleCol?.type === 'boolean' ? (
                          <select
                            id={`pilih-boolean-val-${idx}`}
                            value={rule.value}
                            onChange={(e) => updateRule(rule.id, { value: e.target.value })}
                            className="w-full text-xs text-slate-300 bg-slate-900 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 p-2 rounded outline-none cursor-pointer"
                          >
                            <option value="true">Benar / TRUE (Ya)</option>
                            <option value="false">Salah / FALSE (Tidak)</option>
                          </select>
                        ) : ruleCol?.type === 'date' ? (
                          <input
                            id={`input-date-val-${idx}`}
                            type="date"
                            value={rule.value}
                            onChange={(e) => updateRule(rule.id, { value: e.target.value })}
                            className="w-full text-xs text-slate-300 bg-slate-900 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 p-2 rounded outline-none"
                          />
                        ) : (
                          <>
                            <input
                              id={`input-text-val-${idx}`}
                              list={`datalist-col-${rule.id}`}
                              type={ruleCol?.type === 'number' ? 'number' : 'text'}
                              value={rule.value}
                              onChange={(e) => updateRule(rule.id, { value: e.target.value })}
                              placeholder={ruleCol?.type === 'number' ? 'Keberadaan angka...' : 'Ketik kecocokan teks...'}
                              className="w-full text-xs text-slate-300 bg-slate-900 border border-slate-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500 p-2 rounded outline-none"
                            />
                            <datalist id={`datalist-col-${rule.id}`}>
                              {React.useMemo(() => {
                                if (!rule.columnKey) return [];
                                const uniques = new Set<string>();
                                rows.forEach(r => {
                                  const val = r[rule.columnKey];
                                  if (val !== null && val !== undefined && val !== '') {
                                    uniques.add(String(val).trim());
                                  }
                                });
                                return Array.from(uniques).slice(0, 50).sort();
                              }, [rows, rule.columnKey]).map(val => (
                                <option key={val} value={val} />
                              ))}
                            </datalist>
                          </>
                        )}
                      </div>
                    )}

                    {/* Action delete rule */}
                    <div className="self-end pb-1 md:pb-0 shrink-0">
                      <button
                        id={`btn-hapus-rule-${idx}`}
                        type="button"
                        onClick={() => removeRule(rule.id)}
                        className="p-2 text-slate-500 hover:text-red-400 hover:bg-red-500/10 border border-transparent hover:border-red-500/20 rounded transition-all cursor-pointer"
                        title="Hapus filter ini"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>

                  {/* Checklist Drawer below if one_of operator is selected */}
                  {isOneOfOperator && (
                    <UniqueValuesChecklist
                      rows={rows}
                      columnKey={rule.columnKey}
                      value={rule.value}
                      onChange={(newValue) => updateRule(rule.id, { value: newValue })}
                    />
                  )}
                </div>
              );
            })}
          </div>

          <div className="flex items-center justify-between pt-3 border-t border-slate-800 mt-2">
            <button
              id="btn-tambah-filter-lanjutan"
              type="button"
              onClick={addRule}
              className="px-3 py-1.5 border border-slate-700 hover:border-slate-600 bg-slate-800 hover:bg-slate-700 text-slate-200 font-semibold text-xs rounded transition-all flex items-center gap-1 cursor-pointer hover:text-white"
            >
              <Plus className="w-3.5 h-3.5 text-emerald-400" />
              Tambah Aturan Baru
            </button>

            <button
              id="btn-reset-semua-filter"
              type="button"
              onClick={onClearFilters}
              className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 hover:text-white text-slate-300 border border-slate-700 font-semibold text-xs rounded transition-all flex items-center gap-1 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5 text-slate-400" />
              Reset Semua Filter
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
