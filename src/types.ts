/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ExcelColumn {
  key: string;
  name: string;
  type: 'string' | 'number' | 'boolean' | 'date';
}

export type ExcelRow = Record<string, any>;

export interface WorksheetData {
  name: string;
  columns: ExcelColumn[];
  rows: ExcelRow[];
}

export interface ExcelWorkbookData {
  fileName: string;
  fileSize: string;
  sheets: WorksheetData[];
  activeSheetIndex: number;
}

export type FilterOperator = 'contains' | 'equals' | 'not_equals' | 'starts_with' | 'ends_with' | 'greater_than' | 'less_than' | 'empty' | 'not_empty' | 'one_of' | 'not_one_of';

export interface FilterRule {
  id: string;
  columnKey: string;
  operator: FilterOperator;
  value: string;
}

export interface SortConfig {
  columnKey: string;
  direction: 'asc' | 'desc';
}
