/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import * as XLSX from 'xlsx';
import { ExcelColumn, ExcelRow, ExcelWorkbookData, WorksheetData } from '../types';

/**
 * Format files size into human readable string
 */
export function formatBytes(bytes: number, decimals = 2): string {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

/**
 * Detect the type of a column based on the values in its rows
 */
function detectColumnType(values: any[]): 'string' | 'number' | 'boolean' | 'date' {
  const nonNullValues = values.filter(v => v !== null && v !== undefined && v !== '');
  if (nonNullValues.length === 0) return 'string';

  // Check if they are all booleans
  const isBoolean = nonNullValues.every(v => typeof v === 'boolean' || v === 'true' || v === 'false' || v === 'TRUE' || v === 'FALSE');
  if (isBoolean) return 'boolean';

  // Check if they are all numbers or numeric strings
  const isNumber = nonNullValues.every(v => {
    if (typeof v === 'number') return true;
    return !isNaN(Number(v)) && !isNaN(parseFloat(v));
  });
  if (isNumber) return 'number';

  // Check if they look like dates
  const isDate = nonNullValues.every(v => {
    if (v instanceof Date) return true;
    if (typeof v === 'string') {
      const timestamp = Date.parse(v);
      return !isNaN(timestamp) && v.includes('-') || v.includes('/');
    }
    return false;
  });
  if (isDate) return 'date';

  return 'string';
}

/**
 * Parse an Excel file element in arrayBuffer format
 */
export function parseExcelFile(arrayBuffer: ArrayBuffer, fileName: string, fileSize: number): ExcelWorkbookData {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const sheets: WorksheetData[] = [];

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    // Convert sheet to json array of objects
    // raw: false, defval: '' to ensure we get empty cell support if needed
    const rawData = XLSX.utils.sheet_to_json<ExcelRow>(worksheet, { header: 1, defval: '' });

    if (rawData.length === 0) {
      return; // Skip empty sheet
    }

    // Header is the first row
    const headers = rawData[0] as string[];
    // Filter out completely empty header elements
    const validHeaders = headers.map((h, index) => {
      const name = String(h).trim();
      return {
        key: name || `Kolom_${index + 1}`,
        name: name || `Kolom ${index + 1}`,
        index
      };
    });

    const rows: ExcelRow[] = [];
    // Slice raw data from row 1 to bottom
    const dataRows = rawData.slice(1);

    dataRows.forEach((rowArray: any) => {
      // Check if row has some data to avoid pushing blank rows
      const hasData = rowArray.some((cell: any) => cell !== null && cell !== undefined && cell !== '');
      if (!hasData) return;

      const rowObj: ExcelRow = {};
      validHeaders.forEach((header) => {
        const val = rowArray[header.index];
        rowObj[header.key] = val !== undefined && val !== null ? val : '';
      });
      rows.push(rowObj);
    });

    // Detect column types
    const columns: ExcelColumn[] = validHeaders.map(header => {
      const key = header.key;
      const valuesInCol = rows.map(r => r[key]);
      const type = detectColumnType(valuesInCol);
      return {
        key,
        name: header.name,
        type
      };
    });

    sheets.push({
      name: sheetName,
      columns,
      rows
    });
  });

  return {
    fileName,
    fileSize: formatBytes(fileSize),
    sheets: sheets.length > 0 ? sheets : [{ name: 'Sheet1', columns: [], rows: [] }],
    activeSheetIndex: 0
  };
}

/**
 * Generate highly polished and comprehensive Indonesian Sales Data Sample
 */
export function generateSampleData(): ExcelWorkbookData {
  const columns: ExcelColumn[] = [
    { key: 'ID Transaksi', name: 'ID Transaksi', type: 'string' },
    { key: 'Tanggal', name: 'Tanggal', type: 'date' },
    { key: 'Nama Produk', name: 'Nama Produk', type: 'string' },
    { key: 'Kategori', name: 'Kategori', type: 'string' },
    { key: 'Jumlah Terjual', name: 'Jumlah Terjual', type: 'number' },
    { key: 'Harga Satuan', name: 'Harga Satuan', type: 'number' },
    { key: 'Total Pendapatan', name: 'Total Pendapatan', type: 'number' },
    { key: 'Status Pembayaran', name: 'Status Pembayaran', type: 'string' },
    { key: 'Kota', name: 'Kota', type: 'string' },
    { key: 'Pelanggan Setia', name: 'Pelanggan Setia', type: 'boolean' }
  ];

  const categories = ['Elektronik', 'Fashion', 'Makanan & Minuman', 'Peralatan Rumah', 'Hobi & Hiburan'];
  const products: Record<string, string[]> = {
    'Elektronik': ['Laptop Gaming Ryzen 7', 'Smartphone Redmi 12', 'Earbuds Bluetooth ANC', 'Smartwatch Fit Pro', 'Monitor IPS 24 Inch'],
    'Fashion': ['Kemeja Batik Premium', 'Jaket Bomber Fleece', 'Sepatu Sneakers Canvas', 'Kaos Polos Cotton 30s', 'Tas Ransel Anti Air'],
    'Makanan & Minuman': ['Kopi Arabika Gayo 250g', 'Madu Hutan Murni 500ml', 'Camilan Keripik Tempe Premium', 'Teh Hijau Matcha Bubuk', 'Cokelat Artisan Luar Biasa'],
    'Peralatan Rumah': ['Panci Presto Teflon', 'Mesin Pembuat Kopi Espresso', 'Lampu Meja LED Fleksibel', 'Air Purifier Hepa Filter', 'Sapu Pembersih Microfiber'],
    'Hobi & Hiburan': ['Board Game Monopoli Klasik', 'Buku Novel Fiksi Terlaris', 'Kabel Gitar Premium 3m', 'Set Cat Akrilik 24 Warna', 'Raket Bulutangkis Karbon']
  };

  const cities = ['Jakarta', 'Surabaya', 'Bandung', 'Medan', 'Semarang', 'Yogyakarta', 'Denpasar', 'Makassar'];
  const payments = ['Lunas', 'Menunggu Pembayaran', 'Dicicil', 'Dibatalkan'];

  const rows: ExcelRow[] = [];
  const baseDate = new Date(2026, 0, 1); // 1 Jan 2026

  for (let i = 1; i <= 50; i++) {
    const trxId = `TRX-2026-${String(i).padStart(4, '0')}`;
    // Random dates distributed over early 2026
    const trxDate = new Date(baseDate.getTime() + Math.random() * 90 * 24 * 60 * 60 * 1000);
    const dateStr = trxDate.toISOString().split('T')[0];

    // Pick random category & product
    const cat = categories[Math.floor(Math.random() * categories.length)];
    const prodList = products[cat];
    const prodName = prodList[Math.floor(Math.random() * prodList.length)];

    // Random count and price
    const qty = Math.floor(Math.random() * 8) + 1; // 1 to 8
    let priceScalar = 10000;
    if (cat === 'Elektronik') priceScalar = 250000;
    else if (cat === 'Peralatan Rumah') priceScalar = 75000;
    else if (cat === 'Fashion') priceScalar = 45000;
    else if (cat === 'Makanan & Minuman') priceScalar = 15000;
    else priceScalar = 30000;

    const unitPrice = (Math.floor(Math.random() * 10) + 1) * priceScalar;
    const totalRev = qty * unitPrice;

    // payment status skew towards Lunas
    const payRoll = Math.random();
    const payStatus = payRoll < 0.7 ? 'Lunas' : payRoll < 0.85 ? 'Menunggu Pembayaran' : payRoll < 0.95 ? 'Dicicil' : 'Dibatalkan';

    const city = cities[Math.floor(Math.random() * cities.length)];
    const isLoyal = Math.random() > 0.6; // true/false

    rows.push({
      'ID Transaksi': trxId,
      'Tanggal': dateStr,
      'Nama Produk': prodName,
      'Kategori': cat,
      'Jumlah Terjual': qty,
      'Harga Satuan': unitPrice,
      'Total Pendapatan': totalRev,
      'Status Pembayaran': payStatus,
      'Kota': city,
      'Pelanggan Setia': isLoyal
    });
  }

  // Sort rows by date ascending
  rows.sort((a, b) => new Date(a['Tanggal']).getTime() - new Date(b['Tanggal']).getTime());

  return {
    fileName: 'Data_Sampel_Penjualan_2026.xlsx',
    fileSize: '18.4 KB',
    sheets: [
      {
        name: 'Penjualan Retail',
        columns,
        rows
      },
      {
        name: 'Analisis Sederhana',
        columns: [
          { key: 'Kategori', name: 'Kategori', type: 'string' },
          { key: 'Jumlah Produk', name: 'Jumlah Produk Special', type: 'number' },
          { key: 'Total Unit', name: 'Total Unit', type: 'number' }
        ],
        rows: categories.map(cat => ({
          'Kategori': cat,
          'Jumlah Produk Special': products[cat].length,
          'Total Unit': Math.floor(Math.random() * 100) + 40
        }))
      }
    ],
    activeSheetIndex: 0
  };
}
