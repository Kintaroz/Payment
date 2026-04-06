import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { 
  FileUp, 
  CheckCircle2, 
  AlertCircle, 
  FileSpreadsheet,
  Trash2,
  Settings2,
  Copy,
  Download
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from './lib/utils';

interface RowData {
  [key: string]: any;
}

interface POGroup {
  poNumber: string;
  amount: number;
  netTotal: number;
  vat: number;
  receiptNo: string;
  rows: any[];
}

interface ProcessedData {
  supplier: string;
  totalAmount: number;
  type: 'IOT' | 'ATK';
  date: string;
  pdfDate: string;
  poGroups: POGroup[];
  originalAOA: any[][];
  validRows: any[];
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [data, setData] = useState<RowData[]>([]);
  const [processedData, setProcessedData] = useState<ProcessedData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [docType, setDocType] = useState<'IOT' | 'ATK'>('IOT');
  const [copiedStatement, setCopiedStatement] = useState(false);
  const [copiedDesc, setCopiedDesc] = useState(false);
  const [copiedNetTotal, setCopiedNetTotal] = useState(false);
  const [copiedReceipt, setCopiedReceipt] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (uploadedFile) {
      setFile(uploadedFile);
      processExcel(uploadedFile);
    }
  };

  const processExcel = (file: File) => {
    setIsProcessing(true);
    setError(null);
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws) as RowData[];
        const originalAOA = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        if (jsonData.length === 0) {
          throw new Error("ไฟล์ Excel ไม่มีข้อมูล");
        }

        setData(jsonData);
        
        // Extract valid rows that have Supplier d
        const validRows = jsonData.filter(row => row['Supplier d']);
        
        if (validRows.length === 0) {
          throw new Error("ไม่พบข้อมูลในคอลัมน์ 'Supplier d'");
        }

        const supplier = validRows[0]['Supplier d'];
        const parseAmount = (val: any) => parseFloat(String(val || '0').replace(/,/g, ''));

        // Group by PO
        const poMap = new Map<string, POGroup>();
        let calculatedTotal = 0;

        validRows.forEach(row => {
          const po = row['Purchase o'] || 'ไม่ระบุ PO';
          const amt = parseAmount(row['Total amou']);
          const net = parseAmount(row['Net total']);
          const tax = parseAmount(row['Tax']);
          const receipt = row['Receipt nu'] || '';
          
          if (!poMap.has(po)) {
            poMap.set(po, { poNumber: po, amount: 0, netTotal: 0, vat: 0, receiptNo: receipt, rows: [] });
          }
          const group = poMap.get(po)!;
          group.amount += amt;
          group.netTotal += net;
          group.vat += tax;
          group.rows.push(row);
          if (receipt && !group.receiptNo) group.receiptNo = receipt;
          
          calculatedTotal += amt;
        });

        // Find total amount (Look for summary row first, else sum valid rows)
        const summaryRow = jsonData.find(row => !row['Supplier d'] && row['Total amou']);
        const totalAmount = summaryRow ? parseAmount(summaryRow['Total amou']) : calculatedTotal;

        const d = new Date();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        const formattedDate = `${yyyy}.${d.getMonth() + 1}.${d.getDate()}`;
        const pdfDate = `${yyyy}/${mm}/${dd}`;

        const processed: ProcessedData[] = [{
          supplier: supplier,
          totalAmount: totalAmount,
          type: docType,
          date: formattedDate,
          pdfDate: pdfDate,
          poGroups: Array.from(poMap.values()),
          originalAOA: originalAOA,
          validRows: validRows
        }];

        setProcessedData(processed);
      } catch (err) {
        setError(err instanceof Error ? err.message : "เกิดข้อผิดพลาดในการอ่านไฟล์");
      } finally {
        setIsProcessing(false);
      }
    };

    reader.onerror = () => {
      setError("ไม่สามารถอ่านไฟล์ได้");
      setIsProcessing(false);
    };

    reader.readAsBinaryString(file);
  };

  const handleReceiptChange = (dataIdx: number, poIdx: number, value: string) => {
    const newData = [...processedData];
    newData[dataIdx].poGroups[poIdx].receiptNo = value;
    setProcessedData(newData);
  };

  const copyText = (text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedStatement(true);
    setTimeout(() => setCopiedStatement(false), 2000);
  };

  const copyDescData = (poGroups: POGroup[]) => {
    const text = poGroups.map(po => `${po.receiptNo || po.poNumber} 产线用具以及耗材 Factory consumables`).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      setCopiedDesc(true);
      setTimeout(() => setCopiedDesc(false), 2000);
    });
  };

  const copyNetTotalData = (poGroups: POGroup[]) => {
    const text = poGroups.map(po => po.netTotal.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      setCopiedNetTotal(true);
      setTimeout(() => setCopiedNetTotal(false), 2000);
    });
  };

  const copyReceiptData = (validRows: any[], poGroups: POGroup[]) => {
    const text = validRows.map(row => {
      const poNum = row['Purchase o'] || 'ไม่ระบุ PO';
      const poGroup = poGroups.find(g => g.poNumber === poNum);
      return poGroup?.receiptNo || '';
    }).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      setCopiedReceipt(true);
      setTimeout(() => setCopiedReceipt(false), 2000);
    });
  };

  const exportStatementExcel = async (item: ProcessedData) => {
    if (!file) return;
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.worksheets[0];
      
      // Insert a row at the top
      worksheet.spliceRows(1, 0, []);
      const firstRow = worksheet.getRow(1);
      firstRow.getCell(1).value = `${item.supplier} ${item.date} 对账单 ${item.type}`;
      firstRow.getCell(1).font = { bold: true, size: 14 };
      
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${item.supplier} ${item.date} 对账单 ${item.type}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      setError("เกิดข้อผิดพลาดในการสร้างไฟล์ Excel");
    }
  };

  const exportPaymentExcel = async (item: ProcessedData) => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Payment Voucher');
      
      // Set column widths
      worksheet.columns = [
        { width: 40 }, // Description
        { width: 15 }, // Net Total
        { width: 15 }, // VAT
        { width: 15 }, // Grand Total
        { width: 15 }, // WHT
        { width: 15 }, // AP payment
      ];

      // Header
      worksheet.mergeCells('A1:F1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = item.supplier;
      titleCell.font = { bold: true, size: 14 };
      titleCell.alignment = { horizontal: 'center' };

      worksheet.mergeCells('A2:F2');
      const subtitleCell = worksheet.getCell('A2');
      subtitleCell.value = 'Payment Voucher';
      subtitleCell.font = { bold: true, size: 12 };
      subtitleCell.alignment = { horizontal: 'center' };

      worksheet.mergeCells('A3:F3');
      const dateCell = worksheet.getCell('A3');
      dateCell.value = `Date: ${item.pdfDate}`;
      dateCell.alignment = { horizontal: 'right' };

      // Table Headers
      const headerRow = worksheet.addRow([
        'Description 摘要',
        'Ⓐ Net Total\n未稅',
        'Ⓑ VAT\n稅金',
        'ⒸGrand Total\n=Ⓐ+ Ⓑ 含稅',
        'ⒹWHT 預扣稅\n=Tax rate*Ⓐ',
        'AP payment\n=Ⓒ-Ⓓ應付金額'
      ]);
      headerRow.height = 30;
      headerRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
          top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8FAFC' }
        };
      });

      // Data Rows
      item.poGroups.forEach(po => {
        const row = worksheet.addRow([
          `${po.receiptNo || po.poNumber} 产线用具以及耗材 Factory consumables`,
          po.netTotal,
          po.vat,
          po.amount,
          '',
          po.amount
        ]);
        row.eachCell((cell, colNumber) => {
          cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
          };
          cell.alignment = { vertical: 'middle', wrapText: true };
          if (colNumber > 1) {
            cell.numFmt = '#,##0.00';
          }
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `ทำจ่าย 01 ${item.supplier}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      setError("เกิดข้อผิดพลาดในการสร้างไฟล์ Excel");
    }
  };

  const reset = () => {
    setFile(null);
    setData([]);
    setProcessedData([]);
    setError(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 p-2 rounded-lg">
              <FileSpreadsheet className="text-white w-6 h-6" />
            </div>
            <h1 className="text-xl font-bold tracking-tight text-slate-800">
              Payment Assistant
            </h1>
          </div>
          <div className="flex items-center gap-4">
            <div className="flex items-center bg-slate-100 p-1 rounded-lg">
              <button 
                onClick={() => setDocType('IOT')}
                className={cn(
                  "px-4 py-1.5 rounded-md text-sm font-medium transition-all",
                  docType === 'IOT' ? "bg-white shadow-sm text-indigo-600" : "text-slate-500 hover:text-slate-700"
                )}
              >
                IOT
              </button>
              <button 
                onClick={() => setDocType('ATK')}
                className={cn(
                  "px-4 py-1.5 rounded-md text-sm font-medium transition-all",
                  docType === 'ATK' ? "bg-white shadow-sm text-indigo-600" : "text-slate-500 hover:text-slate-700"
                )}
              >
                ATK
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          {/* Left Column: Upload & Config */}
          <div className="lg:col-span-4 space-y-6">
            <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                <FileUp className="w-5 h-5 text-indigo-500" />
                อัปโหลดไฟล์
              </h2>
              
              {!file ? (
                <div 
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-200 rounded-xl p-8 text-center hover:border-indigo-400 hover:bg-indigo-50/30 transition-all cursor-pointer group"
                >
                  <input 
                    type="file" 
                    ref={fileInputRef}
                    onChange={handleFileUpload}
                    accept=".xlsx, .xls"
                    className="hidden"
                  />
                  <div className="bg-slate-50 w-12 h-12 rounded-full flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                    <FileUp className="w-6 h-6 text-slate-400 group-hover:text-indigo-500" />
                  </div>
                  <p className="text-sm font-medium text-slate-600">คลิกเพื่ออัปโหลดไฟล์ Excel</p>
                  <p className="text-xs text-slate-400 mt-1">รองรับ .xlsx, .xls</p>
                </div>
              ) : (
                <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-3 overflow-hidden">
                      <div className="bg-green-100 p-2 rounded-lg shrink-0">
                        <CheckCircle2 className="w-5 h-5 text-green-600" />
                      </div>
                      <div className="overflow-hidden">
                        <p className="text-sm font-medium text-slate-700 truncate">{file.name}</p>
                        <p className="text-xs text-slate-400">{(file.size / 1024).toFixed(1)} KB</p>
                      </div>
                    </div>
                    <button 
                      onClick={reset}
                      className="p-2 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                  {isProcessing && (
                    <div className="h-1 w-full bg-slate-200 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: '100%' }}
                        className="h-full bg-indigo-500"
                      />
                    </div>
                  )}
                </div>
              )}

              {error && (
                <div className="mt-4 p-3 bg-red-50 border border-red-100 rounded-lg flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 text-red-500 shrink-0 mt-0.5" />
                  <p className="text-sm text-red-600">{error}</p>
                </div>
              )}
            </section>

            <section className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
                <Settings2 className="w-5 h-5 text-indigo-500" />
                คำแนะนำ
              </h2>
              <ul className="space-y-3 text-sm text-slate-600">
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  ระบบจะดึงชื่อซัพพลายเออร์จากคอลัมน์ <strong>Supplier d</strong>
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  ระบบจะดึงยอดเงินรวมจากคอลัมน์ <strong>Total amou</strong> (แถวสรุปสีเหลือง)
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  ระบบจะจัดกลุ่มข้อมูลตาม <strong>Purchase o</strong> (เลข PO)
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  คุณสามารถกรอก <strong>เลขใบกำกับภาษี</strong> แยกตามแต่ละ PO ได้ในหน้าจอ
                </li>
                <li className="flex gap-2">
                  <span className="text-indigo-500 font-bold">•</span>
                  เลือกประเภท (IOT/ATK) ก่อนอัปโหลดเพื่อใช้ในการตั้งชื่อไฟล์
                </li>
              </ul>
            </section>
          </div>

          {/* Right Column: Data & Actions */}
          <div className="lg:col-span-8">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden min-h-[500px] flex flex-col">
              <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                <h2 className="text-lg font-semibold">รายการที่ประมวลผล</h2>
                <span className="bg-slate-100 text-slate-600 px-3 py-1 rounded-full text-xs font-medium">
                  {processedData.length} รายการ
                </span>
              </div>

              <div className="flex-1 overflow-auto p-6">
                <AnimatePresence mode="wait">
                  {processedData.length > 0 ? (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: -10 }}
                      className="space-y-4"
                    >
                      {processedData.map((item, idx) => {
                        const statementTitle = `${item.supplier} ${item.date} 对账单 ${item.type}`;
                        return (
                        <div 
                          key={idx}
                          className="group bg-white border border-slate-200 rounded-xl p-4 hover:border-indigo-300 hover:shadow-md transition-all"
                        >
                          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                            <div className="space-y-1">
                              <div className="flex items-center gap-2">
                                <h3 className="font-bold text-slate-800">{item.supplier}</h3>
                                <span className={cn(
                                  "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                                  item.type === 'IOT' ? "bg-indigo-100 text-indigo-700" : "bg-amber-100 text-amber-700"
                                )}>
                                  {item.type}
                                </span>
                              </div>
                              <div className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-slate-500">
                                <span className="flex items-center gap-1">
                                  <span className="font-medium text-slate-400">ยอดเงินรวม:</span> 
                                  ฿{item.totalAmount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                                </span>
                                <span className="flex items-center gap-1">
                                  <span className="font-medium text-slate-400">จำนวน PO:</span> 
                                  {item.poGroups.length} รายการ
                                </span>
                              </div>
                            </div>
                          </div>

                          {/* PO Groups Input Section */}
                          <div className="mt-4 pt-4 border-t border-slate-100 space-y-3">
                            <h4 className="text-sm font-semibold text-slate-700">ระบุเลขใบกำกับภาษีตาม PO:</h4>
                            <div className="grid grid-cols-1 gap-3">
                              {item.poGroups.map((po, poIdx) => (
                                <div key={poIdx} className="flex flex-col sm:flex-row sm:items-center gap-3 bg-slate-50 p-3 rounded-lg border border-slate-200">
                                  <div className="flex-1 flex justify-between sm:block">
                                    <p className="text-sm font-medium text-slate-800">PO: {po.poNumber}</p>
                                    <p className="text-xs text-slate-500 sm:mt-1">
                                      ยอดเงิน: ฿{po.amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}
                                    </p>
                                  </div>
                                  <div className="flex-1">
                                    <input
                                      type="text"
                                      placeholder="กรอกเลขใบกำกับภาษี..."
                                      value={po.receiptNo}
                                      onChange={(e) => handleReceiptChange(idx, poIdx, e.target.value)}
                                      className="w-full text-sm px-3 py-2 border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white"
                                    />
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>

                          {/* Copy Data Section */}
                          <div className="mt-6 pt-6 border-t border-slate-200">
                            <h3 className="text-base font-semibold text-slate-800 mb-4 flex items-center gap-2">
                              <Copy className="w-4 h-4 text-indigo-500" />
                              ข้อมูลสำหรับคัดลอก (Copy Data)
                            </h3>

                            {/* Statement Header */}
                            <div className="mb-6">
                              <label className="block text-sm font-medium text-slate-700 mb-2">1. หัวเอกสาร Statement (对账单)</label>
                              <div className="flex items-center gap-3">
                                <div className="flex-1 bg-slate-50 p-3 rounded-lg border border-slate-200 font-mono text-sm text-slate-700">
                                  {statementTitle}
                                </div>
                                <button
                                  onClick={() => copyText(statementTitle)}
                                  className="flex items-center gap-2 px-4 py-2.5 bg-indigo-50 text-indigo-600 rounded-lg hover:bg-indigo-100 transition-colors font-medium text-sm shrink-0"
                                >
                                  {copiedStatement ? <CheckCircle2 className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                                  {copiedStatement ? 'คัดลอกแล้ว' : 'คัดลอก'}
                                </button>
                                <button
                                  onClick={() => exportStatementExcel(item)}
                                  className="flex items-center gap-2 px-4 py-2.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors font-medium text-sm shrink-0"
                                >
                                  <Download className="w-4 h-4" />
                                  ดาวน์โหลด Excel
                                </button>
                              </div>
                            </div>

                            {/* Payment Table */}
                            <div className="mb-6">
                              <div className="flex items-center justify-between mb-2">
                                <label className="block text-sm font-medium text-slate-700">2. ข้อมูลทำจ่าย (Payment Voucher)</label>
                                <div className="flex gap-2">
                                  <button
                                    onClick={() => copyDescData(item.poGroups)}
                                    className="flex items-center gap-2 px-4 py-2 bg-emerald-50 text-emerald-600 rounded-lg hover:bg-emerald-100 transition-colors font-medium text-sm"
                                  >
                                    {copiedDesc ? <CheckCircle2 className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                                    {copiedDesc ? 'คัดลอกแล้ว' : 'คัดลอก Description'}
                                  </button>
                                  <button
                                    onClick={() => copyNetTotalData(item.poGroups)}
                                    className="flex items-center gap-2 px-4 py-2 bg-amber-50 text-amber-600 rounded-lg hover:bg-amber-100 transition-colors font-medium text-sm"
                                  >
                                    {copiedNetTotal ? <CheckCircle2 className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                                    {copiedNetTotal ? 'คัดลอกแล้ว' : 'คัดลอก Net Total'}
                                  </button>
                                  <button
                                    onClick={() => exportPaymentExcel(item)}
                                    className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors font-medium text-sm"
                                  >
                                    <Download className="w-4 h-4" />
                                    ดาวน์โหลด Excel
                                  </button>
                                </div>
                              </div>
                              <div className="overflow-x-auto border border-slate-200 rounded-lg bg-white">
                                <table id={`payment-table-${idx}`} className="w-full text-sm text-center" border={1} style={{ borderCollapse: 'collapse', width: '100%', borderColor: '#e2e8f0' }}>
                                  <thead className="bg-slate-50 text-slate-700">
                                    <tr>
                                      <th colSpan={4} className="border p-2 font-medium" style={{ border: '1px solid #e2e8f0', padding: '8px' }}>Description 摘要</th>
                                      <th className="border p-2 font-medium" style={{ border: '1px solid #e2e8f0', padding: '8px' }}>Ⓐ Net Total<br/>未稅</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {item.poGroups.map(po => (
                                      <tr key={po.poNumber} className="border-b">
                                        <td colSpan={4} className="border p-2 text-left" style={{ border: '1px solid #e2e8f0', padding: '8px', textAlign: 'left' }}>
                                          {po.receiptNo || po.poNumber} 产线用具以及耗材 Factory consumables
                                        </td>
                                        <td className="border p-2" style={{ border: '1px solid #e2e8f0', padding: '8px' }}>{po.netTotal.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                            {/* Receipt Column */}
                            <div>
                              <div className="flex items-center justify-between mb-2">
                                <label className="block text-sm font-medium text-slate-700">3. คอลัมน์ Receipt nu (สำหรับนำไปวางในไฟล์ต้นฉบับ)</label>
                                <button
                                  onClick={() => copyReceiptData(item.validRows, item.poGroups)}
                                  className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors font-medium text-sm"
                                >
                                  {copiedReceipt ? <CheckCircle2 className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                                  {copiedReceipt ? 'คัดลอกคอลัมน์แล้ว' : 'คัดลอกคอลัมน์'}
                                </button>
                              </div>
                              <div className="overflow-y-auto max-h-48 border border-slate-200 rounded-lg bg-white w-64">
                                <table id={`receipt-table-${idx}`} className="w-full text-sm text-center" border={1} style={{ borderCollapse: 'collapse', width: '100%', borderColor: '#e2e8f0' }}>
                                  <tbody>
                                    {item.validRows.map((row, i) => {
                                      const poNum = row['Purchase o'] || 'ไม่ระบุ PO';
                                      const poGroup = item.poGroups.find(g => g.poNumber === poNum);
                                      return (
                                        <tr key={i} className="border-b">
                                          <td className="border p-2" style={{ border: '1px solid #e2e8f0', padding: '8px', height: '37px' }}>
                                            {poGroup?.receiptNo || ' '}
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          </div>
                        </div>
                      )})}
                    </motion.div>
                  ) : (
                    <div className="h-full flex flex-col items-center justify-center text-center py-20">
                      <div className="bg-slate-50 w-16 h-16 rounded-full flex items-center justify-center mb-4">
                        <FileSpreadsheet className="w-8 h-8 text-slate-300" />
                      </div>
                      <h3 className="text-slate-500 font-medium">ยังไม่มีข้อมูล</h3>
                      <p className="text-slate-400 text-sm mt-1 max-w-xs">
                        อัปโหลดไฟล์ Excel เพื่อเริ่มประมวลผลเอกสารทำจ่ายและใบแจ้งยอด
                      </p>
                    </div>
                  )}
                </AnimatePresence>
              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="max-w-7xl mx-auto px-4 py-8 text-center text-slate-400 text-xs">
        <p>© 2026 Payment & Document Assistant. All rights reserved.</p>
      </footer>
    </div>
  );
}
