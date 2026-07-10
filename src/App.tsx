import { useState, useEffect, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { 
  Download, 
  Printer, 
  Search, 
  ChevronDown, 
  Loader2, 
  Calendar, 
  Layers, 
  AlertCircle,
  CheckSquare,
  Square,
  X,
  FileSpreadsheet,
  BarChart3
} from 'lucide-react';

const INDEX_SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSMQ_KJlIyAnZJJYKMjKqKWaM226gI8W00KBx7fUBvUFh6JC5RXNL2Ix28rLR7dZA/pub?output=csv";

// Type definitions
interface IndexMonth {
  monthStr: string;
  iframeUrl: string;
  csvUrl: string;
}

interface BudgetItem {
  [key: number]: any;
  _e: number;
  _f: number;
  _g: number;
  _h: number;
  _compC: string;
  _compCSearch: string;
  _invD: string;
  _valM: string;
  _dept: string;
  _status: string;
  _taxId: string;
  _dateRecTs: number | null;
  _datePayTs: number | null;
  _dateProcessTs: number | null;
  _dateRecStr: string;
  _datePayStr: string;
  _dateProcessStr: string;
  _dateProcessStrShort: string;
}

interface MonthConfig {
  m: number;
  label: string;
}

// Utility functions
function parseNumberValue(v: any): number {
  if (!v) return 0;
  if (typeof v === 'string') v = v.replace(/,/g, '');
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function formatNumber(n: number): string {
  return n === 0 ? '0.00' : n.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getJSDate(v: any): Date | null {
  if (!v) return null;
  if (typeof v === 'number') return new Date(new Date(Date.UTC(1899, 11, 30)).getTime() + v * 86400000);
  
  const p = String(v).trim().split(/[-/]/);
  if (p.length === 3) {
    let d = parseInt(p[0]), m = parseInt(p[1]) - 1, y = parseInt(p[2]);
    if (p[0].length === 4) {
      y = parseInt(p[0]); m = parseInt(p[1]) - 1; d = parseInt(p[2]);
    }
    if (y < 100) y += y >= 50 ? 2500 : 2000;
    if (y > 2400) y -= 543;
    return new Date(y, m, d);
  }
  const parsed = new Date(v);
  return isNaN(parsed.getTime()) ? null : parsed;
}

const parseDateForFilter = (s: any): number | null => {
  if (!s || String(s).trim() === '-' || String(s).trim() === '') return null;
  let c = String(s).trim();
  
  const p = c.split(/[\/\-]/);
  if (p.length === 3) {
    let d = parseInt(p[0], 10), m = parseInt(p[1], 10) - 1, y = parseInt(p[2], 10);
    if (p[0].length === 4) {
      y = parseInt(p[0], 10); m = parseInt(p[1], 10) - 1; d = parseInt(p[2], 10);
    }
    if (y < 100) y += y > 40 ? 2500 : 2000;
    if (y >= 2400) y -= 543;
    let dObj = new Date(y, m, d);
    if (!isNaN(dObj.getTime())) return dObj.getTime();
  }
  
  const tS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  const tF = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const tM = c.match(/(\d{1,2})\s+([ก-๙\.]+)\s+(\d{2,4})/);
  
  if (tM) {
    let d = parseInt(tM[1], 10), mo = tM[2], y = parseInt(tM[3], 10);
    let m = tS.findIndex(t => mo.includes(t));
    if (m === -1) m = tF.findIndex(t => mo.includes(t));
    if (m === -1) m = 0;
    if (y < 100) y += y > 40 ? 2500 : 2000;
    if (y >= 2400) y -= 543;
    return new Date(y, m, d).getTime();
  }
  
  if (!isNaN(c as any) && Number(c) > 10000) {
    return new Date(Date.UTC(1899, 11, 30)).getTime() + Number(c) * 86400000;
  }
  
  let dObj = new Date(c);
  return isNaN(dObj.getTime()) ? null : dObj.getTime();
};

function formatThaiDate(v: any, s = false): string {
  if (!v) return '-';
  let d: number, m: number, y: number;
  if (typeof v === 'number') {
    const doObj = new Date(new Date(Date.UTC(1899, 11, 30)).getTime() + v * 86400000);
    d = doObj.getUTCDate();
    m = doObj.getUTCMonth() + 1;
    y = doObj.getUTCFullYear();
    if (y < 2500) y += 543;
  } else {
    let ds = String(v).trim();
    const p = ds.split(/[-/]/);
    if (p.length === 3) {
      d = parseInt(p[0]); m = parseInt(p[1]); y = parseInt(p[2]);
      if (p[0].length === 4) {
        y = parseInt(p[0]); m = parseInt(p[1]); d = parseInt(p[2]);
      }
      if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
        if (y < 100) y += y >= 50 ? 2500 : 2000;
        if (y < 2500) y += 543;
      } else return ds;
    } else return ds;
  }
  return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${s ? String(y).slice(-2) : y}`;
}

function formatThaiDateShort(v: any): string {
  if (!v) return '-';
  const jd = getJSDate(v);
  if (!jd) return String(v);
  const tS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return `${jd.getDate()} ${tS[jd.getMonth()]} ${String(jd.getFullYear() + 543).slice(-2)}`;
}

function getFiscalYear(dateTs: number | null): string | null {
  if (!dateTs) return null;
  const d = new Date(dateTs);
  let y = d.getFullYear(); 
  // Thai fiscal year starts on October 1st (Month 9 in JavaScript Date)
  if (d.getMonth() >= 9) y += 1; 
  return String(y + 543);
}

function formatMonthYearThai(monthStr: string): string {
  const parts = monthStr.split('/');
  if (parts.length === 2) {
    const monthVal = parseInt(parts[0], 10);
    const yearVal = parseInt(parts[1], 10);
    const thaiMonths = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    if (monthVal >= 1 && monthVal <= 12) {
      const thaiYear = yearVal + 543;
      return `${thaiMonths[monthVal - 1]} ${thaiYear}`;
    }
  }
  return monthStr;
}

function extractUrlFromIframe(iframeStr: string): string | null {
  const match = iframeStr.match(/src="([^"]+)"/);
  if (match) {
    return match[1].replace(/&amp;/g, '&');
  }
  const match2 = iframeStr.match(/src='([^']+)'/);
  if (match2) {
    return match2[1].replace(/&amp;/g, '&');
  }
  return null;
}

function convertToCsvUrl(url: string): string {
  if (url.includes('/pubhtml')) {
    return url.split('/pubhtml')[0] + '/pub?output=csv';
  }
  return url;
}

const fiscalMonths: MonthConfig[] = [
  { m: 9, label: 'ต.ค.' }, { m: 10, label: 'พ.ย.' }, { m: 11, label: 'ธ.ค.' },
  { m: 0, label: 'ม.ค.' }, { m: 1, label: 'ก.พ.' }, { m: 2, label: 'มี.ค.' },
  { m: 3, label: 'เม.ย.' }, { m: 4, label: 'พ.ค.' }, { m: 5, label: 'มิ.ย.' },
  { m: 6, label: 'ก.ค.' }, { m: 7, label: 'ส.ค.' }, { m: 8, label: 'ก.ย.' }
];

export default function App() {
  // Index loading states
  const [indexLoading, setIndexLoading] = useState<boolean>(true);
  const [indexError, setIndexError] = useState<string | null>(null);
  const [monthsList, setMonthsList] = useState<IndexMonth[]>([]);
  const [selectedMonth, setSelectedMonth] = useState<string>('');

  // Monthly data loading states
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [dataError, setDataError] = useState<string | null>(null);
  const [data, setData] = useState<BudgetItem[]>([]);

  // Category selection states
  const [availableCategories, setAvailableCategories] = useState<string[]>([]);
  const [selectedCategories, setSelectedCategories] = useState<string[]>([]);
  const [categorySearch, setCategorySearch] = useState<string>('');
  const [dropdownOpen, setDropdownOpen] = useState<boolean>(false);

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown on outside click
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setDropdownOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  // Fetch the INDEX Google Sheet first
  useEffect(() => {
    Papa.parse(INDEX_SHEET_CSV_URL, {
      download: true,
      header: false,
      complete: (results) => {
        try {
          const rows = results.data as string[][];
          // Skip header row if it contains 'ปี'
          const cleanRows = rows.filter(r => r.length >= 2 && r[0] !== 'ปี' && r[0] !== '');
          
          const parsedMonths = cleanRows.map(r => {
            const monthStr = r[0].trim();
            const iframeCode = r[1] || '';
            const iframeUrl = extractUrlFromIframe(iframeCode) || '';
            const csvUrl = iframeUrl ? convertToCsvUrl(iframeUrl) : '';
            return { monthStr, iframeUrl, csvUrl };
          }).filter(m => m.monthStr && m.csvUrl);

          setMonthsList(parsedMonths);
          
          if (parsedMonths.length > 0) {
            // Select the first month by default
            setSelectedMonth(parsedMonths[0].monthStr);
          }
          setIndexLoading(false);
        } catch (err: any) {
          setIndexError(`จัดเตรียมดัชนีรายเดือนล้มเหลว: ${err.message}`);
          setIndexLoading(false);
        }
      },
      error: (err) => {
        setIndexError(`โหลดไฟล์ดัชนีหลักผิดพลาด: ${err.message}`);
        setIndexLoading(false);
      }
    });
  }, []);

  // Fetch the active month's budget sheet whenever selectedMonth changes
  useEffect(() => {
    if (!selectedMonth || monthsList.length === 0) return;
    
    const activeMonthObj = monthsList.find(m => m.monthStr === selectedMonth);
    if (!activeMonthObj) return;

    setDataLoading(true);
    setDataError(null);

    Papa.parse(activeMonthObj.csvUrl, {
      download: true,
      header: false,
      complete: (results) => {
        try {
          if (results.data.length < 5) {
            setDataError("ข้อมูลไฟล์รายเดือนนี้ไม่สมบูรณ์หรือไม่เพียงพอ");
            setDataLoading(false);
            return;
          }

          const rawRows = results.data.slice(4) as any[];
          const cleanRows = rawRows.filter(r => r.length > 0 && r.some((c: any) => c !== ""));
          
          const enriched: BudgetItem[] = cleanRows.map(r => {
            const dateRecTs = parseDateForFilter(r[1]);
            const datePayTs = parseDateForFilter(r[8]);
            const jsDateProc = getJSDate(r[23]);
            
            return {
              ...r,
              _e: parseNumberValue(r[4]),
              _f: parseNumberValue(r[5]),
              _g: parseNumberValue(r[6]),
              _h: parseNumberValue(r[7]),
              _compC: String(r[2] || '').trim(),
              _compCSearch: String(r[2] || '').trim().toLowerCase().replace(/\s+/g, ''),
              _invD: r[3] || '-',
              _valM: String(r[12] || '').trim() || 'ไม่ระบุหมวดหมู่',
              _dept: String(r[13] || '').trim(),
              _status: String(r[22] || '').trim(),
              _taxId: String(r[21] || '').trim(),
              _dateRecTs: dateRecTs,
              _datePayTs: datePayTs,
              _dateProcessTs: jsDateProc ? jsDateProc.getTime() : null,
              _dateRecStr: formatThaiDate(r[1]),
              _datePayStr: formatThaiDate(r[8]),
              _dateProcessStr: formatThaiDate(r[23], true),
              _dateProcessStrShort: formatThaiDateShort(r[23]),
            };
          });

          setData(enriched);

          // Extract and update unique categories
          const catsSet = new Set<string>();
          enriched.forEach(row => {
            if (row._e === 0 && row._f === 0 && row._g === 0 && row._h === 0) return;
            catsSet.add(row._valM);
          });
          const catsList = Array.from(catsSet).sort();
          setAvailableCategories(catsList);
          
          // Auto select all categories on first load of this month
          setSelectedCategories(catsList);

          setDataLoading(false);
        } catch (err: any) {
          setDataError(`ประมวลผลข้อมูลรายเดือนล้มเหลว: ${err.message}`);
          setDataLoading(false);
        }
      },
      error: (err) => {
        setDataError(`ดาวน์โหลดข้อมูลจากลิงก์รายเดือนผิดพลาด: ${err.message}`);
        setDataLoading(false);
      }
    });
  }, [selectedMonth, monthsList]);

  // Filter calculations based on categories
  const baseFilteredData = data.filter(row => {
    if (row._e === 0 && row._f === 0 && row._g === 0 && row._h === 0) return false;
    if (selectedCategories.length === 0) return false;
    if (!selectedCategories.includes(row._valM)) return false;
    return true;
  });

  // Extract the target year of the loaded data to render its table
  const uniqueYearsInLoadedData = Array.from(new Set(baseFilteredData.map(row => getFiscalYear(row._dateRecTs)).filter(Boolean) as string[])).sort((a, b) => Number(b) - Number(a));

  const handleToggleCategory = (cat: string) => {
    setSelectedCategories(prev => 
      prev.includes(cat) ? prev.filter(item => item !== cat) : [...prev, cat]
    );
  };

  const handleSelectAllCategories = () => {
    setSelectedCategories([...availableCategories]);
  };

  const handleClearAllCategories = () => {
    setSelectedCategories([]);
  };

  const filteredDropdownCats = availableCategories.filter(cat => 
    cat.toLowerCase().replace(/\s+/g, '').includes(categorySearch.toLowerCase().replace(/\s+/g, ''))
  );

  const handlePrint = () => {
    window.print();
  };

  const handleExportToExcel = () => {
    if (selectedCategories.length === 0 || uniqueYearsInLoadedData.length === 0) {
      alert('ไม่มีข้อมูลสำหรับดาวน์โหลด กรุณาเลือกหมวดหมู่ก่อนค่ะ');
      return;
    }

    const wb = XLSX.utils.book_new();

    uniqueYearsInLoadedData.forEach((year) => {
      const tableElement = document.getElementById(`budget-table-${year}`);
      if (tableElement) {
        const sheetName = `ปีงบ_${year}`;
        const ws = XLSX.utils.table_to_sheet(tableElement, { raw: true });
        
        const colWidths = [
          { wch: 8 },  // Index
          { wch: 45 }, // Category
        ];
        for (let i = 0; i < 12; i++) {
          colWidths.push({ wch: 15 });
        }
        colWidths.push({ wch: 18 });
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      }
    });

    const activeMonthObj = monthsList.find(m => m.monthStr === selectedMonth);
    const monthLabel = activeMonthObj ? activeMonthObj.monthStr.replace('/', '_') : 'monthly';
    const fileName = `รายงานแยกประเภทรายเดือน_${monthLabel}.xlsx`;

    XLSX.writeFile(wb, fileName);
  };

  const activeMonthObj = monthsList.find(m => m.monthStr === selectedMonth);

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col text-slate-800 font-sans selection:bg-blue-100">
      
      {/* Header Bar (No Back Button, aligning with instruction 3) */}
      <header className="bg-gradient-to-r from-blue-900 to-indigo-900 text-white px-6 py-4 shadow-md no-print flex flex-col sm:flex-row justify-between items-center gap-4 shrink-0 border-b border-blue-950">
        <div className="flex items-center gap-4 w-full sm:w-auto">
          <img 
            src="https://github.com/FinanceNangRong/financial-system/raw/main/NRHOS.png" 
            alt="NRHOS Logo" 
            className="h-10 w-auto bg-white rounded-lg p-1 shadow-inner shrink-0" 
            onError={(e) => { e.currentTarget.style.display = 'none'; }}
          />
          <div>
            <h1 className="text-lg md:text-xl font-bold tracking-tight">รายงานแยกประเภท ประจำปีงบประมาณ ย้อนหลัง</h1>
            <p className="text-xs text-blue-200 font-light">เลือกรายงานแยกประเภทประจำเดือนและปีงบประมาณย้อนหลัง</p>
          </div>
        </div>

        {/* Status indicator badge */}
        <div className="flex items-center gap-2 self-end sm:self-auto">
          {indexLoading || dataLoading ? (
            <div id="statusBadge" className="text-xs px-3 py-1.5 rounded-full bg-amber-500/15 text-amber-200 font-medium flex items-center gap-2 border border-amber-500/30">
              <span className="relative flex h-2 w-2">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-amber-500"></span>
              </span>
              กำลังโหลด...
            </div>
          ) : indexError || dataError ? (
            <div id="statusBadge" className="text-xs px-3 py-1.5 rounded-full bg-red-500/15 text-red-200 font-medium flex items-center gap-2 border border-red-500/30">
              <AlertCircle className="w-3 h-3 text-red-400" />
              เกิดข้อผิดพลาด
            </div>
          ) : (
            <div id="statusBadge" className="text-xs px-3 py-1.5 rounded-full bg-emerald-500/15 text-emerald-200 font-medium flex items-center gap-2 border border-emerald-500/30">
              <span className="relative flex h-2 w-2">
                <span className="animate-pulse absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-500"></span>
              </span>
              พร้อมใช้งาน
            </div>
          )}
        </div>
      </header>

      {/* Main Content View */}
      <main className="flex-1 p-4 md:p-6 lg:p-8 max-w-7xl mx-auto w-full flex flex-col gap-6">
        
        {/* Loading index file state */}
        {indexLoading && (
          <div className="flex-1 flex flex-col items-center justify-center py-20 text-slate-500 gap-4 no-print">
            <Loader2 className="w-12 h-12 text-blue-600 animate-spin" />
            <div className="text-center">
              <p className="font-semibold text-lg text-slate-700">กำลังเชื่อมต่อกับระบบรายงานหลัก...</p>
              <p className="text-sm text-slate-400 mt-1">กรุณารอสักครู่ ระบบกำลังนำเข้ารายชื่อเดือนทั้งหมดจากฐานข้อมูล</p>
            </div>
          </div>
        )}

        {/* Index loading error state */}
        {!indexLoading && indexError && (
          <div className="flex-1 max-w-2xl mx-auto w-full py-12 px-6 bg-red-50 border border-red-200 rounded-2xl flex flex-col items-center gap-4 text-center shadow-sm no-print mt-10">
            <div className="p-3 bg-red-100 rounded-full text-red-600">
              <AlertCircle className="w-8 h-8" />
            </div>
            <div>
              <h3 className="text-lg font-bold text-red-800">เชื่อมต่อระบบหลักล้มเหลว</h3>
              <p className="text-sm text-red-600/90 mt-2 whitespace-pre-wrap">{indexError}</p>
            </div>
            <button 
              onClick={() => window.location.reload()}
              className="mt-2 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-medium transition-colors"
            >
              คลิกเพื่อลองอีกครั้ง
            </button>
          </div>
        )}

        {/* Dynamic active view once index is available */}
        {!indexLoading && !indexError && (
          <>
            {/* Control Panel: Select Month & Filters (Hidden on print) */}
            <section id="filterSection" className="bg-white border border-slate-200 p-4 md:p-6 rounded-2xl shadow-sm no-print flex flex-col gap-6">
              
              <div className="flex flex-col md:flex-row md:items-end justify-between gap-6">
                
                {/* Left side: Month selection */}
                <div className="flex-1 flex flex-col sm:flex-row gap-5">
                  
                  {/* Select Month from Column A */}
                  <div className="flex flex-col gap-1.5 sm:w-64 shrink-0">
                    <label htmlFor="monthSelect" className="text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1.5">
                      <Calendar className="w-3.5 h-3.5 text-blue-600" />
                      เลือกเดือนที่ต้องการดูข้อมูล
                    </label>
                    <select 
                      id="monthSelect"
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(e.target.value)}
                      className="w-full border border-slate-200 hover:border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100 rounded-xl px-3 py-2.5 bg-slate-50 text-slate-800 font-semibold shadow-sm transition-all outline-none cursor-pointer text-sm"
                    >
                      {monthsList.map(m => (
                        <option key={m.monthStr} value={m.monthStr}>
                          {m.monthStr} ({formatMonthYearThai(m.monthStr)})
                        </option>
                      ))}
                    </select>
                  </div>

                </div>

                {/* Right side action buttons */}
                {!dataLoading && !dataError && (
                  <div className="flex gap-2.5 shrink-0 self-stretch md:self-auto justify-end">
                    <button 
                      onClick={handleExportToExcel}
                      className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-semibold rounded-xl text-sm shadow-sm transition-all focus:outline-none focus:ring-2 focus:ring-emerald-400 cursor-pointer"
                    >
                      <Download className="w-4 h-4" />
                      <span>ดาวน์โหลด Excel</span>
                    </button>
                    <button 
                      onClick={handlePrint}
                      className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-4 py-2.5 bg-slate-800 hover:bg-slate-700 active:bg-slate-900 text-white font-semibold rounded-xl text-sm shadow-sm transition-all focus:outline-none focus:ring-2 focus:ring-slate-400 cursor-pointer"
                    >
                      <Printer className="w-4 h-4" />
                      <span>พิมพ์รายงาน</span>
                    </button>
                  </div>
                )}

              </div>

              {/* Category selector option */}
              {!dataLoading && !dataError && (
                <div className="border-t border-slate-100 pt-4 flex flex-col gap-1.5 relative" ref={dropdownRef}>
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                    <Layers className="w-3.5 h-3.5 text-blue-600" />
                    ตัวกรองหมวดหมู่ค่าใช้จ่าย
                  </label>
                  
                  <button 
                    type="button"
                    onClick={() => setDropdownOpen(!dropdownOpen)}
                    className="w-full flex items-center justify-between border border-slate-200 hover:border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100 rounded-xl px-4 py-2.5 bg-slate-50 text-slate-800 font-medium shadow-sm transition-all text-sm text-left"
                  >
                    <span className="truncate max-w-xl">
                      {selectedCategories.length === 0 
                        ? '-- ไม่มีหมวดหมู่ที่เลือก --' 
                        : `${selectedCategories.length} หมวดหมู่ที่เลือก: ${selectedCategories.join(', ')}`
                      }
                    </span>
                    <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform duration-200 ${dropdownOpen ? 'rotate-180' : ''}`} />
                  </button>

                  {/* Dropdown elements */}
                  {dropdownOpen && (
                    <div 
                      className="absolute left-0 right-0 top-full mt-2 bg-white border border-slate-200 shadow-xl rounded-xl z-50 overflow-hidden flex flex-col max-h-80"
                    >
                      <div className="p-3 bg-slate-50 border-b border-slate-100 flex items-center justify-between shrink-0">
                        <div className="flex gap-4">
                          <button 
                            type="button" 
                            onClick={handleSelectAllCategories}
                            className="text-xs font-semibold text-blue-600 hover:text-blue-800 cursor-pointer"
                          >
                            เลือกทั้งหมด
                          </button>
                          <button 
                            type="button" 
                            onClick={handleClearAllCategories}
                            className="text-xs font-semibold text-red-500 hover:text-red-700 cursor-pointer"
                          >
                            ล้างทั้งหมด
                          </button>
                        </div>
                        <span className="text-xs text-slate-400 font-semibold">{selectedCategories.length} เลือกไว้</span>
                      </div>

                      <div className="px-3 py-2 border-b border-slate-100 flex items-center gap-2 shrink-0 bg-white">
                        <Search className="w-4 h-4 text-slate-400" />
                        <input 
                          type="text" 
                          placeholder="พิมพ์เพื่อกรองคำค้นหาหมวดหมู่..." 
                          value={categorySearch}
                          onChange={(e) => setCategorySearch(e.target.value)}
                          className="w-full text-sm outline-none border-0 p-1 text-slate-700"
                        />
                        {categorySearch && (
                          <button 
                            type="button" 
                            onClick={() => setCategorySearch('')}
                            className="text-slate-400 hover:text-slate-600"
                          >
                            <X className="w-3.5 h-3.5" />
                          </button>
                        )}
                      </div>

                      <div className="flex-1 overflow-y-auto py-1 px-1 bg-white">
                        {filteredDropdownCats.length > 0 ? (
                          filteredDropdownCats.map(cat => {
                            const isChecked = selectedCategories.includes(cat);
                            return (
                              <label 
                                key={cat}
                                className={`flex items-center px-3 py-2 rounded-lg cursor-pointer transition-all text-sm select-none ${
                                  isChecked 
                                    ? 'bg-blue-50 text-blue-900 font-medium' 
                                    : 'hover:bg-slate-50 text-slate-700'
                                }`}
                              >
                                <input 
                                  type="checkbox" 
                                  value={cat}
                                  checked={isChecked}
                                  onChange={() => handleToggleCategory(cat)}
                                  className="sr-only"
                                />
                                <div className="shrink-0 mr-3">
                                  {isChecked ? (
                                    <CheckSquare className="w-4.5 h-4.5 text-blue-600" />
                                  ) : (
                                    <Square className="w-4.5 h-4.5 text-slate-300" />
                                  )}
                                </div>
                                <span className="truncate">{cat}</span>
                              </label>
                            );
                          })
                        ) : (
                          <div className="text-center py-4 text-slate-400 text-sm">ไม่พบหมวดหมู่</div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </section>

            {/* Interactive analyzed reports and tables */}
            <section className="flex-1 flex flex-col gap-6">
                
                {/* Loader when active month sheet is parsing */}
                {dataLoading && (
                  <div className="flex-1 flex flex-col items-center justify-center py-20 text-slate-500 gap-4 bg-white rounded-2xl border border-slate-200 shadow-sm">
                    <Loader2 className="w-10 h-10 text-blue-600 animate-spin" />
                    <div className="text-center">
                      <p className="font-semibold text-slate-700">กำลังนำเข้าและคัดแยกหมวดหมู่สำหรับเดือน {selectedMonth}...</p>
                      <p className="text-xs text-slate-400 mt-1">กำลังคำนวณยอดเงินสะสมประจำวันและข้อมูลรายเดือน</p>
                    </div>
                  </div>
                )}

                {/* Monthly sheet loading error */}
                {!dataLoading && dataError && (
                  <div className="flex-1 py-12 px-6 bg-red-50 border border-red-200 rounded-2xl flex flex-col items-center gap-4 text-center shadow-sm">
                    <div className="p-3 bg-red-100 rounded-full text-red-600">
                      <AlertCircle className="w-6 h-6" />
                    </div>
                    <div>
                      <h3 className="font-bold text-red-800 text-sm md:text-base">ไม่สามารถประมวลผลข้อมูลประจำเดือน {selectedMonth} ได้</h3>
                      <p className="text-xs text-red-600/90 mt-1 whitespace-pre-wrap">{dataError}</p>
                    </div>
                  </div>
                )}

                {/* Empty table state */}
                {!dataLoading && !dataError && selectedCategories.length === 0 && (
                  <div className="flex-1 bg-white border border-dashed border-slate-300 py-16 px-6 rounded-2xl flex flex-col items-center text-center justify-center gap-4 shadow-sm">
                    <Layers className="w-10 h-10 text-slate-300" />
                    <p className="text-slate-500 font-semibold text-sm">กรุณาเลือกอย่างน้อยหนึ่งหมวดหมู่เพื่อแสดงผลตารางรายงาน</p>
                  </div>
                )}

                {/* Table display */}
                {!dataLoading && !dataError && selectedCategories.length > 0 && (
                  <div id="tableContainer" className="flex flex-col gap-8">
                    {uniqueYearsInLoadedData.map((currentYear) => {
                      const yearData = baseFilteredData.filter(row => getFiscalYear(row._dateRecTs) === currentYear);
                      
                      if (yearData.length === 0) return null;

                      // Pivot calculation
                      const pivotData: { [cat: string]: { total: number; [m: number]: number } } = {};
                      const totalCols: { total: number; [m: number]: number } = { total: 0 };
                      fiscalMonths.forEach(fm => {
                        totalCols[fm.m] = 0;
                      });

                      yearData.forEach(row => {
                        const cat = row._valM;
                        if (!pivotData[cat]) {
                          pivotData[cat] = { total: 0 };
                          fiscalMonths.forEach(fm => {
                            pivotData[cat][fm.m] = 0;
                          });
                        }
                        if (row._dateRecTs) {
                          const d = new Date(row._dateRecTs);
                          const monthStr = d.getMonth();
                          const valToSum = row._h;
                          
                          if (pivotData[cat][monthStr] !== undefined) {
                            pivotData[cat][monthStr] += valToSum;
                            pivotData[cat].total += valToSum;
                            totalCols[monthStr] += valToSum;
                            totalCols.total += valToSum;
                          }
                        }
                      });

                      const catKeys = Object.keys(pivotData).sort();

                      return (
                        <div key={currentYear} className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden flex flex-col p-4 md:p-6 print:p-0 print:border-none print:shadow-none break-inside-avoid">
                          
                          {/* Print Only Header */}
                          <div className="hidden print-only text-center mb-6">
                            <h2 className="text-xl font-bold text-black">รายงานแยกประเภท ประจำปีงบประมาณ ย้อนหลัง แบบสรุปรายเดือน</h2>
                            <p className="text-sm text-slate-700 mt-1">ประจำปีงบประมาณ พ.ศ. {currentYear} (ข้อมูลเดือน {selectedMonth})</p>
                          </div>

                          {/* Web View Header */}
                          <div className="no-print flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
                            <div className="flex items-center gap-2">
                              <span className="h-4 w-1 bg-blue-600 rounded"></span>
                              <h2 className="text-sm md:text-base font-bold text-slate-800">
                                รายงานประมวลผลข้อมูลประจำปีงบประมาณ พ.ศ. {currentYear}
                              </h2>
                            </div>
                            <span className="text-xs bg-blue-50 text-blue-700 px-2.5 py-1 rounded-full font-semibold">
                              ข้อมูลของเดือน {selectedMonth} ({formatMonthYearThai(selectedMonth)})
                            </span>
                          </div>

                          {/* Responsive Grid Table wrapper */}
                          <div className="overflow-x-auto max-w-full" style={{ maxHeight: '70vh' }}>
                            <table 
                              id={`budget-table-${currentYear}`}
                              className="min-w-full text-xs text-left border-collapse print-fixed-table"
                              style={{ tableLayout: 'auto' }}
                            >
                              <thead className="bg-slate-50 text-slate-700 border-b border-slate-200 sticky top-0 z-30">
                                <tr>
                                  <th className="px-3 py-3 font-bold text-center border border-slate-200 bg-slate-100 sticky left-0 z-40 w-[50px] min-w-[50px] print:static print:bg-slate-100">
                                    ลำดับ
                                  </th>
                                  <th className="px-4 py-3 font-bold border border-slate-200 bg-slate-100 sticky left-[50px] z-40 w-[240px] min-w-[240px] max-w-[240px] print:static print:bg-slate-100">
                                    ค่า
                                  </th>
                                  
                                  {/* 12 Month columns */}
                                  {fiscalMonths.map(fm => {
                                    const fy = parseInt(currentYear);
                                    const yearLabel = fm.m >= 9 ? ` ${(fy - 1).toString().slice(-2)}` : ` ${fy.toString().slice(-2)}`;
                                    return (
                                      <th 
                                        key={fm.m}
                                        className="px-2 py-3 font-semibold text-center border border-slate-200 bg-blue-50/50 text-slate-800 min-w-[95px] whitespace-nowrap print:static print:bg-blue-50/20"
                                      >
                                        {fm.label}{yearLabel}
                                      </th>
                                    );
                                  })}

                                  <th className="px-3 py-3 font-bold text-center border border-slate-200 bg-blue-100/70 text-blue-900 min-w-[120px] print:static print:bg-blue-100/40">
                                    รวมทั้งสิ้น
                                  </th>
                                </tr>
                              </thead>
                              
                              <tbody className="divide-y divide-slate-200 bg-white">
                                {catKeys.map((catKey, idx) => {
                                  const rowData = pivotData[catKey];
                                  return (
                                    <tr key={catKey} className="hover:bg-slate-50/80 transition-colors group">
                                      <td className="px-3 py-2 text-center text-slate-400 border border-slate-200 bg-white sticky left-0 z-20 group-hover:bg-slate-50/80 print:static">
                                        {idx + 1}
                                      </td>
                                      <td 
                                        className="px-4 py-2 text-slate-800 font-semibold border border-slate-200 bg-white sticky left-[50px] z-20 group-hover:bg-slate-50/80 truncate print-truncate w-[240px] max-w-[240px] print:static"
                                        title={catKey}
                                      >
                                        {catKey}
                                      </td>
                                      
                                      {/* 12 months values */}
                                      {fiscalMonths.map(fm => {
                                        const val = rowData[fm.m];
                                        return (
                                          <td 
                                            key={fm.m} 
                                            className="px-2 py-2 text-right text-slate-600 border border-slate-200 tabular-nums"
                                          >
                                            {val !== 0 ? formatNumber(val) : ''}
                                          </td>
                                        );
                                      })}

                                      {/* Total sum cell */}
                                      <td className="px-3 py-2 text-right text-blue-700 font-bold border border-slate-200 bg-blue-50/10 group-hover:bg-blue-50/30 tabular-nums">
                                        {formatNumber(rowData.total)}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>

                              <tfoot className="bg-slate-100 border-t-2 border-slate-300 font-bold text-slate-900 sticky bottom-0 z-30">
                                <tr>
                                  <td 
                                    colSpan={2} 
                                    className="px-4 py-3 text-right border border-slate-200 bg-slate-100 sticky left-0 z-40 font-bold print:static"
                                  >
                                    รวมยอดทั้งสิ้น
                                  </td>
                                  
                                  {/* Total per month footer columns */}
                                  {fiscalMonths.map(fm => {
                                    const val = totalCols[fm.m];
                                    return (
                                      <td 
                                        key={fm.m} 
                                        className="px-2 py-3 text-right border border-slate-200 bg-slate-100 tabular-nums text-slate-800"
                                      >
                                        {val !== 0 ? formatNumber(val) : '-'}
                                      </td>
                                    );
                                  })}

                                  {/* Column Total sum cell */}
                                  <td className="px-3 py-3 text-right text-blue-900 border border-slate-200 bg-blue-100 tabular-nums font-bold">
                                    {formatNumber(totalCols.total)}
                                  </td>
                                </tr>
                              </tfoot>
                            </table>
                          </div>

                        </div>
                      );
                    })}
                  </div>
                )}
              </section>
          </>
        )}
      </main>

      {/* Footer information bar */}
      <footer className="no-print mt-auto py-6 border-t border-slate-200 bg-white text-center text-slate-400 text-xs shrink-0">
        <p>© {new Date().getFullYear() + 543} กลุ่มงานคลัง โรงพยาบาลนางรอง • ระบบติดตามรายจ่ายแยกประเภทอัตโนมัติ</p>
      </footer>
    </div>
  );
}
