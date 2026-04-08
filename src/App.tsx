import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import Holidays from 'date-holidays';
import { saveAs } from 'file-saver';
import { motion, AnimatePresence } from 'motion/react';
import { Upload, FileText, Download, Trash2, AlertCircle, CheckCircle2, ChevronRight, FileSpreadsheet } from 'lucide-react';
import { generateAttendanceReport, generateBulkAttendanceReport } from './services/docGenerator';

type ReportType = '질병' | '출석인정' | '기타';
const koreanHolidays = new Holidays('KR');

interface AttendanceRecord {
  id: string;
  date: string;
  number: string;
  name: string;
  type: string;
  reportType: ReportType;
  reason: string;
  year: string;
  class: string;
  teacherName: string;
}

interface GroupedReport {
  id: string;
  name: string;
  number: string;
  schoolYear: string;
  year: string;
  class: string;
  type: ReportType;
  reason: string;
  startDate: string;
  endDate: string;
  days: number;
  records: AttendanceRecord[];
  teacherName: string;
  guardianName: string;
}

export default function App() {
  const [records, setRecords] = useState<AttendanceRecord[]>([]);
  const [reports, setReports] = useState<GroupedReport[]>([]);
  const [guardianMap, setGuardianMap] = useState<Record<string, string>>({});
  const [isDragging, setIsDragging] = useState(false);
  const [isDraggingMap, setIsDraggingMap] = useState(false);
  const [config, setConfig] = useState({
    year: '1',
    class: '1',
    teacherName: '박선일'
  });

  const getSchoolYear = (dateStr: string) => {
    const matchedYear = dateStr.match(/^(\d{4})[.-]/)?.[1];
    return matchedYear || String(new Date().getFullYear());
  };

  const normalizeReportType = (rawType: string): ReportType | null => {
    const normalized = rawType.trim().toLowerCase();
    if (!normalized) return null;
    if (normalized.includes('미인정')) return null;
    if (normalized.includes('조퇴') || normalized.includes('지각') || normalized.includes('결과')) return null;
    if (normalized.includes('질병')) return '질병';
    if (normalized.includes('출석인정') || normalized.includes('인정')) return '출석인정';
    if (normalized.includes('기타') || normalized.includes('사고') || normalized.includes('결석')) return '기타';
    return null;
  };

  const getDefaultReason = (type: ReportType) => (
    type === '질병' ? '몸살' : '기타 사유'
  );

  const updateYear = (value: string) => {
    setConfig(prev => ({ ...prev, year: value }));
    setRecords(prev => prev.map(record => ({ ...record, year: value })));
    setReports(prev => prev.map(report => ({ ...report, year: value })));
  };

  const updateClass = (value: string) => {
    setConfig(prev => ({ ...prev, class: value }));
    setRecords(prev => prev.map(record => ({ ...record, class: value })));
    setReports(prev => prev.map(report => ({ ...report, class: value })));
  };

  const updateTeacherName = (value: string) => {
    setConfig(prev => ({ ...prev, teacherName: value }));
    setRecords(prev => prev.map(record => ({ ...record, teacherName: value })));
    setReports(prev => prev.map(report => ({ ...report, teacherName: value })));
  };

  const processGuardianMap = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Find the header row dynamically
      const allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      let headerIndex = 0;
      for (let i = 0; i < Math.min(allRows.length, 20); i++) {
        const row = allRows[i];
        if (!row || row.length === 0) continue;
        const rowStr = JSON.stringify(row);
        if (rowStr.includes('성명') || rowStr.includes('이름') || rowStr.includes('보호자') || rowStr.includes('학부모')) {
          headerIndex = i;
          break;
        }
      }

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerIndex }) as any[];

      const map: Record<string, string> = {};
      jsonData.forEach(row => {
        const studentName = String(row['성명'] || row['이름'] || row['학생명'] || row['학생'] || row['성함'] || '').trim();
        const guardianName = String(row['보호자'] || row['보호자명'] || row['학부모'] || row['학부모명'] || row['보호자성명'] || '').trim();
        const studentNumber = String(row['번호'] || row['학번'] || row['출석번호'] || '').trim();

        if (studentName && guardianName) {
          // Map by name
          map[studentName] = guardianName;
          // Also map by number-name if number exists
          if (studentNumber) {
            map[`${studentNumber}-${studentName}`] = guardianName;
          }
        }
      });
      setGuardianMap(map);
      
      // If reports already exist, update them
      if (reports.length > 0) {
        setReports(prev => prev.map(report => ({
          ...report,
          guardianName: map[`${report.number}-${report.name.trim()}`] || map[report.name.trim()] || report.guardianName || ''
        })));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processExcel = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target?.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Find the header row dynamically
      const allRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      let headerIndex = 0;
      for (let i = 0; i < Math.min(allRows.length, 20); i++) {
        const row = allRows[i];
        if (!row || row.length === 0) continue;
        const rowStr = JSON.stringify(row);
        if (rowStr.includes('성명') || rowStr.includes('이름') || rowStr.includes('출결') || rowStr.includes('구분') || rowStr.includes('학생명')) {
          headerIndex = i;
          break;
        }
      }

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerIndex, defval: '' }) as any[];

      let lastValidName = '';
      let lastValidNumber = '';
      let lastValidType = '';
      let lastValidReason = '';

      const parsedRecords: AttendanceRecord[] = jsonData.map((row, index) => {
        // Robust column mapping with more aliases
        let type = String(row['출결구분'] || row['구분'] || row['출결'] || row['출결상황'] || row['종류'] || row['항목'] || row['상태'] || '').trim();
        let reason = String(row['사유'] || row['결석사유'] || row['비고'] || row['상세사유'] || row['내용'] || row['사유(증빙서류)'] || '').trim();
        let name = String(row['성명'] || row['이름'] || row['학생명'] || row['학생'] || row['성함'] || row['피교육자'] || '').trim();
        let number = String(row['번호'] || row['학번'] || row['출석번호'] || '').trim();
        
        // Handle merged cells / fill down
        if (name) {
          lastValidName = name;
          // Reset other fields if name changes to avoid bleeding across students
          lastValidNumber = number || '';
          lastValidType = type || '';
          lastValidReason = reason || '';
        } else {
          name = lastValidName;
          if (!number) number = lastValidNumber;
          if (!type) type = lastValidType;
          if (!reason) reason = lastValidReason;
        }
        
        if (number) lastValidNumber = number;
        if (type) lastValidType = type;
        if (reason) lastValidReason = reason;

        // Handle date
        let rawDate = row['일자'] || row['날짜'] || row['일시'] || row['결석일'] || row['결석일자'] || row['기간'] || '';
        let dateStr = '';
        
        if (rawDate instanceof Date) {
          const y = rawDate.getFullYear();
          const m = String(rawDate.getMonth() + 1).padStart(2, '0');
          const d = String(rawDate.getDate()).padStart(2, '0');
          dateStr = `${y}.${m}.${d}`;
        } else if (typeof rawDate === 'number') {
          try {
            const d = XLSX.SSF.parse_date_code(rawDate);
            dateStr = `${d.y}.${String(d.m).padStart(2, '0')}.${String(d.d).padStart(2, '0')}`;
          } catch (err) {
            dateStr = '';
          }
        } else {
          dateStr = String(rawDate).trim().replace(/-/g, '.');
          // Match YYYY.MM.DD or MM.DD or YYYYMMDD
          if (/^\d{8}$/.test(dateStr)) {
            dateStr = `${dateStr.substring(0, 4)}.${dateStr.substring(4, 6)}.${dateStr.substring(6, 8)}`;
          } else if (/^\d{4}\.\d{2}\.\d{2}/.test(dateStr)) {
            dateStr = dateStr.substring(0, 10);
          }
        }

        const reportType = normalizeReportType(type);
        if (!reportType) return null;

        return {
          id: `record-${index}`,
          date: dateStr,
          number,
          name,
          type,
          reportType,
          reason,
          year: config.year,
          class: config.class,
          teacherName: config.teacherName
        };
      }).filter((r): r is AttendanceRecord => {
        return !!(r && r.name && r.date);
      });

      setRecords(parsedRecords);
      generateReports(parsedRecords);
    };
    reader.readAsArrayBuffer(file);
  };

  const isHolidayOrWeekend = (date: Date): boolean => {
    const day = date.getDay();
    if (day === 0 || day === 6) return true;

    return !!koreanHolidays.isHoliday(date);
  };

  const formatDateWithDay = (dateStr: string) => {
    if (!dateStr) return '';
    const d = new Date(dateStr.replace(/\./g, '-'));
    if (isNaN(d.getTime())) return dateStr;
    const days = ['일', '월', '화', '수', '목', '금', '토'];
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const dayName = days[d.getDay()];
    return `${y}. ${m}. ${day}.(${dayName})`;
  };

  const generateReports = (data: AttendanceRecord[]) => {
    const grouped: GroupedReport[] = [];
    
    // Sort by name then date
    const sorted = [...data].sort((a, b) => {
      if (a.name !== b.name) return a.name.localeCompare(b.name);
      return a.date.localeCompare(b.date);
    });

    let currentGroup: GroupedReport | null = null;

    const isConsecutive = (date1: string, date2: string): boolean => {
      if (date1 === date2) return true;
      const d1 = new Date(date1.replace(/\./g, '-'));
      const d2 = new Date(date2.replace(/\./g, '-'));
      if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return false;
      
      const diff = Math.abs(d2.getTime() - d1.getTime());
      const diffDays = Math.round(diff / (1000 * 60 * 60 * 24));
      
      if (diffDays <= 1) return true;
      
      // Check if all days between d1 and d2 are holidays or weekends
      let tempDate = new Date(d1);
      tempDate.setDate(tempDate.getDate() + 1);
      while (tempDate < d2) {
        if (!isHolidayOrWeekend(tempDate)) return false;
        tempDate.setDate(tempDate.getDate() + 1);
      }
      return true;
    };

    sorted.forEach((record) => {
      const type = record.reportType;
      const schoolYear = getSchoolYear(record.date);
      const normalizedReason = (record.reason || '').trim();
      const effectiveReason = normalizedReason || getDefaultReason(type);

      if (currentGroup && 
          currentGroup.name === record.name && 
          currentGroup.number === record.number &&
          currentGroup.schoolYear === schoolYear &&
          currentGroup.type === type &&
          currentGroup.reason === effectiveReason &&
          isConsecutive(currentGroup.endDate, record.date)) {

        if (currentGroup.endDate !== record.date) {
          currentGroup.endDate = record.date;
          currentGroup.days += 1;
        }
        currentGroup.records.push(record);
      } else {
        if (currentGroup) grouped.push(currentGroup);
        currentGroup = {
          id: `report-${grouped.length}`,
          name: record.name,
          number: record.number,
          schoolYear,
          year: record.year,
          class: record.class,
          type: type,
          reason: effectiveReason,
          startDate: record.date,
          endDate: record.date,
          days: 1,
          records: [record],
          teacherName: record.teacherName,
          guardianName: guardianMap[`${record.number}-${record.name.trim()}`] || guardianMap[record.name.trim()] || ''
        };
      }
    });
    if (currentGroup) grouped.push(currentGroup);
    setReports(grouped);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processExcel(e.target.files[0]);
    }
  };

  const onDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processExcel(e.dataTransfer.files[0]);
    }
  }, [config, guardianMap]);

  const getSubmissionDate = (endDateStr: string): string => {
    // Handle various formats: YYYY.MM.DD or YYYY-MM-DD
    const normalized = endDateStr.replace(/-/g, '.');
    const parts = normalized.split('.');
    
    let date: Date;
    if (parts.length === 3) {
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const day = parseInt(parts[2]);
      date = new Date(year, month, day);
    } else {
      date = new Date(endDateStr);
    }

    if (isNaN(date.getTime())) return endDateStr;
    
    // Add 1 day
    date.setDate(date.getDate() + 1);
    
    // Skip weekends and holidays
    while (isHolidayOrWeekend(date)) {
      date.setDate(date.getDate() + 1);
    }
    
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    
    return `${y}. ${m}. ${d}.`;
  };

  const downloadReport = async (report: GroupedReport) => {
    const submissionDate = getSubmissionDate(report.endDate);
    const blob = await generateAttendanceReport({
      ...report,
      startDate: formatDateWithDay(report.startDate),
      endDate: formatDateWithDay(report.endDate),
      submissionDate
    });
    saveAs(blob, `${report.name}_${report.type}결석신고서.docx`);
  };

  const downloadAll = async () => {
    if (reports.length === 0) return;
    
    const bulkData = reports.map(report => ({
      ...report,
      startDate: formatDateWithDay(report.startDate),
      endDate: formatDateWithDay(report.endDate),
      submissionDate: getSubmissionDate(report.endDate)
    }));

    const now = new Date();
    const fileDate = `${now.getFullYear()}.${String(now.getMonth() + 1).padStart(2, '0')}.${String(now.getDate()).padStart(2, '0')}`;
    const blob = await generateBulkAttendanceReport(bulkData);
    saveAs(blob, `전체_결석신고서_${fileDate}.docx`);
  };

  return (
    <div className="min-h-screen bg-[#F5F5F0] text-[#1A1A1A] font-sans p-6 md:p-12">
      <div className="max-w-5xl mx-auto">
        {/* Header */}
        <header className="mb-12 flex flex-col md:flex-row md:items-end justify-between gap-6">
          <div>
            <h1 className="text-5xl font-serif font-light tracking-tight mb-2">출결 서류 생성기</h1>
            <p className="text-[#5A5A40] italic">Attendance Report Automator</p>
          </div>
          <div className="flex gap-4 bg-white p-4 rounded-3xl shadow-sm border border-black/5">
            <div className="flex flex-col">
              <label className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">학년/반</label>
              <div className="flex items-center gap-2">
                <input 
                  type="text" 
                  value={config.year} 
                  onChange={(e) => updateYear(e.target.value)}
                  className="w-8 text-center border-b border-black/20 focus:border-black outline-none"
                />
                <span>-</span>
                <input 
                  type="text" 
                  value={config.class} 
                  onChange={(e) => updateClass(e.target.value)}
                  className="w-8 text-center border-b border-black/20 focus:border-black outline-none"
                />
              </div>
            </div>
            <div className="w-px bg-black/5 mx-2" />
            <div className="flex flex-col">
              <label className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">담임교사</label>
              <input 
                type="text" 
                value={config.teacherName} 
                onChange={(e) => updateTeacherName(e.target.value)}
                className="w-24 border-b border-black/20 focus:border-black outline-none"
              />
            </div>
          </div>
        </header>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-12">
          {/* Attendance Excel Upload */}
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className={`relative group cursor-pointer h-full`}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={onDrop}
          >
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className={`
              border-2 border-dashed rounded-[40px] p-12 flex flex-col items-center justify-center transition-all duration-500 h-full
              ${isDragging ? 'border-[#5A5A40] bg-[#5A5A40]/5 scale-[0.99]' : 'border-black/10 bg-white hover:border-black/20'}
              ${records.length > 0 ? 'border-emerald-200 bg-emerald-50/30' : ''}
            `}>
              <div className={`w-16 h-16 rounded-full flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-500 ${records.length > 0 ? 'bg-emerald-100' : 'bg-[#F5F5F0]'}`}>
                {records.length > 0 ? <CheckCircle2 className="w-7 h-7 text-emerald-600" /> : <Upload className="w-7 h-7 text-[#5A5A40]" />}
              </div>
              <h2 className="text-xl font-serif mb-1">출결 현황 엑셀</h2>
              <p className="text-[#8E9299] text-xs text-center">
                {records.length > 0 ? '파일이 업로드되었습니다' : '출결 데이터를 업로드하세요'}
              </p>
            </div>
          </motion.div>

          {/* Guardian Map Upload */}
          <motion.div 
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ delay: 0.1 }}
            className={`relative group cursor-pointer h-full`}
            onDragOver={(e) => { e.preventDefault(); setIsDraggingMap(true); }}
            onDragLeave={() => setIsDraggingMap(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDraggingMap(false);
              if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                processGuardianMap(e.dataTransfer.files[0]);
              }
            }}
          >
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              onChange={(e) => e.target.files?.[0] && processGuardianMap(e.target.files[0])}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className={`
              border-2 border-dashed rounded-[40px] p-12 flex flex-col items-center justify-center transition-all duration-500 h-full
              ${isDraggingMap ? 'border-[#5A5A40] bg-[#5A5A40]/5 scale-[0.99]' : 'border-black/10 bg-white hover:border-black/20'}
              ${Object.keys(guardianMap).length > 0 ? 'border-emerald-200 bg-emerald-50/30' : ''}
            `}>
              <div className={`w-16 h-16 rounded-full flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-500 ${Object.keys(guardianMap).length > 0 ? 'bg-emerald-100' : 'bg-[#F5F5F0]'}`}>
                {Object.keys(guardianMap).length > 0 ? <CheckCircle2 className="w-7 h-7 text-emerald-600" /> : <FileSpreadsheet className="w-7 h-7 text-[#5A5A40]" />}
              </div>
              <h2 className="text-xl font-serif mb-1">학생-보호자 명단</h2>
              <p className="text-[#8E9299] text-xs text-center">
                {Object.keys(guardianMap).length > 0 ? `${Object.keys(guardianMap).length}명의 보호자 정보 로드됨` : '학생/보호자 성함 파일을 업로드하세요'}
              </p>
            </div>
          </motion.div>
        </div>

        {records.length > 0 && (
          <div className="space-y-8">
            {/* Action Bar */}
            <div className="flex items-center justify-between bg-white p-6 rounded-[32px] shadow-sm border border-black/5">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center">
                  <CheckCircle2 className="w-6 h-6 text-emerald-600" />
                </div>
                <div>
                  <h3 className="font-medium">{reports.length}개의 서류가 생성되었습니다</h3>
                  <p className="text-xs text-[#8E9299]">내용을 확인하고 다운로드하세요</p>
                </div>
              </div>
              <div className="flex gap-3">
                <button 
                  onClick={() => { setRecords([]); setReports([]); }}
                  className="px-6 py-3 rounded-full border border-black/10 text-sm hover:bg-black/5 transition-colors flex items-center gap-2"
                >
                  <Trash2 className="w-4 h-4" /> 다시 업로드
                </button>
                <button 
                  onClick={downloadAll}
                  className="px-8 py-3 rounded-full bg-[#5A5A40] text-white text-sm hover:bg-[#4A4A30] transition-colors flex items-center gap-2 shadow-lg shadow-[#5A5A40]/20"
                >
                  <Download className="w-4 h-4" /> 전체 다운로드
                </button>
              </div>
            </div>

            {/* Reports List */}
            <div className="grid gap-4">
              <AnimatePresence mode="popLayout">
                {reports.map((report, idx) => (
                  <motion.div
                    key={report.id}
                    layout
                    initial={{ opacity: 0, x: -20 }}
                    animate={{ opacity: 1, x: 0 }}
                    exit={{ opacity: 0, scale: 0.95 }}
                    transition={{ delay: idx * 0.05 }}
                    className="group bg-white p-6 rounded-[32px] border border-black/5 hover:border-black/20 transition-all flex items-center gap-6"
                  >
                    <div className={`w-14 h-14 rounded-2xl flex items-center justify-center shrink-0 ${
                      report.type === '질병' ? 'bg-amber-50' : 
                      report.type === '출석인정' ? 'bg-blue-50' : 
                      report.type === '미인정' ? 'bg-rose-50' : 'bg-slate-50'
                    }`}>
                      <FileText className={`w-7 h-7 ${
                        report.type === '질병' ? 'text-amber-600' : 
                        report.type === '출석인정' ? 'text-blue-600' : 
                        report.type === '미인정' ? 'text-rose-600' : 'text-slate-600'
                      }`} />
                    </div>
                    
                    <div className="flex-1 grid grid-cols-1 md:grid-cols-4 gap-4 items-center">
                      <div>
                        <p className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">학생 / 보호자</p>
                        <div className="flex items-center gap-1">
                          <input 
                            type="text" 
                            value={report.name}
                            onChange={(e) => {
                              const newReports = [...reports];
                              newReports[idx].name = e.target.value;
                              setReports(newReports);
                            }}
                            className="font-medium text-lg border-b border-transparent hover:border-black/20 focus:border-black outline-none bg-transparent w-24"
                          />
                          <span className="text-sm font-normal text-[#8E9299] flex items-center">
                            (<input 
                              type="text" 
                              value={report.number}
                              onChange={(e) => {
                                const newReports = [...reports];
                                newReports[idx].number = e.target.value;
                                setReports(newReports);
                              }}
                              className="w-6 text-center border-b border-transparent hover:border-black/20 focus:border-black outline-none bg-transparent"
                            />번)
                          </span>
                        </div>
                        <div className="flex items-center gap-1">
                          <span className="text-xs text-[#5A5A40] italic">보호자:</span>
                          <input 
                            type="text" 
                            value={report.guardianName}
                            placeholder="보호자 성함"
                            onChange={(e) => {
                              const newReports = [...reports];
                              newReports[idx].guardianName = e.target.value;
                              setReports(newReports);
                            }}
                            className="text-xs text-[#5A5A40] italic border-b border-transparent hover:border-black/20 focus:border-black outline-none bg-transparent w-24"
                          />
                        </div>
                      </div>
                      <div>
                        <p className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">구분 / 사유</p>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-0.5 rounded-md text-[10px] font-bold ${
                            report.type === '질병' ? 'bg-amber-100 text-amber-700' : 
                            report.type === '출석인정' ? 'bg-blue-100 text-blue-700' : 
                            report.type === '미인정' ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-700'
                          }`}>
                            {report.type}
                          </span>
                          <input 
                            type="text" 
                            value={report.reason}
                            onChange={(e) => {
                              const newReports = [...reports];
                              newReports[idx].reason = e.target.value;
                              setReports(newReports);
                            }}
                            className="text-sm border-b border-transparent hover:border-black/20 focus:border-black outline-none bg-transparent"
                          />
                        </div>
                      </div>
                      <div>
                        <p className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">기간 / 제출예정일</p>
                        <p className="text-sm">{report.startDate} ~ {report.endDate}</p>
                        <p className="text-[10px] font-bold text-[#5A5A40] mt-1">제출일: {getSubmissionDate(report.endDate)}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-[10px] uppercase tracking-widest text-[#8E9299] mb-1">일수</p>
                        <p className="font-serif italic text-xl">{report.days}일간</p>
                      </div>
                    </div>

                    <button 
                      onClick={() => downloadReport(report)}
                      className="w-12 h-12 rounded-full border border-black/5 flex items-center justify-center hover:bg-black hover:text-white transition-all group-hover:scale-110"
                    >
                      <Download className="w-5 h-5" />
                    </button>
                  </motion.div>
                ))}
              </AnimatePresence>
            </div>
          </div>
        )}

        {/* Footer Info */}
        <footer className="mt-20 pt-8 border-t border-black/5 flex justify-between items-center text-[#8E9299] text-xs">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1">
              <FileSpreadsheet className="w-3 h-3" />
              <span>XLSX Parser Ready</span>
            </div>
            <div className="flex items-center gap-1">
              <FileText className="w-3 h-3" />
              <span>DOCX Generator Ready</span>
            </div>
          </div>
          <p>© 2026 한올고등학교 출결 시스템 지원 도구</p>
        </footer>
      </div>
    </div>
  );
}
