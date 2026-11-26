import React, { useState, useMemo, useEffect } from 'react';
import {
  Home,
  HelpCircle,
  Trash2,
  Save,
  FileSpreadsheet,
  Upload,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp } from '../../context/AppContext';
import { Attendance, Student } from '../../types';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentRecord } from '../../utils/studentSyncHelper';

const MAX_SHEET_ROWS = 60;

type AttendanceField = 'sakit' | 'izin' | 'alpa';

interface DisplayValues {
  sakit: string;
  izin: string;
  alpa: string;
}

function parseAttendanceInput(raw: string): number {
  const trimmed = raw.trim();
  if (!trimmed || trimmed === '-') return 0;
  const parsed = parseInt(trimmed.replace(/[^\d]/g, ''), 10);
  return Number.isNaN(parsed) ? 0 : Math.max(0, parsed);
}

export const AttendanceSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    saveKehadiran,
    bulkSaveKehadiran,
    showSaveNotification,
    setActiveMenu,
    openClearDataModal
  } = appContext;

  const [showHelp, setShowHelp] = useState(false);

  const siswaKelas = useMemo(
    () => db.siswa.filter((s) => s.kelas === selectedKelas),
    [db.siswa, selectedKelas]
  );

  const displayedSiswa = useMemo(
    () => (filteredStudent ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa) : siswaKelas),
    [filteredStudent, siswaKelas]
  );

  const totalRowSlots = filteredStudent
    ? Math.max(displayedSiswa.length, 1)
    : Math.max(MAX_SHEET_ROWS, displayedSiswa.length);

  const buildAttendanceMaps = () => {
    const map: Record<string, Attendance> = {};
    const dispMap: Record<string, DisplayValues> = {};
    siswaKelas.forEach((s) => {
      const existing = db.kehadiran.find(
        (h) => matchesStudentRecord(h, s) && h.semester === selectedSemester
      );
      if (existing) {
        map[s.idSiswa] = {
          ...existing,
          idSiswa: s.idSiswa,
          nisn: s.nisn,
          namaSiswa: s.namaSiswa,
          kelas: selectedKelas
        };
        dispMap[s.idSiswa] = {
          sakit: String(existing.sakit ?? 0),
          izin: String(existing.izin ?? 0),
          alpa: String(existing.alpa ?? 0)
        };
      } else {
        map[s.idSiswa] = {
          id: `H_${s.idSiswa}_${selectedSemester}`,
          idSiswa: s.idSiswa,
          nisn: s.nisn,
          namaSiswa: s.namaSiswa,
          kelas: selectedKelas,
          semester: selectedSemester,
          tahunPelajaran: db.sekolah.tahunPelajaran,
          sakit: 0,
          izin: 0,
          alpa: 0,
          totalHari: 0
        };
        dispMap[s.idSiswa] = {
          sakit: '0',
          izin: '0',
          alpa: '0'
        };
      }
    });
    return { map, dispMap };
  };

  const initialMaps = buildAttendanceMaps();
  const [attendances, setAttendances] = useState<Record<string, Attendance>>(initialMaps.map);
  const [displayInputs, setDisplayInputs] = useState<Record<string, DisplayValues>>(initialMaps.dispMap);

  useEffect(() => {
    const next = buildAttendanceMaps();
    setAttendances(next.map);
    setDisplayInputs((prev) => {
      const merged: Record<string, DisplayValues> = { ...next.dispMap };
      Object.keys(merged).forEach((id) => {
        if (prev[id]) {
          // Preserve '-' display when numeric value is 0
          (['sakit', 'izin', 'alpa'] as AttendanceField[]).forEach((field) => {
            if (prev[id][field] === '-' && next.map[id][field] === 0) {
              merged[id][field] = '-';
            }
          });
        }
      });
      return merged;
    });
  }, [db.siswa, db.kehadiran, selectedKelas, selectedSemester]);

  const buildFallbackAtt = (s: Student): Attendance => ({
    id: `H_${s.idSiswa}_${selectedSemester}`,
    idSiswa: s.idSiswa,
    nisn: s.nisn,
    namaSiswa: s.namaSiswa,
    kelas: selectedKelas,
    semester: selectedSemester,
    tahunPelajaran: db.sekolah.tahunPelajaran,
    sakit: 0,
    izin: 0,
    alpa: 0,
    totalHari: 0
  });

  const handleCellChange = (student: Student, field: AttendanceField, rawValue: string) => {
    // Allow digits or single '-' or empty while editing
    const sanitized = rawValue === '-' ? '-' : rawValue.replace(/[^\d]/g, '');
    setDisplayInputs((prev) => ({
      ...prev,
      [student.idSiswa]: {
        ...(prev[student.idSiswa] || { sakit: '0', izin: '0', alpa: '0' }),
        [field]: sanitized
      }
    }));

    const numericValue = parseAttendanceInput(sanitized);
    const current = attendances[student.idSiswa] || buildFallbackAtt(student);
    const updated: Attendance = {
      ...current,
      idSiswa: student.idSiswa,
      nisn: student.nisn,
      namaSiswa: student.namaSiswa,
      kelas: selectedKelas,
      [field]: numericValue
    };
    updated.totalHari = Number(updated.sakit) + Number(updated.izin) + Number(updated.alpa);

    setAttendances((prev) => ({
      ...prev,
      [student.idSiswa]: updated
    }));
    saveKehadiran(updated, true);
  };

  const handleCellBlur = (student: Student, field: AttendanceField) => {
    const currentDisp = displayInputs[student.idSiswa]?.[field] ?? '';
    if (currentDisp.trim() === '') {
      setDisplayInputs((prev) => ({
        ...prev,
        [student.idSiswa]: {
          ...(prev[student.idSiswa] || { sakit: '0', izin: '0', alpa: '0' }),
          [field]: '0'
        }
      }));
    }
  };

  const handleSaveAll = () => {
    const list = siswaKelas.map((s) => attendances[s.idSiswa] || buildFallbackAtt(s));
    bulkSaveKehadiran(list);
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['ABSENSI SISWA (KETIDAKHADIRAN)'],
      [`Kelas: ${selectedKelas} • Semester: ${selectedSemester} • Tahun Ajaran: ${db.sekolah.tahunPelajaran}`],
      [],
      ['NO', 'NISN', 'NAMA SISWA', 'SAKIT', 'IZIN', 'TANPA KETERANGAN', 'TOTAL ABSEN'],
      ...siswaKelas.map((s, idx) => {
        const item = attendances[s.idSiswa] || { sakit: 0, izin: 0, alpa: 0, totalHari: 0 };
        return [
          idx + 1,
          s.nisn,
          s.namaSiswa,
          item.sakit,
          item.izin,
          item.alpa,
          item.totalHari
        ];
      })
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 18 },
      { wch: 30 },
      { wch: 12 },
      { wch: 12 },
      { wch: 18 },
      { wch: 14 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Absensi Siswa');
    XLSX.writeFile(wb, `absensi-siswa-kelas-${selectedKelas}-${selectedSemester}.xlsx`);
    showSaveNotification('Data Absensi Siswa berhasil diekspor ke Excel (.xlsx)!');
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const result = await parseAndImportTemplate(file, 'template_kehadiran', appContext);
      showSaveNotification(
        result.success ? result.message : `Gagal impor: ${result.message}`
      );
    } catch (err: any) {
      showSaveNotification(`Gagal mengimpor file: ${err?.message || 'Format tidak sesuai'}`);
    }
    e.target.value = '';
  };

  return (
    <div className="bg-[#00ffff] min-h-screen p-2 sm:p-3 rounded-xl shadow-inner">
      {/* TOP TOOLBAR STRIP */}
      <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
        <div className="flex flex-wrap items-center gap-6 sm:gap-10">
          {/* Black Box with Home, Help, Delete (matching screenshot) */}
          <div className="bg-black border-2 border-red-600 px-2.5 py-1 flex items-center gap-2 shadow">
            <button
              onClick={() => setActiveMenu('dashboard')}
              className="w-8 h-8 rounded-full bg-sky-600 hover:bg-sky-500 text-white border border-sky-300 flex items-center justify-center cursor-pointer transition-transform active:scale-95"
              title="Kembali ke Dashboard"
            >
              <Home size={16} />
            </button>
            <button
              onClick={() => setShowHelp((prev) => !prev)}
              className="w-8 h-8 rounded-full bg-purple-700 hover:bg-purple-600 text-white border border-purple-300 flex items-center justify-center cursor-pointer"
              title="Panduan Pengisian Absensi Siswa"
            >
              <HelpCircle size={17} />
            </button>
            <button
              onClick={() => openClearDataModal()}
              className="w-8 h-8 rounded-full bg-red-600 hover:bg-red-500 text-white border border-red-300 flex items-center justify-center cursor-pointer"
              title="Bersihkan Semua Data Absensi"
            >
              <Trash2 size={15} />
            </button>
          </div>

          <h1 className="text-lg sm:text-xl font-extrabold tracking-[0.15em] uppercase text-black font-serif">
            ABSENSI SISWA
          </h1>
        </div>

        {/* Right Action Controls */}
        <div className="flex flex-wrap items-center gap-1.5">
          <div className="bg-white rounded border border-black px-1 py-0.5">
            <StudentFilterSelect />
          </div>

          <button
            onClick={handleSaveAll}
            className="px-2.5 py-1.5 bg-sky-700 hover:bg-sky-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Simpan & Hubungkan semua data Absensi Siswa ke Rapor"
          >
            <Save size={13} />
            <span>Simpan</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Ekspor Absensi Siswa ke Excel (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>Excel</span>
          </button>

          <label
            className="px-2.5 py-1.5 bg-indigo-700 hover:bg-indigo-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Impor Data Kehadiran dari Excel"
          >
            <Upload size={13} />
            <span>Impor</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} className="hidden" />
          </label>
        </div>
      </div>

      {/* HELP PANEL */}
      {showHelp && (
        <div className="mb-2 bg-white border-2 border-black p-3 rounded shadow-lg text-xs text-slate-800 max-w-3xl">
          <div className="flex items-center justify-between border-b border-slate-200 pb-1 mb-1.5">
            <h3 className="font-bold text-slate-900">
              Panduan Pengisian Absensi Siswa (Kelas {selectedKelas} • Semester {selectedSemester})
            </h3>
            <button
              onClick={() => setShowHelp(false)}
              className="p-0.5 hover:bg-slate-100 rounded text-slate-600 cursor-pointer"
              aria-label="Tutup Panduan"
            >
              <X size={15} />
            </button>
          </div>
          <ul className="list-disc pl-5 space-y-1">
            <li>
              <strong>Pengisian Ketidakhadiran:</strong> Ketik jumlah hari <em>Sakit</em>, <em>Izin</em>, dan <em>Tanpa Keterangan</em> pada kolom siswa. Tulis <strong>0 (nol)</strong> atau tanda strip (<strong>-</strong>) jika siswa tidak pernah absen.
            </li>
            <li>
              <strong>Terhubung ke Rapor:</strong> Data absensi ini otomatis masuk ke bagian <em>C. KETIDAKHADIRAN</em> pada Rapor Akademik, Buku Induk, dan Cetak Rapor.
            </li>
          </ul>
        </div>
      )}

      {/* MAIN TABLE + RIGHT INSTRUCTION BOX */}
      <div className="flex flex-col lg:flex-row items-start gap-5">
        {/* SPREADSHEET TABLE */}
        <div className="bg-white border-2 border-black shadow-xl overflow-x-auto w-full lg:w-auto">
          <table className="border-collapse text-xs text-black">
            <thead>
              {/* Row 1 */}
              <tr>
                <th
                  rowSpan={2}
                  className="border border-black bg-[#76ff03] text-black font-bold italic text-center px-2 py-1.5 w-11"
                >
                  No.
                </th>
                <th
                  rowSpan={2}
                  className="border border-black bg-[#76ff03] text-black font-bold italic text-center px-3 py-1.5 min-w-[210px] w-[240px]"
                >
                  Nama
                </th>
                <th
                  colSpan={3}
                  className="border border-black bg-[#4f81bd] text-black font-extrabold uppercase text-center py-1 px-3 text-[11px] tracking-wide"
                >
                  KETIDAKHADIRAN
                </th>
              </tr>

              {/* Row 2 */}
              <tr>
                <th className="border border-black bg-[#ffff00] text-black font-bold text-center px-2 py-1.5 w-28">
                  Sakit
                </th>
                <th className="border border-black bg-[#76ff03] text-black font-bold text-center px-2 py-1.5 w-28">
                  Izin
                </th>
                <th className="border border-black bg-[#ff0000] text-black font-bold text-center px-2 py-1.5 w-32 leading-tight">
                  Tanpa
                  <br />
                  Keterangan
                </th>
              </tr>
            </thead>

            <tbody>
              {Array.from({ length: totalRowSlots }).map((_, rowIdx) => {
                const student = displayedSiswa[rowIdx];
                const disp = student
                  ? displayInputs[student.idSiswa] || { sakit: '0', izin: '0', alpa: '0' }
                  : null;
                const rowNum = filteredStudent ? 1 : rowIdx + 1;

                return (
                  <tr
                    key={student ? student.idSiswa : `empty-att-row-${rowIdx}`}
                    className="hover:bg-sky-50/40"
                  >
                    <td className="border border-black bg-white text-center font-medium py-0.5 px-2 h-[25px]">
                      {rowNum}
                    </td>
                    <td className="border border-black bg-white px-2.5 py-0.5 font-semibold text-left whitespace-nowrap">
                      {student ? student.namaSiswa : ''}
                    </td>

                    {(['sakit', 'izin', 'alpa'] as AttendanceField[]).map((field) => (
                      <td key={`${rowIdx}-${field}`} className="border border-black bg-white p-0">
                        {student && disp ? (
                          <input
                            type="text"
                            inputMode="numeric"
                            value={disp[field]}
                            onChange={(e) => handleCellChange(student, field, e.target.value)}
                            onBlur={() => handleCellBlur(student, field)}
                            className="w-full h-full px-2 py-0.5 text-center font-semibold text-xs text-slate-900 bg-transparent focus:bg-yellow-50 focus:outline-none font-mono"
                            aria-label={`${field} untuk ${student.namaSiswa}`}
                          />
                        ) : (
                          <div className="h-[23px]" />
                        )}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* CALLOUT BOX ON RIGHT SIDE (matching screenshot) */}
        <div className="bg-white border-2 border-slate-700 shadow-md px-3.5 py-2.5 text-xs text-slate-900 font-medium max-w-[230px] lg:mt-14">
          tulis 0 (nol) atau tanda strip (-) jika tidak ada absen.
        </div>
      </div>
    </div>
  );
};
