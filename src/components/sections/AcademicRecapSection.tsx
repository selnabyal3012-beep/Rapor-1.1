import React, { useMemo, useState, useEffect } from 'react';
import {
  Home,
  HelpCircle,
  Printer,
  FileSpreadsheet,
  Download,
  Upload,
  Eye,
  BookOpen,
  X,
  Check
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp } from '../../context/AppContext';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { exportClassComprehensiveWorkbook, exportClassLegerToCsv } from '../../utils/exportHelper';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentScore } from '../../utils/studentSyncHelper';
import { DEFAULT_KKM_SETTINGS } from '../../data/defaultData';
import { AcademicScore, Student } from '../../types';

interface LegerSubjectColumn {
  id: string;
  label: string;
  shortLabel: string;
  group?: 'seni' | 'mulok';
  headerBg: string;
  headerText: string;
  barColor: string;
  matchNames: string[];
  isAgama?: boolean;
  defaultKkm: number;
  visible: boolean;
}

const MAX_SHEET_ROWS = 60;

export const AcademicRecapSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    setActiveMenu,
    setPrintTargetStudentId,
    openReportPreview,
    updateSekolah,
    showSaveNotification
  } = appContext;

  const kkmSettings = db.kkmSettings || DEFAULT_KKM_SETTINGS;

  // Toggle visibility for Seni Pilihan (1..4) and Muatan Lokal (1..5), synced initially with kkmSettings
  const [seniVisible, setSeniVisible] = useState<Record<'seniMusik' | 'seniTari' | 'seniRupa' | 'seniTeater', boolean>>({
    seniMusik: kkmSettings.seniPilihanAktif?.seniMusik ?? true,
    seniTari: kkmSettings.seniPilihanAktif?.seniTari ?? true,
    seniRupa: kkmSettings.seniPilihanAktif?.seniRupa ?? true,
    seniTeater: kkmSettings.seniPilihanAktif?.seniTeater ?? true
  });

  const [mulokVisible, setMulokVisible] = useState<boolean[]>([
    (kkmSettings.muatanLokalCount ?? 5) >= 1,
    (kkmSettings.muatanLokalCount ?? 5) >= 2,
    (kkmSettings.muatanLokalCount ?? 5) >= 3,
    (kkmSettings.muatanLokalCount ?? 5) >= 4,
    (kkmSettings.muatanLokalCount ?? 5) >= 5
  ]);

  useEffect(() => {
    setSeniVisible({
      seniMusik: kkmSettings.seniPilihanAktif?.seniMusik ?? true,
      seniTari: kkmSettings.seniPilihanAktif?.seniTari ?? true,
      seniRupa: kkmSettings.seniPilihanAktif?.seniRupa ?? true,
      seniTeater: kkmSettings.seniPilihanAktif?.seniTeater ?? true
    });
    const cnt = kkmSettings.muatanLokalCount ?? 5;
    setMulokVisible([cnt >= 1, cnt >= 2, cnt >= 3, cnt >= 4, cnt >= 5]);
  }, [kkmSettings]);

  const [showHelp, setShowHelp] = useState(false);
  const [pemutihanGrafik, setPemutihanGrafik] = useState(false);
  const [tempatTanggalInput, setTempatTanggalInput] = useState(
    db.sekolah.tempatTanggalRapor ||
      `${db.sekolah.tempatRapor || 'Oksibil'}, ${db.sekolah.tanggalRapor || '10 Juli 2026'}`
  );
  const [pengesahanOleh, setPengesahanOleh] = useState('Kepala Sekolah');

  useEffect(() => {
    setTempatTanggalInput(
      db.sekolah.tempatTanggalRapor ||
        `${db.sekolah.tempatRapor || 'Oksibil'}, ${db.sekolah.tanggalRapor || '10 Juli 2026'}`
    );
  }, [db.sekolah.tempatTanggalRapor, db.sekolah.tempatRapor, db.sekolah.tanggalRapor]);

  const siswaKelas = useMemo(
    () => db.siswa.filter((s) => s.kelas === selectedKelas),
    [db.siswa, selectedKelas]
  );

  const displayedStudents = useMemo(
    () => (filteredStudent ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa) : siswaKelas),
    [filteredStudent, siswaKelas]
  );

  // Check if IPAS exists in db.mataPelajaran (used in Fase B/C classes III-VI)
  const hasIpas = useMemo(
    () =>
      db.mataPelajaran.some(
        (m) =>
          m.mataPelajaran.toLowerCase().includes('ipas') ||
          m.mataPelajaran.toLowerCase().includes('ilmu pengetahuan alam dan sosial')
      ),
    [db.mataPelajaran]
  );

  // Build ordered columns matching screenshot + Kurikulum Merdeka
  const allColumns = useMemo<LegerSubjectColumn[]>(() => {
    const mulokNames = kkmSettings.muatanLokalNames || [];
    const cols: LegerSubjectColumn[] = [
      {
        id: 'agama',
        label: 'Pendidikan Agama dan Budi Pekerti',
        shortLabel: 'Pendidikan Agama dan Budi Pekerti',
        headerBg: '#fff9c4',
        headerText: '#0f172a',
        barColor: '#eab308',
        matchNames: ['pendidikan agama dan budi pekerti'],
        isAgama: true,
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'pancasila',
        label: 'Pendidikan Pancasila',
        shortLabel: 'Pendidikan Pancasila',
        headerBg: '#76ff03',
        headerText: '#0f172a',
        barColor: '#22c55e',
        matchNames: ['pendidikan pancasila', 'ppkn'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'bindo',
        label: 'Bahasa Indonesia',
        shortLabel: 'Bahasa Indonesia',
        headerBg: '#ff66ff',
        headerText: '#0f172a',
        barColor: '#d946ef',
        matchNames: ['bahasa indonesia'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'mtk',
        label: 'Matematika',
        shortLabel: 'Matematika',
        headerBg: '#00ffff',
        headerText: '#0f172a',
        barColor: '#06b6d4',
        matchNames: ['matematika'],
        defaultKkm: 75,
        visible: true
      },
      ...(hasIpas
        ? [
            {
              id: 'ipas',
              label: 'Ilmu Pengetahuan Alam dan Sosial (IPAS)',
              shortLabel: 'IPAS',
              headerBg: '#bef264',
              headerText: '#0f172a',
              barColor: '#65a30d',
              matchNames: ['ilmu pengetahuan alam dan sosial (ipas)', 'ipas'],
              defaultKkm: 75,
              visible: true
            }
          ]
        : []),
      {
        id: 'seni_musik',
        label: 'Seni Musik',
        shortLabel: 'Seni Musik',
        group: 'seni',
        headerBg: '#ffff00',
        headerText: '#0f172a',
        barColor: '#facc15',
        matchNames: ['seni musik'],
        defaultKkm: 75,
        visible: seniVisible.seniMusik
      },
      {
        id: 'seni_tari',
        label: 'Seni Tari',
        shortLabel: 'Seni Tari',
        group: 'seni',
        headerBg: '#e9d5ff',
        headerText: '#0f172a',
        barColor: '#a855f7',
        matchNames: ['seni tari'],
        defaultKkm: 75,
        visible: seniVisible.seniTari
      },
      {
        id: 'seni_rupa',
        label: 'Seni Rupa',
        shortLabel: 'Seni Rupa',
        group: 'seni',
        headerBg: '#0284c7',
        headerText: '#ffffff',
        barColor: '#0284c7',
        matchNames: ['seni rupa', 'seni dan budaya (seni rupa & musik)'],
        defaultKkm: 75,
        visible: seniVisible.seniRupa
      },
      {
        id: 'seni_teater',
        label: 'Seni Teater',
        shortLabel: 'Seni Teater',
        group: 'seni',
        headerBg: '#f59e0b',
        headerText: '#0f172a',
        barColor: '#ea580c',
        matchNames: ['seni teater'],
        defaultKkm: 75,
        visible: seniVisible.seniTeater
      },
      {
        id: 'pjok',
        label: 'Pendidikan Jasmani, Olahraga dan Kesehatan',
        shortLabel: 'Pendidikan Jasmani, Olahraga dan Kesehatan',
        headerBg: '#d8b4fe',
        headerText: '#0f172a',
        barColor: '#8b5cf6',
        matchNames: [
          'pendidikan jasmani, olahraga, dan kesehatan',
          'pendidikan jasmani, olahraga dan kesehatan',
          'pjok'
        ],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'bing',
        label: 'Bahasa Inggris',
        shortLabel: 'Bahasa Inggris',
        headerBg: '#fed7aa',
        headerText: '#0f172a',
        barColor: '#fb923c',
        matchNames: ['bahasa inggris'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'mulok_1',
        label: mulokNames[0]?.trim() || 'Mulok 1',
        shortLabel: mulokNames[0]?.trim() || 'Mulok 1',
        group: 'mulok',
        headerBg: '#cbd5e1',
        headerText: '#0f172a',
        barColor: '#64748b',
        matchNames: [
          (mulokNames[0] || '').trim().toLowerCase(),
          'muatan lokal (bahasa daerah & noken papua)',
          'muatan lokal 1',
          'mulok 1'
        ].filter(Boolean),
        defaultKkm: 75,
        visible: mulokVisible[0]
      },
      {
        id: 'mulok_2',
        label: mulokNames[1]?.trim() || 'Mulok 2',
        shortLabel: mulokNames[1]?.trim() || 'Mulok 2',
        group: 'mulok',
        headerBg: '#94a3b8',
        headerText: '#0f172a',
        barColor: '#475569',
        matchNames: [(mulokNames[1] || '').trim().toLowerCase(), 'muatan lokal 2', 'mulok 2'].filter(Boolean),
        defaultKkm: 75,
        visible: mulokVisible[1]
      },
      {
        id: 'mulok_3',
        label: mulokNames[2]?.trim() || 'Mulok 3',
        shortLabel: mulokNames[2]?.trim() || 'Mulok 3',
        group: 'mulok',
        headerBg: '#00ffff',
        headerText: '#0f172a',
        barColor: '#0891b2',
        matchNames: [(mulokNames[2] || '').trim().toLowerCase(), 'muatan lokal 3', 'mulok 3'].filter(Boolean),
        defaultKkm: 75,
        visible: mulokVisible[2]
      },
      {
        id: 'mulok_4',
        label: mulokNames[3]?.trim() || 'Mulok 4',
        shortLabel: mulokNames[3]?.trim() || 'Mulok 4',
        group: 'mulok',
        headerBg: '#84cc16',
        headerText: '#0f172a',
        barColor: '#4d7c0f',
        matchNames: [(mulokNames[3] || '').trim().toLowerCase(), 'muatan lokal 4', 'mulok 4'].filter(Boolean),
        defaultKkm: 75,
        visible: mulokVisible[3]
      },
      {
        id: 'mulok_5',
        label: mulokNames[4]?.trim() || 'Mulok 5',
        shortLabel: mulokNames[4]?.trim() || 'Mulok 5',
        group: 'mulok',
        headerBg: '#ffffff',
        headerText: '#0f172a',
        barColor: '#0ea5e9',
        matchNames: [(mulokNames[4] || '').trim().toLowerCase(), 'muatan lokal 5', 'mulok 5'].filter(Boolean),
        defaultKkm: 75,
        visible: mulokVisible[4]
      }
    ];
    return cols;
  }, [hasIpas, kkmSettings.muatanLokalNames, mulokVisible, seniVisible]);

  const visibleColumns = useMemo(() => allColumns.filter((c) => c.visible), [allColumns]);
  const seniColSpan = useMemo(
    () => visibleColumns.filter((c) => c.group === 'seni').length,
    [visibleColumns]
  );
  const mulokColSpan = useMemo(
    () => visibleColumns.filter((c) => c.group === 'mulok').length,
    [visibleColumns]
  );

  // Resolve KKM per column from db.mataPelajaran
  const getColumnKkm = (col: LegerSubjectColumn): number => {
    const found = db.mataPelajaran.find((m) => {
      const mName = m.mataPelajaran.trim().toLowerCase();
      if (col.isAgama && mName.includes('agama')) return true;
      return col.matchNames.some((mn) => mn === mName);
    });
    return found?.kkm ?? col.defaultKkm;
  };

  // Resolve AcademicScore for a student + column
  const findStudentColScore = (student: Student, col: LegerSubjectColumn): AcademicScore | undefined => {
    const studentScores = db.nilaiAkademik.filter(
      (n) => matchesStudentScore(n, student) && n.semester === selectedSemester
    );

    if (col.isAgama) {
      const agamaName = (student.agama || '').trim().toLowerCase();
      // Prioritize exact student religion match first
      const exactReligion = agamaName
        ? studentScores.find((n) => {
            const mp = n.mataPelajaran.trim().toLowerCase();
            return mp.includes('agama') && mp.includes(agamaName);
          })
        : undefined;
      if (exactReligion) return exactReligion;
      return studentScores.find((n) => n.mataPelajaran.trim().toLowerCase().includes('agama'));
    }

    return studentScores.find((n) => {
      const mp = n.mataPelajaran.trim().toLowerCase();
      return col.matchNames.some((mn) => mn === mp);
    });
  };

  // Build matrix for ALL students in the class (so rank is always accurate across the whole class)
  const classMatrix = useMemo(() => {
    return siswaKelas.map((student) => {
      const scores = visibleColumns.map((col) => {
        const rec = findStudentColScore(student, col);
        return rec && rec.nilaiAkhir > 0 ? rec.nilaiAkhir : null;
      });
      const valid = scores.filter((v): v is number => v !== null);
      const total = valid.reduce((acc, v) => acc + v, 0);
      const avg = valid.length > 0 ? parseFloat((total / valid.length).toFixed(1)) : 0;
      return {
        student,
        scores,
        total,
        avg,
        hasAnyScore: valid.length > 0
      };
    });
  }, [siswaKelas, visibleColumns, db.nilaiAkademik, selectedSemester]);

  // Compute class ranking (dense/standard competition rank among students who have at least 1 score)
  const rankMap = useMemo(() => {
    const map = new Map<string, number | string>();
    const withScores = [...classMatrix]
      .filter((r) => r.hasAnyScore)
      .sort((a, b) => b.total - a.total || b.avg - a.avg);
    withScores.forEach((row, idx) => {
      map.set(row.student.idSiswa, idx + 1);
    });
    return map;
  }, [classMatrix]);

  const displayedMatrix = useMemo(() => {
    if (filteredStudent) {
      return classMatrix.filter((r) => r.student.idSiswa === filteredStudent.idSiswa);
    }
    return classMatrix;
  }, [classMatrix, filteredStudent]);

  // Pad to 60 rows when unfiltered to match classic 60-student leger layout
  const totalRowSlots = filteredStudent
    ? Math.max(displayedMatrix.length, 1)
    : Math.max(MAX_SHEET_ROWS, displayedMatrix.length);

  // Compute 5 footer summary stats across classMatrix for each visible column
  const columnStats = useMemo(() => {
    return visibleColumns.map((col, colIdx) => {
      const kkm = getColumnKkm(col);
      const values = classMatrix
        .map((row) => row.scores[colIdx])
        .filter((v): v is number => v !== null);
      const sum = values.reduce((a, b) => a + b, 0);
      const avg = values.length > 0 ? parseFloat((sum / values.length).toFixed(1)) : 0;
      const tuntas = values.filter((v) => v >= kkm).length;
      const dayaSerap = values.length > 0 ? Math.round((tuntas / values.length) * 100) : 0;
      const maxVal = values.length > 0 ? Math.max(...values) : 0;
      const minVal = values.length > 0 ? Math.min(...values) : 0;

      return {
        colId: col.id,
        label: col.shortLabel,
        barColor: col.barColor,
        kkm,
        count: values.length,
        sum,
        avg,
        dayaSerap,
        maxVal,
        minVal
      };
    });
  }, [visibleColumns, classMatrix, db.mataPelajaran]);

  // Overall summary stats for right-hand summary columns
  const totalSummaryStats = useMemo(() => {
    const totals = classMatrix.filter((r) => r.hasAnyScore).map((r) => r.total);
    const avgs = classMatrix.filter((r) => r.hasAnyScore).map((r) => r.avg);
    return {
      sumOfTotals: totals.reduce((a, b) => a + b, 0),
      avgOfAvgs:
        avgs.length > 0
          ? parseFloat((avgs.reduce((a, b) => a + b, 0) / avgs.length).toFixed(1))
          : 0
    };
  }, [classMatrix]);

  const handleSaveTempatTanggal = (val: string) => {
    setTempatTanggalInput(val);
    updateSekolah({ tempatTanggalRapor: val });
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const headerNames = visibleColumns.map((c) => c.label);
    const data: (string | number)[][] = [
      ['REKAP NILAI SISWA (LEGER NILAI AKADEMIK)'],
      [`Nama Sekolah : ${db.sekolah.namaSekolah}`, '', `Semester : ${selectedSemester === 'Ganjil' ? 'I (Ganjil)' : 'II (Genap)'}`],
      [`Kelas : ${db.sekolah.rombel || selectedKelas}${db.sekolah.kelasParalel || ''}`, '', `Tahun Pel. : ${db.sekolah.tahunPelajaran}`],
      [],
      ['NO', 'NAMA SISWA', ...headerNames, 'JUMLAH', 'RATA-RATA', 'PERINGKAT'],
      ...classMatrix.map((row, idx) => [
        idx + 1,
        row.student.namaSiswa,
        ...row.scores.map((sc) => sc ?? ''),
        row.hasAnyScore ? row.total : '',
        row.hasAnyScore ? row.avg : '',
        rankMap.get(row.student.idSiswa) ?? ''
      ]),
      ['Jumlah', '', ...columnStats.map((st) => (st.count > 0 ? st.sum : 0)), totalSummaryStats.sumOfTotals, '', ''],
      ['Rata-rata Kelas', '', ...columnStats.map((st) => (st.count > 0 ? st.avg : 0)), '', totalSummaryStats.avgOfAvgs, ''],
      ['Daya Serap (%)', '', ...columnStats.map((st) => `${st.dayaSerap}%`), '', '', ''],
      ['Ns Tertinggi', '', ...columnStats.map((st) => (st.count > 0 ? st.maxVal : 0)), '', '', ''],
      ['Ns Terendah', '', ...columnStats.map((st) => (st.count > 0 ? st.minVal : 0)), '', '', '']
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 30 },
      ...visibleColumns.map(() => ({ wch: 14 })),
      { wch: 10 },
      { wch: 11 },
      { wch: 11 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Rekap Nilai');
    XLSX.writeFile(wb, `rekap-nilai-kelas-${selectedKelas}-${selectedSemester}.xlsx`);
    showSaveNotification('Leger Rekap Nilai Siswa berhasil diekspor ke Excel (.xlsx)!');
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await parseAndImportTemplate(file, 'template_nilai_akademik', appContext);
    e.target.value = '';
  };

  const toggleSeni = (key: 'seniMusik' | 'seniTari' | 'seniRupa' | 'seniTeater') => {
    setSeniVisible((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const toggleMulok = (index: number) => {
    setMulokVisible((prev) => prev.map((val, idx) => (idx === index ? !val : val)));
  };

  const kelasDisplay = `${db.sekolah.rombel || selectedKelas}${db.sekolah.kelasParalel || ''}`.trim();
  const semesterDisplay = selectedSemester === 'Ganjil' ? 'I (Ganjil)' : 'II (Genap)';

  return (
    <div className="bg-[#0066aa] min-h-screen p-3 sm:p-4 rounded-xl shadow-inner print:bg-white print:p-0 print:rounded-none print:shadow-none">
      {/* TOP BLACK CONTROL STRIP (hidden on print) */}
      <div className="bg-black border-2 border-red-600 p-2 mb-3 flex flex-wrap items-center justify-between gap-2 print:hidden">
        <div className="flex flex-wrap items-center gap-2">
          {/* Circular Blue Home Button */}
          <button
            onClick={() => setActiveMenu('dashboard')}
            className="w-10 h-10 rounded-full bg-sky-600 hover:bg-sky-500 text-white border-2 border-sky-300 flex items-center justify-center shadow cursor-pointer transition-transform active:scale-95"
            title="Kembali ke Dashboard Utama"
          >
            <Home size={18} />
          </button>

          {/* Cyan Aturan Mata Pelajaran Button */}
          <button
            onClick={() => setActiveMenu('kkm_pelajaran')}
            className="px-3 py-1.5 bg-cyan-400 hover:bg-cyan-300 text-slate-950 font-extrabold text-[11px] uppercase tracking-tight border border-white/60 shadow cursor-pointer leading-tight text-center"
            title="Buka Pengaturan KKM, Bobot & Mata Pelajaran"
          >
            Aturan
            <br />
            Mata Pelajaran
          </button>

          {/* Red Cetak / Print Leger Button */}
          <button
            onClick={() => window.print()}
            className="px-3 py-1.5 bg-red-600 hover:bg-red-500 text-white font-extrabold text-[11px] uppercase tracking-tight border border-white/60 shadow cursor-pointer leading-tight inline-flex items-center gap-1.5"
            title="Cetak Rekap Nilai Siswa (Legér & Grafik Daya Serap)"
          >
            <Printer size={15} />
            <span className="text-left">
              Cetak Leger
              <br />
              Nilai Rapor
            </span>
          </button>

          {/* Purple Help Button */}
          <button
            onClick={() => setShowHelp((prev) => !prev)}
            className="w-9 h-9 rounded-lg bg-purple-700 hover:bg-purple-600 text-white border border-purple-300 flex items-center justify-center shadow cursor-pointer"
            title="Panduan & Rumus Rekap Nilai"
          >
            <HelpCircle size={20} />
          </button>

          {/* Compact 2-row Seni Pilihan & Mulok Column Toggles (matching screenshot) */}
          <div className="grid grid-cols-5 gap-1 bg-slate-900/90 p-1 border border-slate-700 rounded">
            <button
              onClick={() => toggleSeni('seniMusik')}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                seniVisible.seniMusik
                  ? 'bg-[#ffff00] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Seni Musik"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {seniVisible.seniMusik ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Seni Pil. 1</span>
            </button>

            <button
              onClick={() => toggleSeni('seniRupa')}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                seniVisible.seniRupa
                  ? 'bg-[#0284c7] text-white border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Seni Rupa"
            >
              <span className="w-3 h-3 bg-white text-black border border-black inline-flex items-center justify-center text-[8px]">
                {seniVisible.seniRupa ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Seni Pil. 3</span>
            </button>

            <button
              onClick={() => toggleMulok(0)}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                mulokVisible[0]
                  ? 'bg-[#cbd5e1] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Muatan Lokal 1"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {mulokVisible[0] ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Mulok 1</span>
            </button>

            <button
              onClick={() => toggleMulok(2)}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                mulokVisible[2]
                  ? 'bg-[#00ffff] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Muatan Lokal 3"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {mulokVisible[2] ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Mulok 3</span>
            </button>

            <button
              onClick={() => toggleMulok(4)}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                mulokVisible[4]
                  ? 'bg-white text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Muatan Lokal 5"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {mulokVisible[4] ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Mulok 5</span>
            </button>

            <button
              onClick={() => toggleSeni('seniTari')}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                seniVisible.seniTari
                  ? 'bg-[#e9d5ff] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Seni Tari"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {seniVisible.seniTari ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Seni Pil. 2</span>
            </button>

            <button
              onClick={() => toggleSeni('seniTeater')}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                seniVisible.seniTeater
                  ? 'bg-[#f59e0b] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Seni Teater"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {seniVisible.seniTeater ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Seni Pil. 4</span>
            </button>

            <button
              onClick={() => toggleMulok(1)}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                mulokVisible[1]
                  ? 'bg-[#94a3b8] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Muatan Lokal 2"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {mulokVisible[1] ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Mulok 2</span>
            </button>

            <button
              onClick={() => toggleMulok(3)}
              className={`px-1.5 py-0.5 text-[10px] font-bold border flex items-center gap-1 cursor-pointer ${
                mulokVisible[3]
                  ? 'bg-[#84cc16] text-slate-950 border-black'
                  : 'bg-slate-700 text-slate-300 border-slate-600 opacity-60'
              }`}
              title="Tampilkan / Sembunyikan Muatan Lokal 4"
            >
              <span className="w-3 h-3 bg-white border border-black inline-flex items-center justify-center text-[8px]">
                {mulokVisible[3] ? <Check size={9} strokeWidth={3} /> : ''}
              </span>
              <span>Mulok 4</span>
            </button>

            <button
              onClick={() => setActiveMenu('input_nilai')}
              className="px-1.5 py-0.5 text-[10px] font-bold bg-emerald-500 hover:bg-emerald-400 text-slate-950 border border-black flex items-center justify-center gap-1 cursor-pointer"
              title="Buka Input Nilai Akademik"
            >
              <BookOpen size={10} />
              <span>Input Nilai</span>
            </button>
          </div>
        </div>

        {/* Right Utilities: Filter, Export Excel/CSV/Workbook, Import */}
        <div className="flex flex-wrap items-center gap-2">
          <div className="bg-white rounded px-1.5 py-0.5">
            <StudentFilterSelect />
          </div>

          <button
            onClick={handleExportExcel}
            className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-[11px] font-bold rounded flex items-center gap-1 border border-emerald-300 cursor-pointer"
            title="Ekspor Leger Rekap Nilai (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>Excel</span>
          </button>

          <button
            onClick={() => exportClassLegerToCsv(db, selectedKelas, selectedSemester)}
            className="px-2.5 py-1.5 bg-slate-700 hover:bg-slate-600 text-white text-[11px] font-bold rounded flex items-center gap-1 border border-slate-400 cursor-pointer"
            title="Ekspor Leger ke CSV"
          >
            <Download size={13} />
            <span>CSV</span>
          </button>

          <button
            onClick={() => exportClassComprehensiveWorkbook(db, selectedKelas, selectedSemester)}
            className="px-2.5 py-1.5 bg-amber-500 hover:bg-amber-400 text-slate-950 text-[11px] font-bold rounded flex items-center gap-1 border border-amber-200 cursor-pointer"
            title="Ekspor Buku Nilai Sekelas Multi-Sheet (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>Buku Nilai</span>
          </button>

          <label
            className="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-[11px] font-bold rounded flex items-center gap-1 border border-indigo-300 cursor-pointer"
            title="Impor Nilai dari Template Excel"
          >
            <Upload size={13} />
            <span>Impor</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} className="hidden" />
          </label>
        </div>
      </div>

      {/* HELP PANEL */}
      {showHelp && (
        <div className="mb-3 bg-white border-2 border-purple-700 rounded-lg p-3 shadow-lg text-xs text-slate-800 print:hidden">
          <div className="flex items-center justify-between border-b border-purple-200 pb-1.5 mb-2">
            <h3 className="font-bold text-purple-900">Panduan & Rumus Rekap Nilai Siswa (Legér Kelas)</h3>
            <button
              onClick={() => setShowHelp(false)}
              className="p-1 hover:bg-purple-100 rounded text-purple-700 cursor-pointer"
              aria-label="Tutup Panduan"
            >
              <X size={15} />
            </button>
          </div>
          <ul className="list-disc pl-5 space-y-1">
            <li>
              <strong>Rumus Terhubung Otomatis:</strong> Seluruh nilai pada tabel ini diambil langsung dari menu{' '}
              <em>Input Nilai</em> (menggunakan bobot S. TP, S. LM, S. AS &amp; KKM dari menu <em>Aturan Mata Pelajaran</em>).
            </li>
            <li>
              <strong>Pendidikan Agama:</strong> Kolom 1 otomatis membaca nilai agama sesuai agama masing-masing siswa pada{' '}
              <em>Identitas Siswa</em> (Islam, Kristen, Katolik, Hindu, Buddha, Konghucu).
            </li>
            <li>
              <strong>Daya Serap &amp; Grafik:</strong> Daya Serap menghitung persentase siswa yang memperoleh nilai &ge; KKM per mata pelajaran dan langsung digambarkan pada <em>Grafik Daya Serap</em> di bawah tabel.
            </li>
            <li>
              <strong>Klik Cepat:</strong> Klik <strong>Nama Siswa</strong> untuk membuka pratinjau Rapor Siswa, atau klik <strong>sel nilai</strong> untuk langsung mengedit di menu Input Nilai.
            </li>
          </ul>
        </div>
      )}

      {/* MAIN WORKSHEET + RIGHT SIDEBAR LAYOUT */}
      <div className="flex flex-col xl:flex-row items-start gap-4">
        {/* WHITE PRINTABLE SHEET */}
        <div className="flex-1 w-full bg-white border-2 border-black shadow-2xl p-4 sm:p-6 print:border-0 print:shadow-none print:p-2 overflow-x-auto">
          {/* Sheet Title */}
          <div className="text-center mb-3">
            <h1 className="text-base sm:text-lg font-extrabold tracking-[0.18em] uppercase text-black font-serif">
              REKAP NILAI SISWA
            </h1>
          </div>

          {/* School / Class / Semester / Year Metadata */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-1 text-xs font-bold text-black mb-3 font-serif">
            <div className="space-y-0.5">
              <div className="flex">
                <span className="w-28 shrink-0">Nama Sekolah</span>
                <span>: {db.sekolah.namaSekolah}</span>
              </div>
              <div className="flex">
                <span className="w-28 shrink-0">Kelas</span>
                <span>: {kelasDisplay}</span>
              </div>
            </div>
            <div className="space-y-0.5">
              <div className="flex">
                <span className="w-24 shrink-0">Semester</span>
                <span>: {semesterDisplay}</span>
              </div>
              <div className="flex">
                <span className="w-24 shrink-0">Tahun Pel.</span>
                <span>: {db.sekolah.tahunPelajaran}</span>
              </div>
            </div>
          </div>

          {/* MAIN LEGER MATRIX TABLE */}
          <div className="overflow-x-auto">
            <table className="w-full border-collapse border-2 border-black text-[10px] text-black">
              <thead>
                {/* Row 1: Top Header */}
                <tr>
                  <th
                    rowSpan={4}
                    className="border border-black bg-[#0099ff] text-black font-extrabold text-center px-1.5 py-1 w-8 sticky left-0 z-20"
                  >
                    NO
                  </th>
                  <th
                    rowSpan={4}
                    className="border border-black bg-[#0099ff] text-black font-extrabold text-center px-2 py-1 min-w-[165px] sticky left-8 z-20"
                  >
                    NAMA SISWA
                  </th>
                  <th
                    colSpan={Math.max(visibleColumns.length, 1)}
                    className="border border-black bg-[#00ff00] text-black font-extrabold tracking-[0.2em] text-center py-1 text-xs font-serif"
                  >
                    CAPAIAN MATA PELAJARAN
                  </th>
                  <th
                    rowSpan={4}
                    className="border border-black bg-[#00aaff] text-black font-extrabold text-center w-10 px-1 py-1 align-middle"
                  >
                    <div className="[writing-mode:vertical-rl] rotate-180 mx-auto text-[10px] tracking-wider">
                      JUMLAH
                    </div>
                  </th>
                  <th
                    rowSpan={4}
                    className="border border-black bg-[#00ff00] text-black font-extrabold text-center w-10 px-1 py-1 align-middle"
                  >
                    <div className="[writing-mode:vertical-rl] rotate-180 mx-auto text-[10px] tracking-wider">
                      RATA-RATA
                    </div>
                  </th>
                  <th
                    rowSpan={4}
                    className="border border-black bg-[#fbbf24] text-black font-extrabold text-center w-10 px-1 py-1 align-middle"
                  >
                    <div className="[writing-mode:vertical-rl] rotate-180 mx-auto text-[10px] tracking-wider">
                      PERINGKAT
                    </div>
                  </th>
                </tr>

                {/* Row 2: Subject Group & Standalone Subject Names */}
                <tr>
                  {(() => {
                    const cells: React.ReactNode[] = [];
                    let i = 0;
                    while (i < visibleColumns.length) {
                      const col = visibleColumns[i];
                      if (col.group === 'seni') {
                        cells.push(
                          <th
                            key="group-seni"
                            colSpan={seniColSpan}
                            className="border border-black bg-white text-black font-bold italic text-center py-1 px-1 font-serif text-[11px]"
                          >
                            Seni Pilihan
                          </th>
                        );
                        while (i < visibleColumns.length && visibleColumns[i].group === 'seni') i++;
                      } else if (col.group === 'mulok') {
                        cells.push(
                          <th
                            key="group-mulok"
                            colSpan={mulokColSpan}
                            className="border border-black bg-white text-black font-bold italic text-center py-1 px-1 font-serif text-[11px]"
                          >
                            Muatan Lokal
                          </th>
                        );
                        while (i < visibleColumns.length && visibleColumns[i].group === 'mulok') i++;
                      } else {
                        cells.push(
                          <th
                            key={col.id}
                            rowSpan={2}
                            style={{ backgroundColor: col.headerBg, color: col.headerText }}
                            className="border border-black font-bold text-center px-1 py-1 min-w-[54px] max-w-[78px] leading-tight text-[9px] align-middle"
                            title={`${col.label} (KKM: ${getColumnKkm(col)})`}
                          >
                            {col.label}
                          </th>
                        );
                        i++;
                      }
                    }
                    return cells;
                  })()}
                </tr>

                {/* Row 3: Sub-columns inside Seni Pilihan & Muatan Lokal */}
                <tr>
                  {visibleColumns
                    .filter((c) => c.group === 'seni' || c.group === 'mulok')
                    .map((col) => (
                      <th
                        key={col.id}
                        style={{ backgroundColor: col.headerBg, color: col.headerText }}
                        className="border border-black font-bold text-center px-1 py-1 min-w-[50px] max-w-[72px] leading-tight text-[9px] align-middle"
                        title={`${col.label} (KKM: ${getColumnKkm(col)})`}
                      >
                        {col.shortLabel}
                      </th>
                    ))}
                </tr>

                {/* Row 4: NILAI AKHIR Green Sub-bar */}
                <tr>
                  <th
                    colSpan={Math.max(visibleColumns.length, 1)}
                    className="border border-black bg-[#00ff00] text-black font-extrabold tracking-[0.22em] text-center py-0.5 text-[10px] font-serif"
                  >
                    NILAI AKHIR
                  </th>
                </tr>
              </thead>

              <tbody>
                {Array.from({ length: totalRowSlots }).map((_, rowIdx) => {
                  const row = displayedMatrix[rowIdx];
                  const rowNum = filteredStudent ? 1 : rowIdx + 1;
                  const rank = row ? rankMap.get(row.student.idSiswa) : undefined;

                  return (
                    <tr
                      key={row ? row.student.idSiswa : `empty-${rowIdx}`}
                      className={row ? 'hover:bg-sky-50/70 transition-colors' : ''}
                    >
                      <td className="border border-black text-center font-medium py-0.5 px-1 bg-white sticky left-0 z-10 h-[19px]">
                        {rowNum}
                      </td>
                      <td className="border border-black px-2 py-0.5 font-semibold text-left bg-white sticky left-8 z-10 whitespace-nowrap">
                        {row ? (
                          <div className="flex items-center justify-between gap-1.5">
                            <button
                              onClick={() => openReportPreview(row.student.idSiswa)}
                              className="text-left hover:text-sky-700 hover:underline cursor-pointer truncate max-w-[180px]"
                              title={`Klik untuk pratinjau rapor ${row.student.namaSiswa}`}
                            >
                              {row.student.namaSiswa}
                            </button>
                            <button
                              onClick={() => {
                                setPrintTargetStudentId(row.student.idSiswa);
                                setActiveMenu('rapor_akademik');
                              }}
                              className="print:hidden text-slate-400 hover:text-emerald-700 cursor-pointer shrink-0"
                              title="Buka halaman Cetak Rapor siswa"
                            >
                              <Eye size={11} />
                            </button>
                          </div>
                        ) : (
                          ''
                        )}
                      </td>

                      {visibleColumns.map((col, colIdx) => {
                        const score = row ? row.scores[colIdx] : null;
                        const kkm = getColumnKkm(col);
                        const isBelowKkm = score !== null && score < kkm;
                        return (
                          <td
                            key={col.id}
                            onClick={row ? () => setActiveMenu('input_nilai') : undefined}
                            className={`border border-black text-center py-0.5 px-1 font-mono ${
                              row ? 'cursor-pointer hover:bg-amber-50' : ''
                            } ${isBelowKkm ? 'text-red-600 font-bold bg-red-50/60' : 'text-slate-900 font-semibold'}`}
                            title={
                              row
                                ? score !== null
                                  ? `${row.student.namaSiswa} • ${col.label}: ${score} (KKM ${kkm}) — Klik untuk ubah nilai`
                                  : `Belum ada nilai ${col.label} — Klik untuk input`
                                : undefined
                            }
                          >
                            {score !== null ? score : ''}
                          </td>
                        );
                      })}

                      <td className="border border-black text-center py-0.5 px-1 font-mono font-bold bg-sky-50/40">
                        {row && row.hasAnyScore ? row.total : ''}
                      </td>
                      <td className="border border-black text-center py-0.5 px-1 font-mono font-bold bg-emerald-50/40">
                        {row && row.hasAnyScore ? row.avg : ''}
                      </td>
                      <td className="border border-black text-center py-0.5 px-1 font-mono font-extrabold bg-amber-50/50">
                        {rank ?? ''}
                      </td>
                    </tr>
                  );
                })}

                {/* 5 Summary Footer Rows matching screenshot */}
                <tr className="bg-slate-50 font-bold">
                  <td colSpan={2} className="border border-black px-2 py-1 text-left sticky left-0 bg-slate-50 z-10">
                    Jumlah
                  </td>
                  {columnStats.map((st) => (
                    <td key={`sum-${st.colId}`} className="border border-black text-center py-1 px-1 font-mono">
                      {st.count > 0 ? st.sum : 0}
                    </td>
                  ))}
                  <td className="border border-black text-center py-1 px-1 font-mono bg-sky-50">
                    {totalSummaryStats.sumOfTotals > 0 ? totalSummaryStats.sumOfTotals : ''}
                  </td>
                  <td colSpan={2} className="border border-black bg-slate-100" />
                </tr>

                <tr className="bg-white font-bold">
                  <td colSpan={2} className="border border-black px-2 py-1 text-left sticky left-0 bg-white z-10">
                    Rata-rata Kelas
                  </td>
                  {columnStats.map((st) => (
                    <td key={`avg-${st.colId}`} className="border border-black text-center py-1 px-1 font-mono">
                      {st.count > 0 ? st.avg : 0}
                    </td>
                  ))}
                  <td className="border border-black bg-slate-100" />
                  <td className="border border-black text-center py-1 px-1 font-mono bg-emerald-50">
                    {totalSummaryStats.avgOfAvgs > 0 ? totalSummaryStats.avgOfAvgs : ''}
                  </td>
                  <td className="border border-black bg-slate-100" />
                </tr>

                <tr className="bg-slate-50 font-bold">
                  <td colSpan={2} className="border border-black px-2 py-1 text-left sticky left-0 bg-slate-50 z-10">
                    Daya Serap
                  </td>
                  {columnStats.map((st) => (
                    <td key={`ds-${st.colId}`} className="border border-black text-center py-1 px-1 font-mono">
                      {st.dayaSerap}%
                    </td>
                  ))}
                  <td colSpan={3} className="border border-black bg-slate-100" />
                </tr>

                <tr className="bg-white font-bold">
                  <td colSpan={2} className="border border-black px-2 py-1 text-left sticky left-0 bg-white z-10">
                    Ns Tertinggi
                  </td>
                  {columnStats.map((st) => (
                    <td key={`max-${st.colId}`} className="border border-black text-center py-1 px-1 font-mono">
                      {st.count > 0 ? st.maxVal : 0}
                    </td>
                  ))}
                  <td colSpan={3} className="border border-black bg-slate-100" />
                </tr>

                <tr className="bg-slate-50 font-bold">
                  <td colSpan={2} className="border border-black px-2 py-1 text-left sticky left-0 bg-slate-50 z-10">
                    Ns Terendah
                  </td>
                  {columnStats.map((st) => (
                    <td key={`min-${st.colId}`} className="border border-black text-center py-1 px-1 font-mono">
                      {st.count > 0 ? st.minVal : 0}
                    </td>
                  ))}
                  <td colSpan={3} className="border border-black bg-slate-100" />
                </tr>
              </tbody>
            </table>
          </div>

          {/* GRAFIK DAYA SERAP SECTION */}
          <div className="mt-6 pt-4 border-t border-slate-300">
            <div className="flex items-start gap-2">
              <div className="text-[10px] font-bold text-slate-700 pt-1 whitespace-nowrap">
                DAYA SERAP (%)
              </div>
              <div className="flex-1">
                <h3 className="text-center text-xs sm:text-sm font-extrabold tracking-wider uppercase text-slate-800 mb-2">
                  GRAFIK DAYA SERAP
                </h3>

                <div className="relative pl-7 pr-2">
                  {/* Plot Area */}
                  <div className="relative h-44 border-l border-b border-slate-400 flex flex-col justify-between">
                    {[100, 80, 60, 40, 20, 0].map((tick) => (
                      <div key={tick} className="relative w-full border-t border-slate-200 first:border-t">
                        <span className="absolute -left-7 -top-2 text-[9px] font-mono text-slate-600 w-5 text-right">
                          {tick}
                        </span>
                      </div>
                    ))}

                    {/* Bars Row */}
                    <div
                      className="absolute inset-0 flex items-end justify-around px-2 pb-0.5 z-10"
                      style={{ bottom: '4px' }}
                    >
                      {columnStats.map((st) => {
                        const heightPct = Math.max(2, st.dayaSerap);
                        const barFill = pemutihanGrafik ? '#f8fafc' : st.barColor;
                        return (
                          <div
                            key={`bar-${st.colId}`}
                            className="flex flex-col items-center justify-end h-full flex-1 max-w-[38px]"
                            title={`${st.label}: Daya Serap ${st.dayaSerap}% (KKM ${st.kkm})`}
                          >
                            <span className="text-[9px] font-bold text-slate-800 mb-0.5">
                              {st.dayaSerap}
                            </span>
                            <div
                              className="w-5 sm:w-6 border border-slate-700 transition-all duration-300 shadow-xs"
                              style={{
                                height: st.dayaSerap === 0 ? '4px' : `${heightPct}%`,
                                backgroundColor: st.dayaSerap === 0 ? '#22c55e' : barFill
                              }}
                            />
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Cyan 3D-style Chart Floor Strip */}
                  <div
                    className={`h-3 w-full border-x border-b border-slate-500 ${
                      pemutihanGrafik ? 'bg-slate-100' : 'bg-cyan-400'
                    }`}
                  />

                  {/* Vertical Subject Labels Below Chart */}
                  <div className="flex items-start justify-around px-2 pt-1.5">
                    {columnStats.map((st) => (
                      <div key={`lbl-${st.colId}`} className="flex-1 max-w-[38px] flex justify-center">
                        <span className="[writing-mode:vertical-rl] rotate-180 text-[8.5px] font-semibold text-slate-800 leading-tight max-h-28 truncate">
                          {st.label}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* SIGNATURE BLOCK AT BOTTOM OF WHITE SHEET */}
          <div className="mt-8 grid grid-cols-2 gap-8 text-xs text-black px-4 sm:px-8 pb-4">
            <div>
              <p>Mengetahui,</p>
              <p className="font-semibold">{pengesahanOleh || 'Kepala Sekolah'}</p>
              <div className="h-16" />
              <p className="font-bold underline">
                {db.sekolah.namaKepalaSekolah || '................................'}
              </p>
              <p>NIP. {db.sekolah.nipKepalaSekolah || '-'}</p>
            </div>

            <div className="pl-4 sm:pl-16">
              <p>{tempatTanggalInput || db.sekolah.tempatTanggalRapor || 'Oksibil, 10 Juli 2026'}</p>
              <p className="font-semibold">Guru Kelas,</p>
              <div className="h-16" />
              <p className="font-bold underline">
                {db.sekolah.waliKelas || '................................'}
              </p>
              <p>NIP. {db.sekolah.nipWaliKelas || '-'}</p>
            </div>
          </div>
        </div>

        {/* RIGHT BLUE CANVAS CONTROLS (Grafik Toggle + Ganti Tempat & Tanggal + Pengesahan Oleh) */}
        <div className="w-full xl:w-64 shrink-0 flex flex-col justify-between gap-6 xl:self-stretch print:hidden">
          <div className="xl:mt-auto xl:mb-12 flex flex-col items-start gap-4">
            {/* Pink Pemutihan Grafik Button matching screenshot */}
            <button
              onClick={() => setPemutihanGrafik((prev) => !prev)}
              className={`px-3 py-1.5 border-2 border-black shadow font-extrabold text-[10px] uppercase leading-tight cursor-pointer transition-colors ${
                pemutihanGrafik
                  ? 'bg-white text-fuchsia-700'
                  : 'bg-fuchsia-400 hover:bg-fuchsia-300 text-slate-950'
              }`}
              title="Ubah tampilan warna grafik menjadi putih polos (hemat tinta cetak) atau berwarna"
            >
              {pemutihanGrafik ? 'Tampilkan Warna' : 'Pemutihan'}
              <br />
              Grafik Daya Serap
            </button>

            {/* Ganti Tempat dan Tanggal */}
            <div className="w-full">
              <div className="bg-black text-white font-bold italic text-[11px] px-2 py-0.5 mb-1 inline-block">
                Ganti Tempat dan Tanggal
              </div>
              <input
                type="text"
                value={tempatTanggalInput}
                onChange={(e) => handleSaveTempatTanggal(e.target.value)}
                placeholder="Contoh: Oksibil, 10 Juli 2026"
                className="w-full bg-white text-slate-950 font-semibold text-xs px-2 py-1 border-2 border-red-600 focus:outline-none focus:ring-2 focus:ring-yellow-400"
              />
            </div>

            {/* Pengesahan Oleh */}
            <div className="w-full">
              <div className="bg-black text-white font-bold italic text-[11px] px-2 py-0.5 mb-1 inline-block">
                Pengesahan Oleh
              </div>
              <input
                type="text"
                list="pengesahan-options"
                value={pengesahanOleh}
                onChange={(e) => setPengesahanOleh(e.target.value)}
                placeholder="Kepala Sekolah"
                className="w-full bg-white text-slate-950 font-semibold text-xs px-2 py-1 border-2 border-red-600 focus:outline-none focus:ring-2 focus:ring-yellow-400"
              />
              <datalist id="pengesahan-options">
                <option value="Kepala Sekolah" />
                <option value="Plt. Kepala Sekolah" />
                <option value="Komite Sekolah" />
              </datalist>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
