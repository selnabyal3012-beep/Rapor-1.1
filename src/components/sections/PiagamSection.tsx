import React, { useState, useEffect, useRef } from 'react';
import {
  Printer,
  Download,
  Upload,
  Home,
  Lock,
  Unlock,
  X,
  RotateCcw,
  Check,
  ChevronUp,
  ChevronDown,
  FileSpreadsheet,
  SlidersHorizontal
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { LogoKabupaten, LogoTutWuriHandayani } from '../common/OfficialLogos';
import * as XLSX from 'xlsx';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentScore } from '../../utils/studentSyncHelper';

interface AwardPreset {
  id: string;
  label: string;
  rankTitle: string;
  verseLine1: string;
  verseLine2: string;
}

const AWARD_PRESETS: AwardPreset[] = [
  {
    id: 'peringkat_otomatis',
    label: 'Otomatis Sesuai Peringkat Kelas',
    rankTitle: '',
    verseLine1: 'Semoga Prestasi Yang Diraih Menjadi Motivasi Untuk Meraih Kesuksesan',
    verseLine2: 'Di Masa Yang Akan Datang'
  },
  {
    id: 'juara1',
    label: 'Peringkat Pertama (Juara I)',
    rankTitle: 'Peringkat Pertama',
    verseLine1: 'Semoga Prestasi Yang Diraih Menjadi Motivasi Untuk Meraih Kesuksesan',
    verseLine2: 'Di Masa Yang Akan Datang'
  },
  {
    id: 'juara2',
    label: 'Peringkat Kedua (Juara II)',
    rankTitle: 'Peringkat Kedua',
    verseLine1: 'Semoga Prestasi Yang Diraih Menjadi Motivasi Untuk Meraih Kesuksesan',
    verseLine2: 'Di Masa Yang Akan Datang'
  },
  {
    id: 'juara3',
    label: 'Peringkat Ketiga (Juara III)',
    rankTitle: 'Peringkat Ketiga',
    verseLine1: 'Semoga Prestasi Yang Diraih Menjadi Motivasi Untuk Meraih Kesuksesan',
    verseLine2: 'Di Masa Yang Akan Datang'
  },
  {
    id: 'teladan_p5',
    label: 'Siswa Teladan Profil Pelajar Pancasila',
    rankTitle: 'Siswa Teladan P5',
    verseLine1: 'Atas Pengamalan Nilai-Nilai Profil Pelajar Pancasila Dengan Akhlak Mulia',
    verseLine2: 'Semoga Menjadi Teladan Di Lingkungan Sekolah Dan Masyarakat'
  },
  {
    id: 'rajin',
    label: 'Kehadiran & Disiplin Terbaik',
    rankTitle: 'Siswa Teladan Disiplin & Kehadiran Terbaik',
    verseLine1: 'Atas Keteladanan Kedisiplinan Dan Kehadiran Penuh Sepanjang Semester',
    verseLine2: 'Semoga Tetap Konsisten Meraih Prestasi Gemilang'
  }
];

const FONT_OPTIONS = [
  { label: 'Times New Roman (Standar Resmi)', value: '"Tinos", "Times New Roman", Times, serif' },
  { label: 'Georgia (Klasik Elegan)', value: 'Georgia, "Times New Roman", serif' },
  { label: 'EB Garamond (Buku Resmi)', value: '"EB Garamond", Georgia, serif' },
  { label: 'Lora (Serif Modern)', value: '"Lora", Georgia, serif' },
  { label: 'Calibri / Carlito (Bersih)', value: '"Carlito", Calibri, Arial, sans-serif' },
  { label: 'Arial / Arimo (Tegas)', value: '"Arimo", Arial, Helvetica, sans-serif' }
];

const ORDINAL_ID = [
  '',
  'Pertama',
  'Kedua',
  'Ketiga',
  'Keempat',
  'Kelima',
  'Keenam',
  'Ketujuh',
  'Kedelapan',
  'Kesembilan',
  'Kesepuluh'
];

function getRankText(rank: number): string {
  if (rank <= 0) return 'Peringkat Pertama';
  if (rank < ORDINAL_ID.length) return `Peringkat ${ORDINAL_ID[rank]}`;
  return `Peringkat Ke-${rank}`;
}

/**
 * Ornamen Sudut Geometri Navy + Oranye + Slate Biru
 * Didesain presisi vektor agar tajam di layar dan saat dicetak A4 Landscape
 */
const CertificateCornerOrnament: React.FC<{ position: 'top-left' | 'bottom-right' }> = ({ position }) => {
  const isBottomRight = position === 'bottom-right';
  return (
    <svg
      viewBox="0 0 260 220"
      className={`pointer-events-none select-none absolute z-20 w-[180px] h-[152px] sm:w-[220px] sm:h-[186px] ${
        isBottomRight ? 'bottom-[-4px] right-[-4px] rotate-180' : 'top-[-4px] left-[-4px]'
      }`}
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <defs>
        <linearGradient id={`navySq-${position}`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#3c6696" />
          <stop offset="50%" stopColor="#264b73" />
          <stop offset="100%" stopColor="#142b47" />
        </linearGradient>
        <linearGradient id={`orangeSq-${position}`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#ff9429" />
          <stop offset="55%" stopColor="#f97316" />
          <stop offset="100%" stopColor="#c2410c" />
        </linearGradient>
        <linearGradient id={`slateSq-${position}`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#4f7cac" />
          <stop offset="100%" stopColor="#264b73" />
        </linearGradient>
        <filter id={`drop-${position}`} x="-10%" y="-10%" width="130%" height="130%">
          <feDropShadow dx="2" dy="2" stdDeviation="2" floodColor="#000000" floodOpacity="0.32" />
        </filter>
      </defs>

      {/* Bar Horizontal Atas: Navy tapered + Orange garis tipis */}
      <polygon points="56,14 196,14 226,32 56,32" fill="#2d5580" />
      <polygon points="56,33 230,33 244,39 56,39" fill="#f97316" />

      {/* Bar Vertikal Kiri: Navy tapered + Orange garis tipis */}
      <polygon points="14,56 32,56 32,202 14,172" fill="#2d5580" />
      <polygon points="33,56 39,56 39,214 33,200" fill="#f97316" />

      {/* 3 Kotak 3D Bertumpuk di Sudut */}
      <g filter={`url(#drop-${position})`}>
        {/* Kotak Besar Navy */}
        <rect x="4" y="4" width="58" height="58" fill={`url(#navySq-${position})`} stroke="#0f233a" strokeWidth="2.5" />
        <polyline points="7,60 7,7 60,7" stroke="#5b87b8" strokeWidth="2" fill="none" />

        {/* Kotak Sedang Oranye */}
        <rect x="34" y="34" width="38" height="38" fill={`url(#orangeSq-${position})`} stroke="#9a3412" strokeWidth="2" />
        <polyline points="36,70 36,36 70,36" stroke="#fdba74" strokeWidth="1.5" fill="none" />

        {/* Kotak Kecil Slate Biru */}
        <rect x="54" y="54" width="26" height="26" fill={`url(#slateSq-${position})`} stroke="#1e3a5f" strokeWidth="1.8" />
        <polyline points="56,78 56,56 78,56" stroke="#93c5fd" strokeWidth="1.2" fill="none" />
      </g>
    </svg>
  );
};

/**
 * Ilustrasi Trofi Emas 3D (Figur Emas Mengangkat Piala)
 */
const GoldTrophyIllustration: React.FC = () => (
  <svg
    viewBox="0 0 140 150"
    className="w-24 h-26 sm:w-28 sm:h-30 mx-auto drop-shadow-[0_4px_6px_rgba(161,98,7,0.28)] select-none"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <defs>
      <radialGradient id="goldHead" cx="38%" cy="32%" r="68%">
        <stop offset="0%" stopColor="#fff7cc" />
        <stop offset="40%" stopColor="#facc15" />
        <stop offset="80%" stopColor="#ca8a04" />
        <stop offset="100%" stopColor="#854d0e" />
      </radialGradient>
      <linearGradient id="goldBody" x1="0%" y1="0%" x2="100%" y2="100%">
        <stop offset="0%" stopColor="#fef08a" />
        <stop offset="45%" stopColor="#eab308" />
        <stop offset="100%" stopColor="#92400e" />
      </linearGradient>
      <linearGradient id="cupGold" x1="15%" y1="0%" x2="85%" y2="100%">
        <stop offset="0%" stopColor="#fffbeb" />
        <stop offset="30%" stopColor="#fde047" />
        <stop offset="65%" stopColor="#d97706" />
        <stop offset="100%" stopColor="#78350f" />
      </linearGradient>
    </defs>

    {/* Bayangan lembut di bawah kaki */}
    <ellipse cx="70" cy="142" rx="34" ry="5" fill="#d97706" opacity="0.18" />

    {/* Kepala Figur Emas */}
    <circle cx="46" cy="48" r="19" fill="url(#goldHead)" />

    {/* Tubuh & Kaki Figur Emas (pose melompat gembira) */}
    <path
      d="M52 67 C63 72 68 88 64 102 C60 113 55 123 64 134 C66 137 62 140 58 136 C48 126 48 112 52 99 C48 86 43 75 52 67 Z"
      fill="url(#goldBody)"
    />
    <path
      d="M57 96 C68 94 79 102 82 115 C83 119 78 120 75 116 C72 108 65 104 58 104 Z"
      fill="url(#goldBody)"
    />

    {/* Lengan kiri & kanan mengangkat piala */}
    <path
      d="M55 71 C64 60 72 51 77 45 C80 42 84 46 80 51 C74 59 66 69 59 77 Z"
      fill="url(#goldBody)"
    />
    <path
      d="M60 78 C71 78 82 74 88 67 C91 64 94 69 90 73 C82 81 71 86 58 84 Z"
      fill="url(#goldBody)"
    />

    {/* Gagang Piala Kiri & Kanan */}
    <path
      d="M74 26 C61 26 58 43 68 50 C71 52 75 49 72 46 C66 42 66 32 74 32 Z"
      fill="url(#goldBody)"
    />
    <path
      d="M111 26 C124 26 127 43 117 50 C114 52 110 49 113 46 C119 42 119 32 111 32 Z"
      fill="url(#goldBody)"
    />

    {/* Mangkuk Piala Utama */}
    <path
      d="M72 22 L113 22 C112 46 104 61 92.5 64 C81 61 73 46 72 22 Z"
      fill="url(#cupGold)"
      stroke="#b45309"
      strokeWidth="1"
    />
    {/* Bibir Atas Piala */}
    <ellipse cx="92.5" cy="22" rx="20.5" ry="4.5" fill="#fef08a" stroke="#b45309" strokeWidth="1" />
    <ellipse cx="92.5" cy="22" rx="16.5" ry="2.8" fill="#a16207" />

    {/* Tangkai & Alas Piala */}
    <rect x="89" y="63" width="7" height="14" rx="2" fill="url(#cupGold)" />
    <path d="M80 77 L105 77 L108 84 L77 84 Z" fill="url(#cupGold)" stroke="#92400e" strokeWidth="0.8" />
    <rect x="74" y="84" width="37" height="7" rx="1.5" fill="url(#goldBody)" stroke="#78350f" strokeWidth="0.8" />

    {/* Kilau Cahaya pada Piala */}
    <ellipse cx="83" cy="35" rx="3" ry="8" fill="#ffffff" opacity="0.45" transform="rotate(-12 83 35)" />
  </svg>
);

export const PiagamSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    printTargetStudentId,
    setPrintTargetStudentId,
    updateSekolah,
    setActiveMenu,
    showSaveNotification
  } = appContext;

  const importInputRef = useRef<HTMLInputElement | null>(null);

  const studentsInClass = db.siswa.filter(
    (s) => s.kelas === selectedKelas && s.statusSiswa === 'Aktif'
  );

  // Hitung rata-rata nilai tiap siswa untuk menentukan ranking kelas
  const studentAverages = studentsInClass.map((s) => {
    const scores = db.nilaiAkademik.filter(
      (n) =>
        matchesStudentScore(n, s) &&
        n.kelas === selectedKelas &&
        n.semester === selectedSemester &&
        n.nilaiAkhir > 0
    );
    const total = scores.reduce((acc, curr) => acc + curr.nilaiAkhir, 0);
    const avg = scores.length > 0 ? total / scores.length : 0;
    return {
      student: s,
      avg: parseFloat(avg.toFixed(1)),
      total
    };
  });

  const rankedStudents = [...studentAverages].sort((a, b) => b.avg - a.avg);

  const [selectedStudentId, setSelectedStudentId] = useState<string>(() => {
    if (filteredStudent && studentsInClass.some((s) => s.idSiswa === filteredStudent.idSiswa)) {
      return filteredStudent.idSiswa;
    }
    if (printTargetStudentId && studentsInClass.some((s) => s.idSiswa === printTargetStudentId)) {
      return printTargetStudentId;
    }
    return rankedStudents[0]?.student.idSiswa || studentsInClass[0]?.idSiswa || '';
  });

  useEffect(() => {
    if (filteredStudent && studentsInClass.some((s) => s.idSiswa === filteredStudent.idSiswa)) {
      setSelectedStudentId(filteredStudent.idSiswa);
      if (printTargetStudentId !== filteredStudent.idSiswa) {
        setPrintTargetStudentId(filteredStudent.idSiswa);
      }
      return;
    }
    if (printTargetStudentId && studentsInClass.some((s) => s.idSiswa === printTargetStudentId)) {
      setSelectedStudentId(printTargetStudentId);
      return;
    }
    const exists = studentsInClass.some((s) => s.idSiswa === selectedStudentId);
    if (!exists) {
      setSelectedStudentId(rankedStudents[0]?.student.idSiswa || studentsInClass[0]?.idSiswa || '');
    }
  }, [db.siswa, selectedKelas, filteredStudent, printTargetStudentId]);

  const currentIndex = Math.max(
    0,
    rankedStudents.findIndex((r) => r.student.idSiswa === selectedStudentId)
  );
  const currentRankNumber = rankedStudents.length > 0 ? currentIndex + 1 : 1;
  const currentRanked = rankedStudents[currentIndex];
  const currentStudent = currentRanked?.student || studentsInClass[0] || db.siswa[0];

  // Mode Protect / Unprotect (saat Unprotect, KOP & teks piagam bisa diedit manual langsung)
  const [isProtected, setIsProtected] = useState<boolean>(true);
  const [showSettingsModal, setShowSettingsModal] = useState<boolean>(false);

  // KOP & Teks Piagam yang dapat diedit manual saat Unprotect
  const defaultKop1 = db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN';
  const defaultKop2 =
    db.sekolah.namaInstansi && db.sekolah.namaSekolah
      ? `${db.sekolah.namaInstansi} • ${db.sekolah.namaSekolah}`
      : db.sekolah.namaInstansi || 'KOORDINATOR WILAYAH DIKPORA KECAMATAN';
  const defaultKopTlp = db.sekolah.alamat
    ? `Alamat: ${db.sekolah.alamat}${db.sekolah.telepon ? ` • Tlp. ${db.sekolah.telepon}` : ''}`
    : db.sekolah.telepon
    ? `Tlp. ${db.sekolah.telepon}`
    : 'Tlp';

  const [kopLine1, setKopLine1] = useState<string>(defaultKop1);
  const [kopLine2, setKopLine2] = useState<string>(defaultKop2);
  const [kopTlp, setKopTlp] = useState<string>(defaultKopTlp);
  const [titleText, setTitleText] = useState<string>('Piagam Penghargaan');
  const [diberikanKepadaText, setDiberikanKepadaText] = useState<string>('Diberikan Kepada');
  const [customRecipientName, setCustomRecipientName] = useState<string>('');
  const [sebagaiText, setSebagaiText] = useState<string>('Sebagai');
  const [selectedPresetId, setSelectedPresetId] = useState<string>('peringkat_otomatis');
  const [customRankLabel, setCustomRankLabel] = useState<string>('');
  const [verseLine1, setVerseLine1] = useState<string>(
    'Semoga Prestasi Yang Diraih Menjadi Motivasi Untuk Meraih Kesuksesan'
  );
  const [verseLine2, setVerseLine2] = useState<string>('Di Masa Yang Akan Datang');

  // Pengaturan Font, Logo, & Tanda Tangan Kanan
  const [certificateFont, setCertificateFont] = useState<string>(FONT_OPTIONS[0].value);
  const [fontScale, setFontScale] = useState<number>(100);
  const [showLeftLogo, setShowLeftLogo] = useState<boolean>(true);
  const [showRightLogo, setShowRightLogo] = useState<boolean>(false);
  const [showStudentSubInfo, setShowStudentSubInfo] = useState<boolean>(true);

  const [tempatTanggalInput, setTempatTanggalInput] = useState<string>(
    db.sekolah.tempatTanggalRapor ||
      (db.sekolah.tempatRapor && db.sekolah.tanggalRapor
        ? `${db.sekolah.tempatRapor}, ${db.sekolah.tanggalRapor}`
        : db.sekolah.tanggalRapor || '')
  );
  const [pengesahanOleh, setPengesahanOleh] = useState<string>('Kepala Sekolah');

  useEffect(() => {
    if (db.sekolah.tempatTanggalRapor) {
      setTempatTanggalInput(db.sekolah.tempatTanggalRapor);
    }
  }, [db.sekolah.tempatTanggalRapor]);

  const displayedRecipientName =
    customRecipientName.trim().length > 0
      ? customRecipientName
      : currentStudent?.namaSiswa || '';

  const displayedRankLabel =
    customRankLabel.trim().length > 0
      ? customRankLabel
      : getRankText(currentRankNumber);

  const handleSelectRankIndex = (nextOneBasedIndex: number) => {
    if (rankedStudents.length === 0) return;
    const bounded = Math.max(1, Math.min(rankedStudents.length, nextOneBasedIndex));
    const target = rankedStudents[bounded - 1];
    if (target) {
      setSelectedStudentId(target.student.idSiswa);
      setPrintTargetStudentId(target.student.idSiswa);
      setCustomRecipientName('');
      if (selectedPresetId === 'peringkat_otomatis') {
        setCustomRankLabel('');
      }
    }
  };

  const handleTempatTanggalChange = (val: string) => {
    setTempatTanggalInput(val);
    updateSekolah({
      ...db.sekolah,
      tempatTanggalRapor: val
    });
  };

  const handleUploadLogoKanan = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result as string;
      updateSekolah({
        ...db.sekolah,
        logoKananUrl: result
      });
      showSaveNotification('Logo Tut Wuri Handayani berhasil diganti!');
    };
    reader.readAsDataURL(file);
  };

  const handleUploadLogoKiri = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result as string;
      updateSekolah({
        ...db.sekolah,
        logoKiriUrl: result
      });
      showSaveNotification('Logo Kabupaten berhasil diganti!');
    };
    reader.readAsDataURL(file);
  };

  const handleSelectPreset = (preset: AwardPreset) => {
    setSelectedPresetId(preset.id);
    setCustomRankLabel(preset.rankTitle);
    setVerseLine1(preset.verseLine1);
    setVerseLine2(preset.verseLine2);
    if (preset.id === 'juara1' && rankedStudents[0]) {
      handleSelectRankIndex(1);
    } else if (preset.id === 'juara2' && rankedStudents[1]) {
      handleSelectRankIndex(2);
    } else if (preset.id === 'juara3' && rankedStudents[2]) {
      handleSelectRankIndex(3);
    }
  };

  const handleResetKopFromSchool = () => {
    setKopLine1(defaultKop1);
    setKopLine2(defaultKop2);
    setKopTlp(defaultKopTlp);
    setCustomRecipientName('');
    setCustomRankLabel('');
    setSelectedPresetId('peringkat_otomatis');
    showSaveNotification('KOP & teks piagam disinkronkan ulang dengan Identitas Sekolah!');
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [kopLine1],
      [kopLine2],
      [kopTlp],
      [titleText],
      [diberikanKepadaText],
      [displayedRecipientName || '-'],
      [`NISN: ${currentStudent?.nisn || '-'} | Kelas: ${selectedKelas}`],
      [sebagaiText],
      [displayedRankLabel],
      [verseLine1],
      [verseLine2],
      [''],
      ['Mengetahui / ' + pengesahanOleh, db.sekolah.namaKepalaSekolah || '-', `NIP. ${db.sekolah.nipKepalaSekolah || '-'}`],
      [tempatTanggalInput || '-', 'Wali Kelas: ' + (db.sekolah.waliKelas || '-'), `NIP. ${db.sekolah.nipWaliKelas || '-'}`]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Piagam');
    XLSX.writeFile(wb, `piagam-${displayedRecipientName || 'siswa'}.xlsx`);
    showSaveNotification('Data Piagam berhasil diekspor ke Excel!');
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="bg-[#00e5ff] min-h-[calc(100vh-5rem)] p-2.5 sm:p-4 rounded-xl print:bg-white print:p-0">
      {/* Aturan Cetak A4 Landscape */}
      <style>{`
        @media print {
          @page {
            size: A4 landscape;
            margin: 5mm;
          }
          .piagam-printable-area {
            width: 100% !important;
            max-width: 100% !important;
            min-height: 194mm !important;
            box-shadow: none !important;
            margin: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
        }
      `}</style>

      <div className="flex flex-col xl:flex-row items-start justify-between gap-4">
        {/* KIRI: LEMBAR PIAGAM PUTIH A4 LANDSCAPE */}
        <div className="w-full xl:flex-1 overflow-x-auto pb-2 print:overflow-visible print:pb-0">
          <div
            className="w-[297mm] min-h-[202mm] bg-white relative shadow-[0_10px_28px_rgba(0,0,0,0.22)] mx-auto printable-area piagam-printable-area text-black overflow-visible"
            style={{
              fontFamily: certificateFont,
              fontSize: `${fontScale}%`,
              WebkitPrintColorAdjust: 'exact',
              printColorAdjust: 'exact'
            }}
          >
            {/* Bingkai Ganda Biru Cyan */}
            <div className="absolute inset-[14px] border-[3.5px] border-[#0099dd] pointer-events-none z-10" />
            <div className="absolute inset-[21px] border-[2.5px] border-[#0099dd] pointer-events-none z-10" />

            {/* Ornamen Sudut Geometri Navy + Oranye (Kiri Atas & Kanan Bawah) */}
            <CertificateCornerOrnament position="top-left" />
            <CertificateCornerOrnament position="bottom-right" />

            {/* Konten Utama Piagam */}
            <div className="relative z-20 px-12 sm:px-16 pt-9 pb-8 min-h-[202mm] flex flex-col justify-between">
              {/* HEADER KOP + LOGO */}
              <div className="relative">
                {/* Logo Kiri (Tut Wuri Handayani) */}
                {showLeftLogo && (
                  <div className="absolute left-2 sm:left-3 top-5 z-20 flex items-center justify-center">
                    <LogoTutWuriHandayani
                      url={db.sekolah.logoKananUrl}
                      size={84}
                      className="w-[78px] h-[78px] sm:w-[86px] sm:h-[86px]"
                      label="Logo Tut Wuri Handayani"
                    />
                  </div>
                )}

                {/* Logo Kanan Opsional (Pemda Kabupaten) */}
                {showRightLogo && (
                  <div className="absolute right-2 sm:right-3 top-5 z-20 flex items-center justify-center">
                    <LogoKabupaten
                      url={db.sekolah.logoKiriUrl}
                      size={82}
                      className="w-[76px] h-[76px] sm:w-[84px] sm:h-[84px]"
                      label="Logo Kabupaten"
                    />
                  </div>
                )}

                {/* Baris Teks Kop Tengah */}
                <div className={`text-center ${showLeftLogo ? 'px-20 sm:px-24' : 'px-8'} pt-2`}>
                  {isProtected ? (
                    <div className="text-[15px] sm:text-[18px] font-bold uppercase tracking-wide text-black leading-snug">
                      {kopLine1}
                    </div>
                  ) : (
                    <input
                      type="text"
                      value={kopLine1}
                      onChange={(e) => setKopLine1(e.target.value)}
                      className="w-full text-center text-[15px] sm:text-[18px] font-bold uppercase tracking-wide text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 focus:outline-none"
                    />
                  )}

                  {isProtected ? (
                    <div className="text-[15px] sm:text-[18px] font-bold uppercase tracking-wide text-black leading-snug mt-2">
                      {kopLine2}
                    </div>
                  ) : (
                    <input
                      type="text"
                      value={kopLine2}
                      onChange={(e) => setKopLine2(e.target.value)}
                      className="w-full text-center text-[15px] sm:text-[18px] font-bold uppercase tracking-wide text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 mt-2 focus:outline-none"
                    />
                  )}

                  {isProtected ? (
                    <div className="text-[12px] sm:text-[13px] italic font-bold text-black mt-4">
                      {kopTlp}
                    </div>
                  ) : (
                    <input
                      type="text"
                      value={kopTlp}
                      onChange={(e) => setKopTlp(e.target.value)}
                      className="w-3/4 mx-auto block text-center text-[12px] sm:text-[13px] italic font-bold text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 mt-4 focus:outline-none"
                    />
                  )}
                </div>

                {/* Garis Pembatas KOP dengan Shadow Halus */}
                <div
                  className="w-[84%] mx-auto mt-1 h-[2.5px] bg-[#64748b]"
                  style={{ boxShadow: '0 2px 3px rgba(0,0,0,0.35)' }}
                />
              </div>

              {/* ISI TENGAH SERTIFIKAT */}
              <div className="text-center my-auto py-3 px-4 flex flex-col items-center justify-center">
                {/* Judul Kaligrafi "Piagam Penghargaan" */}
                {isProtected ? (
                  <h1
                    className="text-[38px] sm:text-[46px] text-[#0066ff] leading-tight tracking-wide font-normal"
                    style={{
                      fontFamily: '"UnifrakturMaguntia", "Old English Text MT", "EB Garamond", Georgia, serif'
                    }}
                  >
                    {titleText}
                  </h1>
                ) : (
                  <input
                    type="text"
                    value={titleText}
                    onChange={(e) => setTitleText(e.target.value)}
                    className="w-full max-w-xl text-center text-[36px] sm:text-[42px] text-[#0066ff] bg-amber-50/70 border border-dashed border-sky-500 rounded px-3 py-0.5 focus:outline-none"
                    style={{
                      fontFamily: '"UnifrakturMaguntia", "Old English Text MT", "EB Garamond", Georgia, serif'
                    }}
                  />
                )}

                {/* "Diberikan Kepada" */}
                {isProtected ? (
                  <p className="text-[14px] sm:text-[15px] font-bold text-black mt-1">
                    {diberikanKepadaText}
                  </p>
                ) : (
                  <input
                    type="text"
                    value={diberikanKepadaText}
                    onChange={(e) => setDiberikanKepadaText(e.target.value)}
                    className="text-center text-[14px] sm:text-[15px] font-bold text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 mt-1 focus:outline-none"
                  />
                )}

                {/* Nama Siswa Penerima */}
                <div className="my-2.5 min-h-[42px] flex flex-col items-center justify-center">
                  {isProtected ? (
                    <div className="text-[22px] sm:text-[27px] font-extrabold uppercase tracking-wider text-slate-950 border-b-2 border-slate-800/80 px-8 pb-0.5">
                      {displayedRecipientName || '—'}
                    </div>
                  ) : (
                    <input
                      type="text"
                      value={customRecipientName || displayedRecipientName}
                      onChange={(e) => setCustomRecipientName(e.target.value)}
                      placeholder="Nama Lengkap Siswa"
                      className="text-center text-[22px] sm:text-[26px] font-extrabold uppercase tracking-wider text-slate-950 bg-amber-50/70 border border-dashed border-sky-500 rounded px-6 py-0.5 focus:outline-none"
                    />
                  )}

                  {showStudentSubInfo && currentStudent && (
                    <div className="text-[11px] font-semibold text-slate-600 mt-1 tracking-wide">
                      NISN: {currentStudent.nisn || '-'} • Kelas {selectedKelas} • Semester {selectedSemester} ({db.sekolah.tahunPelajaran})
                    </div>
                  )}
                </div>

                {/* "Sebagai" */}
                {isProtected ? (
                  <p className="text-[14px] sm:text-[15px] font-bold text-black mt-1">
                    {sebagaiText}
                  </p>
                ) : (
                  <input
                    type="text"
                    value={sebagaiText}
                    onChange={(e) => setSebagaiText(e.target.value)}
                    className="text-center text-[14px] sm:text-[15px] font-bold text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 mt-1 focus:outline-none"
                  />
                )}

                {/* "Peringkat Pertama" */}
                {isProtected ? (
                  <div className="text-[24px] sm:text-[30px] font-extrabold text-[#0066ff] tracking-tight leading-tight mt-1">
                    {displayedRankLabel}
                  </div>
                ) : (
                  <input
                    type="text"
                    value={customRankLabel || displayedRankLabel}
                    onChange={(e) => setCustomRankLabel(e.target.value)}
                    className="text-center text-[24px] sm:text-[28px] font-extrabold text-[#0066ff] bg-amber-50/70 border border-dashed border-sky-500 rounded px-4 py-0.5 mt-1 focus:outline-none"
                  />
                )}

                {/* 2 Baris Motivasi */}
                <div className="mt-5 space-y-0.5 max-w-2xl w-full">
                  {isProtected ? (
                    <>
                      <p className="text-[14px] sm:text-[16px] text-black leading-snug">
                        {verseLine1}
                      </p>
                      <p className="text-[14px] sm:text-[16px] text-black leading-snug">
                        {verseLine2}
                      </p>
                    </>
                  ) : (
                    <>
                      <input
                        type="text"
                        value={verseLine1}
                        onChange={(e) => setVerseLine1(e.target.value)}
                        className="w-full text-center text-[14px] sm:text-[15px] text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 focus:outline-none"
                      />
                      <input
                        type="text"
                        value={verseLine2}
                        onChange={(e) => setVerseLine2(e.target.value)}
                        className="w-full text-center text-[14px] sm:text-[15px] text-black bg-amber-50/70 border border-dashed border-sky-500 rounded px-2 py-0.5 focus:outline-none"
                      />
                    </>
                  )}
                </div>
              </div>

              {/* FOOTER: TANDA TANGAN KIRI, TROFI TENGAH, WALI KELAS KANAN */}
              <div className="grid grid-cols-3 items-end gap-4 px-4 sm:px-10 pt-2">
                {/* Kiri: Mengetahui / Pengesahan */}
                <div className="text-left text-[13px] sm:text-[14px] text-black leading-snug pl-2 sm:pl-6">
                  <div>Mengetahui</div>
                  <div>{pengesahanOleh || 'Kepala Sekolah'}</div>
                  <div className="h-14 sm:h-16" />
                  {db.sekolah.namaKepalaSekolah && (
                    <div className="font-bold underline">{db.sekolah.namaKepalaSekolah}</div>
                  )}
                  <div>NIP. {db.sekolah.nipKepalaSekolah || ''}</div>
                </div>

                {/* Tengah: Ilustrasi Trofi Emas */}
                <div className="flex items-center justify-center pb-1">
                  <GoldTrophyIllustration />
                </div>

                {/* Kanan: Tempat Tanggal + Wali Kelas */}
                <div className="text-left text-[13px] sm:text-[14px] text-black leading-snug pl-6 sm:pl-14">
                  <div className="min-h-[20px]">
                    {tempatTanggalInput.trim() !== '' ? tempatTanggalInput : '\u00A0'}
                  </div>
                  <div>Wali Kelas</div>
                  <div className="h-14 sm:h-16" />
                  {db.sekolah.waliKelas && (
                    <div className="font-bold underline">{db.sekolah.waliKelas}</div>
                  )}
                  <div>NIP. {db.sekolah.nipWaliKelas || ''}</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* KANAN: PANEL KONTROL HITAM + PETUNJUK + INPUT TANDA TANGAN (NO-PRINT) */}
        <aside className="w-full xl:w-[270px] shrink-0 flex flex-col justify-between gap-6 no-print print:hidden">
          {/* Kotak Kontrol Hitam Berbingkai Merah */}
          <div className="bg-black border-[2.5px] border-[#ff0000] p-2 space-y-2.5 shadow-xl">
            {/* Baris Tombol Atas */}
            <div className="flex items-center justify-between gap-1">
              {/* Pemilih Nomor Peringkat / Siswa [1] ▲ ▼ */}
              <div className="flex items-stretch shrink-0" title="Ganti Urutan Siswa / Peringkat">
                <div className="w-9 h-9 bg-[#0066ff] border border-white/80 flex items-center justify-center font-black text-black text-lg select-none">
                  {currentRankNumber}
                </div>
                <div className="flex flex-col w-5 h-9 border border-slate-400 bg-slate-200">
                  <button
                    type="button"
                    onClick={() => handleSelectRankIndex(currentRankNumber + 1)}
                    disabled={currentRankNumber >= rankedStudents.length}
                    title="Peringkat Berikutnya"
                    className="flex-1 flex items-center justify-center hover:bg-slate-300 disabled:opacity-40 text-black border-b border-slate-400 cursor-pointer"
                  >
                    <ChevronUp size={12} strokeWidth={3} />
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSelectRankIndex(currentRankNumber - 1)}
                    disabled={currentRankNumber <= 1}
                    title="Peringkat Sebelumnya"
                    className="flex-1 flex items-center justify-center hover:bg-slate-300 disabled:opacity-40 text-black cursor-pointer"
                  >
                    <ChevronDown size={12} strokeWidth={3} />
                  </button>
                </div>
              </div>

              {/* Tombol Home Bulat Biru */}
              <button
                type="button"
                onClick={() => setActiveMenu('dashboard')}
                title="Kembali ke Menu Utama"
                className="w-9 h-9 rounded-full bg-[#1d4ed8] hover:bg-[#2563eb] border-2 border-[#38bdf8] flex items-center justify-center text-black shadow-md cursor-pointer shrink-0 transition-transform active:scale-95"
              >
                <Home size={18} strokeWidth={2.5} />
              </button>

              {/* Tombol Cyan ATUR HALAMAN */}
              <button
                type="button"
                onClick={() => setShowSettingsModal(true)}
                className="px-2 h-9 bg-[#00e5ff] hover:bg-[#40ecff] text-black font-black text-[10px] leading-[1.08] uppercase border border-white/90 shadow-inner cursor-pointer text-center shrink-0"
              >
                ATUR
                <br />
                HALAMAN
              </button>

              {/* Tombol Bertumpuk Protect / Unprotect */}
              <div className="flex flex-col gap-0.5 flex-1 min-w-[66px]">
                <button
                  type="button"
                  onClick={() => {
                    setIsProtected(true);
                    showSaveNotification('Mode Protect aktif: Teks Piagam dikunci siap cetak.');
                  }}
                  className={`w-full text-[10px] font-bold py-0.5 px-1.5 border cursor-pointer transition-all flex items-center justify-center gap-1 ${
                    isProtected
                      ? 'bg-[#ff0000] text-white border-white shadow-xs'
                      : 'bg-[#7f1d1d] text-white/80 border-slate-500 hover:bg-[#dc2626]'
                  }`}
                >
                  <Lock size={9} />
                  <span>Protect</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setIsProtected(false);
                    showSaveNotification('Mode Unprotect aktif: Klik langsung Kop atau teks di piagam untuk edit manual.');
                  }}
                  className={`w-full text-[10px] font-bold py-0.5 px-1.5 border cursor-pointer transition-all flex items-center justify-center gap-1 ${
                    !isProtected
                      ? 'bg-[#fef08a] text-black border-amber-400 shadow-xs'
                      : 'bg-white text-[#991b1b] border-slate-400 hover:bg-slate-100'
                  }`}
                >
                  <Unlock size={9} />
                  <span>Unprotect</span>
                </button>
              </div>
            </div>

            {/* Tombol Tengah Ganti Logo, Font & Cetak */}
            <div className="flex flex-col items-center gap-1.5">
              <button
                type="button"
                onClick={() => setShowSettingsModal(true)}
                className="w-[78%] py-1.5 px-3 bg-gradient-to-b from-white via-slate-100 to-slate-300 hover:from-slate-50 hover:to-slate-200 text-black font-extrabold italic text-[10.5px] leading-tight border border-slate-400 shadow-md cursor-pointer text-center"
              >
                GANTI LOGO, FONT DLL
              </button>

              <div className="w-full grid grid-cols-2 gap-1.5 pt-0.5">
                <button
                  type="button"
                  onClick={handlePrint}
                  className="flex items-center justify-center gap-1 py-1.5 px-2 bg-[#0066ff] hover:bg-[#1d4ed8] text-white font-bold text-[10.5px] border border-sky-300 cursor-pointer"
                >
                  <Printer size={12} />
                  <span>Cetak A4</span>
                </button>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="flex items-center justify-center gap-1 py-1.5 px-2 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-[10.5px] border border-emerald-300 cursor-pointer"
                >
                  <Download size={12} />
                  <span>Excel</span>
                </button>
              </div>
            </div>

            {/* Kotak Putih Petunjuk */}
            <div className="bg-white text-black p-2.5 text-[11px] leading-[1.35] border border-slate-400">
              <div className="font-semibold mb-1">Petunjuk :</div>
              <ol className="list-decimal pl-4 space-y-0.5">
                <li>Cetak dengan kertas A4 (210x297)</li>
                <li>Kop. dapat di edit manual</li>
                <li>Format huruf dapat diganti</li>
                <li>Gunakan printer Epson untuk hasil cetakan yang lebih baik</li>
                <li>
                  Jika menggunakan printer selain epson, atur terlebih dahulu ukuran lebar kolom (ceklis Pada Headings)
                </li>
              </ol>
              <p className="text-[#ff0000] font-bold mt-2 leading-tight">
                Pada printer selain Epson hanya tersedia ukuran kertas A4 (tidak tersedia ukuran kertas A4(210x297))
              </p>
            </div>
          </div>

          {/* Info Ringkas Siswa Terpilih di Panel Kanan */}
          <div className="bg-white/90 border-2 border-[#005a9e] p-2.5 text-[11px] text-slate-900 shadow-sm space-y-1.5">
            <div className="font-bold text-[#005a9e] uppercase tracking-wider text-[10px]">
              Siswa Terpilih (Peringkat #{currentRankNumber})
            </div>
            <select
              value={currentStudent?.idSiswa || ''}
              onChange={(e) => {
                const idx = rankedStudents.findIndex((r) => r.student.idSiswa === e.target.value);
                if (idx >= 0) handleSelectRankIndex(idx + 1);
              }}
              className="w-full px-2 py-1.5 border border-slate-400 bg-white text-xs font-bold text-slate-900 focus:outline-none focus:border-[#0066ff]"
            >
              {rankedStudents.map((r, idx) => (
                <option key={r.student.idSiswa} value={r.student.idSiswa}>
                  #{idx + 1} - {r.student.namaSiswa} ({r.avg})
                </option>
              ))}
            </select>
          </div>

          {/* Bagian Bawah Kanan: Ganti Tempat dan Tanggal + Pengesahan Oleh */}
          <div className="space-y-2.5 mt-auto pt-4">
            <div>
              <label className="block bg-[#005a9e] text-white font-serif italic font-bold text-[12px] px-2 py-0.5">
                Ganti Tempat dan Tanggal
              </label>
              <input
                type="text"
                value={tempatTanggalInput}
                onChange={(e) => handleTempatTanggalChange(e.target.value)}
                placeholder="Contoh: Balil, 23 Desember 2026"
                className="w-full bg-white border-2 border-[#ff0000] px-2 py-1 text-[12px] font-semibold text-black focus:outline-none"
              />
            </div>

            <div>
              <label className="block bg-[#005a9e] text-white font-serif italic font-bold text-[12px] px-2 py-0.5">
                Pengesahan Oleh
              </label>
              <input
                type="text"
                list="pengesahan-options"
                value={pengesahanOleh}
                onChange={(e) => setPengesahanOleh(e.target.value)}
                placeholder="Kepala Sekolah"
                className="w-full bg-white border-2 border-[#ff0000] px-2 py-1 text-[12px] font-semibold text-black focus:outline-none"
              />
              <datalist id="pengesahan-options">
                <option value="Kepala Sekolah" />
                <option value="Plt. Kepala Sekolah" />
                <option value="Komite Sekolah" />
              </datalist>
            </div>
          </div>
        </aside>
      </div>

      {/* MODAL ATUR HALAMAN, GANTI LOGO, FONT & IMPOR/EKSPOR */}
      {showSettingsModal && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4 print:hidden">
          <div className="bg-white rounded-xl max-w-2xl w-full shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between px-5 py-3.5 bg-[#005a9e] text-white">
              <div className="flex items-center gap-2">
                <SlidersHorizontal size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wide">
                  Atur Halaman, Font, Logo & Kategori Piagam
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowSettingsModal(false)}
                className="p-1 rounded hover:bg-white/20 cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-5 text-xs text-slate-800">
              {/* Format Huruf & Skala Font */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Format Huruf Piagam
                  </label>
                  <select
                    value={certificateFont}
                    onChange={(e) => setCertificateFont(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg font-medium bg-white"
                  >
                    {FONT_OPTIONS.map((f) => (
                      <option key={f.label} value={f.value}>
                        {f.label}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Ukuran Huruf ({fontScale}%)
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="range"
                      min={85}
                      max={115}
                      step={5}
                      value={fontScale}
                      onChange={(e) => setFontScale(Number(e.target.value))}
                      className="flex-1 accent-[#0066ff]"
                    />
                    <button
                      type="button"
                      onClick={() => setFontScale(100)}
                      className="px-2 py-1 border rounded text-[11px] hover:bg-slate-50 cursor-pointer"
                    >
                      100%
                    </button>
                  </div>
                </div>
              </div>

              {/* Kategori / Teks Penghargaan */}
              <div>
                <label className="block font-bold text-slate-700 mb-1.5">
                  Kategori Penghargaan & Predikat
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {AWARD_PRESETS.map((preset) => (
                    <button
                      key={preset.id}
                      type="button"
                      onClick={() => handleSelectPreset(preset)}
                      className={`text-left px-3 py-2 rounded-lg border cursor-pointer transition-all ${
                        selectedPresetId === preset.id
                          ? 'bg-sky-50 border-[#0066ff] text-sky-950 font-bold'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      {preset.label}
                    </button>
                  ))}
                </div>
              </div>

              {/* Pengaturan Logo */}
              <div className="space-y-3 border-t pt-4">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <span className="font-bold text-slate-800 uppercase text-[11px]">
                    Pengaturan Logo Piagam
                  </span>
                  <div className="flex flex-wrap items-center gap-3">
                    <label className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={showLeftLogo}
                        onChange={(e) => setShowLeftLogo(e.target.checked)}
                      />
                      <span>Logo Kiri (Tut Wuri)</span>
                    </label>
                    <label className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={showRightLogo}
                        onChange={(e) => setShowRightLogo(e.target.checked)}
                      />
                      <span>Logo Kanan (Pemda Kab.)</span>
                    </label>
                    <label className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={showStudentSubInfo}
                        onChange={(e) => setShowStudentSubInfo(e.target.checked)}
                      />
                      <span>Tampilkan NISN & Kelas</span>
                    </label>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {/* Logo Tut Wuri Handayani */}
                  <div className="flex items-center gap-3 p-2.5 rounded-lg border border-slate-200 bg-slate-50">
                    <div className="w-12 h-12 bg-white border rounded flex items-center justify-center shrink-0 overflow-hidden">
                      <LogoTutWuriHandayani url={db.sekolah.logoKananUrl} size={42} className="w-10 h-10" />
                    </div>
                    <div className="flex-1 space-y-1">
                      <div className="font-bold flex items-center gap-1">
                        <span>Logo Tut Wuri Handayani</span>
                        {db.sekolah.logoKananUrl && <Check size={12} className="text-emerald-600" />}
                      </div>
                      <div className="flex items-center gap-2">
                        <label className="inline-flex items-center gap-1 px-2 py-1 bg-sky-100 hover:bg-sky-200 text-sky-900 font-bold rounded cursor-pointer text-[10px]">
                          <Upload size={11} />
                          <span>Ganti Logo</span>
                          <input type="file" accept="image/*" onChange={handleUploadLogoKanan} className="hidden" />
                        </label>
                        {db.sekolah.logoKananUrl && (
                          <button
                            type="button"
                            onClick={() => updateSekolah({ ...db.sekolah, logoKananUrl: '' })}
                            className="text-[10px] text-rose-600 font-semibold cursor-pointer"
                          >
                            Reset
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Logo Kabupaten */}
                  <div className="flex items-center gap-3 p-2.5 rounded-lg border border-slate-200 bg-slate-50">
                    <div className="w-12 h-12 bg-white border rounded flex items-center justify-center shrink-0 overflow-hidden">
                      <LogoKabupaten url={db.sekolah.logoKiriUrl} size={42} className="w-10 h-10" />
                    </div>
                    <div className="flex-1 space-y-1">
                      <div className="font-bold flex items-center gap-1">
                        <span>Logo Kabupaten / Pemda</span>
                        {db.sekolah.logoKiriUrl && <Check size={12} className="text-emerald-600" />}
                      </div>
                      <div className="flex items-center gap-2">
                        <label className="inline-flex items-center gap-1 px-2 py-1 bg-indigo-100 hover:bg-indigo-200 text-indigo-900 font-bold rounded cursor-pointer text-[10px]">
                          <Upload size={11} />
                          <span>Ganti Logo</span>
                          <input type="file" accept="image/*" onChange={handleUploadLogoKiri} className="hidden" />
                        </label>
                        {db.sekolah.logoKiriUrl && (
                          <button
                            type="button"
                            onClick={() => updateSekolah({ ...db.sekolah, logoKiriUrl: '' })}
                            className="text-[10px] text-rose-600 font-semibold cursor-pointer"
                          >
                            Reset
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Impor / Sinkronisasi KOP */}
              <div className="flex flex-wrap items-center justify-between gap-2 border-t pt-4">
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={handleResetKopFromSchool}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-50 font-semibold cursor-pointer"
                  >
                    <RotateCcw size={13} />
                    <span>Sinkronkan KOP Sekolah</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => importInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-emerald-300 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-semibold cursor-pointer"
                  >
                    <FileSpreadsheet size={13} />
                    <span>Impor Excel Piagam</span>
                  </button>
                  <input
                    ref={importInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      try {
                        const res = await parseAndImportTemplate(file, 'template_piagam', appContext);
                        showSaveNotification(res.success ? res.message : `Gagal Impor: ${res.message}`);
                      } catch (err: unknown) {
                        const msg = err instanceof Error ? err.message : 'Format tidak valid';
                        showSaveNotification('Gagal mengimpor file piagam: ' + msg);
                      }
                      e.target.value = '';
                    }}
                  />
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setShowSettingsModal(false);
                      setTimeout(handlePrint, 150);
                    }}
                    className="inline-flex items-center gap-1.5 px-4 py-1.5 rounded-lg bg-[#0066ff] hover:bg-[#1d4ed8] text-white font-bold cursor-pointer"
                  >
                    <Printer size={14} />
                    <span>Cetak Piagam A4</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowSettingsModal(false)}
                    className="px-4 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-900 text-white font-bold cursor-pointer"
                  >
                    Tutup
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
