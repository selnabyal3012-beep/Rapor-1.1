import React, { useState, useMemo, useEffect } from 'react';
import {
  FileText,
  Printer,
  ChevronLeft,
  ChevronRight,
  User,
  School,
  Calendar,
  MapPin,
  Edit3,
  CheckCircle2,
  X,
  Save,
  Users,
  Award,
  AlertCircle,
  Filter,
  CheckSquare,
  Square,
  BookOpen,
  Layers,
  Eye,
  EyeOff,
  Minimize2,
  Maximize2,
  ChevronDown,
  ChevronUp,
  Zap,
  ArrowUp,
  ArrowDown,
  FileCheck2,
  Image as ImageIcon
} from 'lucide-react';
import { useApp, PAPER_SIZE_MAP, getReportFontCssStack } from '../../context/AppContext';
import { ClassLevel, StudentDecision, SemesterType } from '../../types';
import { LogoKabupaten, LogoTutWuriHandayani } from '../common/OfficialLogos';
import { StandardActionToolbar } from '../common/StandardActionToolbar';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { ReportCoverPage } from '../common/ReportCoverPage';
import { ReportSchoolIdentityPage } from '../common/ReportSchoolIdentityPage';
import { ReportStudentIdentityPage } from '../common/ReportStudentIdentityPage';
import { PrintTypographyControl } from '../common/PrintTypographyControl';
import { generateReportDescriptionFromTP, getNamaPelajaranAgama } from '../../utils/tpDescriptionHelper';
import { matchesStudentScore, matchesStudentRecord } from '../../utils/studentSyncHelper';
import { ReportBottomSection } from '../print/ReportBottomSection';
import {
  ReportAcademicHeaderAndScores,
  ReportExtracurricularTable,
  splitAcademicScoresForPages
} from '../print/ReportAcademicTables';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import * as XLSX from 'xlsx';

function getNextClass(current: ClassLevel): string {
  switch (current) {
    case 'I':
      return 'II (Dua)';
    case 'II':
      return 'III (Tiga)';
    case 'III':
      return 'IV (Empat)';
    case 'IV':
      return 'V (Lima)';
    case 'V':
      return 'VI (Enam)';
    case 'VI':
      return 'Lulus dari Satuan Pendidikan SD';
    case 'VII':
      return 'VIII (Delapan)';
    case 'VIII':
      return 'IX (Sembilan)';
    case 'IX':
      return 'Lulus dari Satuan Pendidikan SMP';
    case 'X':
      return 'XI (Sebelas)';
    case 'XI':
      return 'XII (Dua Belas)';
    case 'XII':
      return 'Lulus dari Satuan Pendidikan SMA/SMK';
    case 'XIII':
      return 'Lulus dari Satuan Pendidikan SMK Program 4 Tahun';
    default:
      return 'Tingkat Berikutnya';
  }
}

function getCurrentClassText(current: ClassLevel): string {
  switch (current) {
    case 'I':
      return 'I (Satu)';
    case 'II':
      return 'II (Dua)';
    case 'III':
      return 'III (Tiga)';
    case 'IV':
      return 'IV (Empat)';
    case 'V':
      return 'V (Lima)';
    case 'VI':
      return 'VI (Enam)';
    case 'VII':
      return 'VII (Tujuh)';
    case 'VIII':
      return 'VIII (Delapan)';
    case 'IX':
      return 'IX (Sembilan)';
    case 'X':
      return 'X (Sepuluh)';
    case 'XI':
      return 'XI (Sebelas)';
    case 'XII':
      return 'XII (Dua Belas)';
    case 'XIII':
      return 'XIII (Tiga Belas)';
    default:
      return `${current}`;
  }
}

interface AcademicReportViewSectionProps {
  forcedSemester?: 'Ganjil' | 'Genap';
}

export const AcademicReportViewSection: React.FC<AcademicReportViewSectionProps> = ({ forcedSemester }) => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    setSelectedSemester,
    selectedFase,
    filteredStudent,
    printTargetStudentId,
    setPrintTargetStudentId,
    saveKeputusanSiswa,
    bulkSaveKeputusanSiswa,
    syncAllScoresWithTPDescriptions,
    showSaveNotification,
    openReportPreview,
    printFontFamily,
    printFontSize,
    setPrintFontSize,
    printLineHeight,
    printLogoScale,
    setPrintLogoScale,
    paperSize,
    paperOrientation,
    paperMargins,
    setPaperMargins,
    printScale,
    setPrintScale,
    updateSekolah
  } = appContext;

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

  const handleTempatTanggalChange = (val: string) => {
    setTempatTanggalInput(val);
    updateSekolah({
      ...db.sekolah,
      tempatTanggalRapor: val
    });
  };

  const adjustFontSize = (direction: 'up' | 'down') => {
    const cur = parseFloat(printFontSize) || 11;
    const next =
      direction === 'up'
        ? Math.min(24, Math.round((cur + (cur >= 16 ? 2 : cur >= 12 ? 1 : 0.5)) * 2) / 2)
        : Math.max(8, Math.round((cur - (cur > 16 ? 2 : cur > 12 ? 1 : 0.5)) * 2) / 2);
    const nextStr = `${next}pt`;
    setPrintFontSize(nextStr);
    showSaveNotification(`Ukuran tulisan rapor diubah ke ${nextStr}`);
  };

  const adjustPaperMargin = (side: 'top' | 'bottom' | 'left' | 'right', delta: number) => {
    const nextVal = Math.max(0, Math.min(60, (paperMargins[side] ?? 0) + delta));
    setPaperMargins({
      ...paperMargins,
      [side]: nextVal
    });
    const labels = { top: 'Atas', bottom: 'Bawah', left: 'Kiri', right: 'Kanan' };
    showSaveNotification(`Batas kertas ${labels[side]} diatur ke ${nextVal} mm`);
  };

  const adjustPaperScale = (delta: number) => {
    const next = Math.max(60, Math.min(140, (printScale || 100) + delta));
    setPrintScale(next);
  };

  const resetPaperLayout = () => {
    setPaperMargins({ top: 20, bottom: 20, left: 25, right: 20 });
    setPrintScale(100);
    showSaveNotification('Ukuran kertas & margin (Kiri/Kanan/Atas/Bawah) dikembalikan ke standar');
  };

  const paperSpec = PAPER_SIZE_MAP[paperSize] || PAPER_SIZE_MAP.F4;
  const finalPageCssSize =
    paperOrientation === 'landscape'
      ? `${paperSpec.heightMm}mm ${paperSpec.widthMm}mm`
      : `${paperSpec.widthMm}mm ${paperSpec.heightMm}mm`;
  const reportFontCss = getReportFontCssStack(printFontFamily);
  const sheetMaxWClass =
    paperOrientation === 'landscape'
      ? 'max-w-6xl report-paper-sheet--landscape'
      : 'max-w-4xl report-paper-sheet--portrait';

  const handleImportFile = async (file: File) => {
    try {
      const res = await parseAndImportTemplate(file, undefined, appContext);
      if (res.success) {
        showSaveNotification(res.message);
      } else {
        showSaveNotification(`Gagal Impor: ${res.message}`);
      }
    } catch (err: any) {
      showSaveNotification(`Gagal membaca file: ${err.message || 'Format tidak valid'}`);
    }
  };

  const activeSemester: SemesterType = forcedSemester || selectedSemester;

  useEffect(() => {
    if (
      filteredStudent &&
      filteredStudent.kelas === selectedKelas &&
      printTargetStudentId !== filteredStudent.idSiswa
    ) {
      setPrintTargetStudentId(filteredStudent.idSiswa);
    }
  }, [filteredStudent, selectedKelas, printTargetStudentId, setPrintTargetStudentId]);

  const siswaKelas = db.siswa.filter((s) => s.kelas === selectedKelas);
  const currentStudent = siswaKelas.find((s) => s.idSiswa === printTargetStudentId) || siswaKelas[0];

  const [openDecisionModal, setOpenDecisionModal] = useState(false);
  const [activePage, setActivePage] = useState<'kaver' | 'rapor_peserta_didik' | 'identitas' | 'rapor' | 'semua'>('semua');
  const [scrollBarMode, setScrollBarMode] = useState<'full' | 'compact' | 'hidden'>('full');

  // Cari keputusan siswa saat ini dari database
  const currentDecision = currentStudent
    ? db.keputusanSiswa?.find(
        (k) =>
          matchesStudentRecord(k, currentStudent) &&
          k.kelas === selectedKelas &&
          k.semester === activeSemester
      )
    : undefined;

  // Form state untuk modal keputusan
  const [decisionForm, setDecisionForm] = useState<{
    status: 'Naik Kelas' | 'Tinggal di Kelas' | 'Lulus' | 'Belum Ditetapkan';
    naikKelas: string;
    tinggalKelas: string;
    applyToAll: boolean;
  }>({
    status: currentDecision?.statusKeputusan || 'Naik Kelas',
    naikKelas: currentDecision?.naikKelas || getNextClass(selectedKelas),
    tinggalKelas: currentDecision?.tinggalKelas || '-',
    applyToAll: false
  });

  const handleOpenDecisionModal = () => {
    if (currentDecision) {
      setDecisionForm({
        status: currentDecision.statusKeputusan,
        naikKelas: currentDecision.naikKelas,
        tinggalKelas: currentDecision.tinggalKelas,
        applyToAll: false
      });
    } else {
      // Default naik kelas
      setDecisionForm({
        status: 'Naik Kelas',
        naikKelas: getNextClass(selectedKelas),
        tinggalKelas: '-',
        applyToAll: false
      });
    }
    setOpenDecisionModal(true);
  };

  // Filter Mata Pelajaran (Wajib / Pilihan)
  type MapelCategoryFilter = 'semua' | 'wajib' | 'pilihan';
  const [mapelCategoryFilter, setMapelCategoryFilter] = useState<MapelCategoryFilter>('semua');

  // Daftar mapel pilihan & muatan lokal yang ada di DB
  const mapelPilihanList = useMemo(() => {
    return db.mataPelajaran.filter((m) => m.kelompok === 'Pilihan' || m.kelompok === 'Muatan Lokal');
  }, [db.mataPelajaran]);

  // Set mata pelajaran pilihan yang aktif dicentang untuk ditampilkan di rapor
  const [activePilihanMapels, setActivePilihanMapels] = useState<string[]>(() => {
    return db.mataPelajaran
      .filter((m) => m.kelompok === 'Pilihan' || m.kelompok === 'Muatan Lokal')
      .map((m) => m.mataPelajaran);
  });

  const togglePilihanMapel = (namaMapel: string) => {
    setActivePilihanMapels((prev) =>
      prev.includes(namaMapel) ? prev.filter((m) => m !== namaMapel) : [...prev, namaMapel]
    );
  };

  const selectAllPilihan = () => {
    setActivePilihanMapels(mapelPilihanList.map((m) => m.mataPelajaran));
  };

  const clearAllPilihan = () => {
    setActivePilihanMapels([]);
  };

  // Filter Mata Pelajaran berdasarkan kategori dan pilihan aktif
  const filteredMataPelajaran = useMemo(() => {
    return db.mataPelajaran.filter((m) => {
      if (mapelCategoryFilter === 'wajib') {
        return m.kelompok === 'Wajib';
      }
      if (mapelCategoryFilter === 'pilihan') {
        return (m.kelompok === 'Pilihan' || m.kelompok === 'Muatan Lokal') && activePilihanMapels.includes(m.mataPelajaran);
      }
      // 'semua': Wajib selalu ada, Pilihan/Mulok ada jika aktif dicentang
      if (m.kelompok === 'Pilihan' || m.kelompok === 'Muatan Lokal') {
        return activePilihanMapels.includes(m.mataPelajaran);
      }
      return true;
    });
  }, [db.mataPelajaran, mapelCategoryFilter, activePilihanMapels]);

  const handleSelectStatus = (status: 'Naik Kelas' | 'Tinggal di Kelas' | 'Lulus' | 'Belum Ditetapkan') => {
    if (status === 'Naik Kelas') {
      setDecisionForm((prev) => ({
        ...prev,
        status,
        naikKelas: getNextClass(selectedKelas),
        tinggalKelas: '-'
      }));
    } else if (status === 'Tinggal di Kelas') {
      setDecisionForm((prev) => ({
        ...prev,
        status,
        naikKelas: '-',
        tinggalKelas: getCurrentClassText(selectedKelas)
      }));
    } else if (status === 'Lulus') {
      setDecisionForm((prev) => ({
        ...prev,
        status,
        naikKelas: 'Lulus dari Satuan Pendidikan SD Negeri Balil',
        tinggalKelas: '-'
      }));
    } else {
      setDecisionForm((prev) => ({
        ...prev,
        status,
        naikKelas: '........................................................',
        tinggalKelas: '........................................................'
      }));
    }
  };

  const handleSaveDecision = (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentStudent) return;

    if (decisionForm.applyToAll) {
      // Terapkan ke semua siswa di kelas ini
      const decisions: StudentDecision[] = siswaKelas.map((s) => ({
        nisn: s.nisn,
        namaSiswa: s.namaSiswa,
        kelas: selectedKelas,
        semester: selectedSemester,
        tahunPelajaran: db.sekolah.tahunPelajaran,
        statusKeputusan: decisionForm.status,
        naikKelas: decisionForm.naikKelas,
        tinggalKelas: decisionForm.tinggalKelas
      }));
      bulkSaveKeputusanSiswa(decisions);
    } else {
      // Simpan hanya untuk siswa saat ini
      const decision: StudentDecision = {
        nisn: currentStudent.nisn,
        namaSiswa: currentStudent.namaSiswa,
        kelas: selectedKelas,
        semester: selectedSemester,
        tahunPelajaran: db.sekolah.tahunPelajaran,
        statusKeputusan: decisionForm.status,
        naikKelas: decisionForm.naikKelas,
        tinggalKelas: decisionForm.tinggalKelas
      };
      saveKeputusanSiswa(decision);
    }

    setOpenDecisionModal(false);
  };

  // Handler Cetak Otomatis Semua Halaman (Buku Rapor Lengkap Hal 1 s.d. 4)
  const handlePrintAllPages = () => {
    setActivePage('semua');
    showSaveNotification('Mempersiapkan pencetakan otomatis semua halaman rapor (Kaver, Sekolah, Biodata Siswa, Nilai & TTD)...');
    setTimeout(() => {
      window.print();
    }, 150);
  };

  // Handler Cetak Halaman Aktif Saja
  const handlePrintCurrentPage = () => {
    window.print();
  };

  // Default print: Otomatis semua halaman
  const handlePrint = () => {
    handlePrintAllPages();
  };

  const reportPageIds = ['hal-1-kaver', 'hal-2-sekolah', 'hal-3-siswa', 'hal-4-rapor', 'hal-5-nilai', 'hal-5-keputusan'];

  // Navigasi scroll halus ke halaman spesifik rapor
  const scrollToReportPage = (elementId: string) => {
    if (activePage !== 'semua') {
      setActivePage('semua');
      setTimeout(() => {
        const el = document.getElementById(elementId);
        el?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }, 100);
    } else {
      const el = document.getElementById(elementId);
      el?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  const scrollToNextReportPage = () => {
    if (activePage !== 'semua') setActivePage('semua');
    const scrollContainer = document.querySelector('main');
    if (!scrollContainer) return;
    const currentScroll = scrollContainer.scrollTop;
    for (const id of reportPageIds) {
      const el = document.getElementById(id);
      if (el && el.offsetTop > currentScroll + 80) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
      }
    }
    scrollContainer.scrollTo({ top: scrollContainer.scrollHeight, behavior: 'smooth' });
  };

  const scrollToPrevReportPage = () => {
    if (activePage !== 'semua') setActivePage('semua');
    const scrollContainer = document.querySelector('main');
    if (!scrollContainer) return;
    const currentScroll = scrollContainer.scrollTop;
    for (let i = reportPageIds.length - 1; i >= 0; i--) {
      const id = reportPageIds[i];
      const el = document.getElementById(id);
      if (el && el.offsetTop < currentScroll - 80) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        return;
      }
    }
    scrollContainer.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleDownloadReportTemplate = () => {
    const schoolName = db.sekolah.namaSekolah || 'SD NEGERI BALIL';
    const schoolYear = db.sekolah.tahunPelajaran || '2026/2027';
    const pemda = db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG';
    const instansi = db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN';

    const wb = XLSX.utils.book_new();
    const columns = [
      'NISN',
      'Nama Siswa',
      'Kelas',
      'Semester',
      'Mata Pelajaran',
      'KKM',
      'Formatif',
      'Nilai PTS',
      'Nilai Akhir',
      'Capaian Kompetensi',
      'Deskripsi Capaian Pembelajaran',
      'Kehadiran (S/I/A)',
      'Catatan Wali Kelas',
      'Keputusan Kenaikan Kelas'
    ];

    const aoaData: (string | number)[][] = [
      [pemda],
      [instansi],
      [schoolName],
      [`FORMAT RESMI RAPOR AKHIR SEMESTER: ${activeSemester === 'Ganjil' || activeSemester === '1' ? 'SEMESTER 1 (GANJIL)' : 'SEMESTER 2 (GENAP)'}`],
      [`Nama Siswa: ${currentStudent?.namaSiswa || '-'} • NISN: ${currentStudent?.nisn || '-'} • Kelas: ${selectedKelas}`],
      [],
      ['NO', ...columns],
      ['(1)', ...columns.map((_, i) => `(${i + 2})`)],
      ...studentScores.map((sc, i) => [
        i + 1,
        currentStudent?.nisn || '-',
        currentStudent?.namaSiswa || '-',
        selectedKelas,
        activeSemester,
        sc.mapel.mataPelajaran,
        sc.score.kkm,
        sc.score.formatif ?? '',
        sc.score.nilaiPts ?? '',
        sc.score.nilaiAkhir > 0 ? sc.score.nilaiAkhir : '',
        sc.score.capaianKompetensi,
        sc.score.deskripsiRapor,
        `${studentAttendance.sakit}S / ${studentAttendance.izin}I / ${studentAttendance.alpa}A`,
        studentNote,
        displayNaikKelas
      ]),
      ['', '', '', '', '', 'JUMLAH NILAI', '', '', '', totalNilai, '', '', '', '', ''],
      ['', '', '', '', '', 'RATA-RATA NILAI', '', '', '', Number(rataRataNilai), '', '', '', '', '']
    ];

    for (let i = 0; i < 5; i++) {
      aoaData.push([studentScores.length + 1 + i, ...columns.map(() => '')]);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 18 },
      { wch: 30 },
      { wch: 8 },
      { wch: 10 },
      { wch: 32 },
      { wch: 8 },
      { wch: 10 },
      { wch: 10 },
      { wch: 12 },
      { wch: 22 },
      { wch: 60 },
      { wch: 16 },
      { wch: 45 },
      { wch: 25 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Template Rapor');
    XLSX.writeFile(wb, `template-rapor-semester-${activeSemester}-kelas-${selectedKelas}.xlsx`);
    showSaveNotification('Template resmi Rapor Akhir Semester (.xlsx) berhasil diunduh!');
  };

  const handlePrevStudent = () => {
    if (!currentStudent) return;
    const currentIdx = siswaKelas.findIndex((s) => s.idSiswa === currentStudent.idSiswa);
    if (currentIdx > 0) {
      setPrintTargetStudentId(siswaKelas[currentIdx - 1].idSiswa);
    }
  };

  const handleNextStudent = () => {
    if (!currentStudent) return;
    const currentIdx = siswaKelas.findIndex((s) => s.idSiswa === currentStudent.idSiswa);
    if (currentIdx < siswaKelas.length - 1) {
      setPrintTargetStudentId(siswaKelas[currentIdx + 1].idSiswa);
    }
  };

  if (!currentStudent) {
    return (
      <div className="bg-white rounded-2xl p-12 text-center text-slate-500 border border-slate-200">
        Belum ada data siswa di kelas {selectedKelas}. Silakan tambahkan data siswa terlebih dahulu.
      </div>
    );
  }

  // Get matching academic scores dynamically linked to Tujuan Pembelajaran (TP)
  const studentScores = filteredMataPelajaran.map((m) => {
    const sc = db.nilaiAkademik.find(
      (n) =>
        matchesStudentScore(n, currentStudent) &&
        n.mataPelajaran === m.mataPelajaran &&
        n.kelas === selectedKelas &&
        n.semester === activeSemester
    );

    const kkm = sc?.kkm || m.kkm || 75;
    const formatif = sc?.formatif;
    const nilaiPts = sc?.pts;
    const nilaiAkhir = sc ? sc.nilaiAkhir : 0;
    const tpResult =
      nilaiAkhir > 0
        ? generateReportDescriptionFromTP(
            m.mataPelajaran,
            selectedKelas,
            nilaiAkhir,
            db.tujuanPembelajaran,
            currentStudent.namaSiswa,
            currentStudent.agama
          )
        : null;

    // If score has custom non-empty description, respect it; otherwise use the connected TP description
    const deskripsiRapor =
      sc?.deskripsiRapor && sc.deskripsiRapor.trim() !== ''
        ? sc.deskripsiRapor
        : tpResult?.deskripsiRapor || '-';

    const capaianKompetensi =
      sc?.capaianKompetensi ||
      (nilaiAkhir > 0
        ? nilaiAkhir >= 88
          ? 'Sangat Baik'
          : nilaiAkhir >= 78
          ? 'Baik'
          : nilaiAkhir >= 68
          ? 'Cukup'
          : 'Perlu Pendampingan'
        : '-');

    return {
      mapel: m,
      score: {
        kkm,
        formatif,
        nilaiPts,
        nilaiAkhir,
        capaianKompetensi,
        deskripsiRapor,
        deskripsiTercapai: sc?.deskripsiTercapai || tpResult?.deskripsiTercapai || '',
        deskripsiPerluPeningkatan: sc?.deskripsiPerluPeningkatan || tpResult?.deskripsiPerluPeningkatan || ''
      },
      linkedTpKodes: tpResult?.linkedTpKodes || []
    };
  });

  // Hitung Jumlah Nilai & Rata-Rata Nilai
  const totalNilai = studentScores.reduce((acc, curr) => acc + curr.score.nilaiAkhir, 0);
  const scoredCount = studentScores.filter((curr) => curr.score.nilaiAkhir > 0).length;
  const rataRataNilai = scoredCount > 0 ? (totalNilai / scoredCount).toFixed(2) : '0';

  // Ekstrakurikuler siswa
  const studentEkskul = db.ekstrakurikuler.filter(
    (e) => matchesStudentRecord(e, currentStudent) && e.kelas === selectedKelas && e.semester === activeSemester
  );

  // Kehadiran siswa
  const studentAttendance = db.kehadiran.find(
    (k) => matchesStudentRecord(k, currentStudent) && k.kelas === selectedKelas && k.semester === activeSemester
  ) || {
    sakit: 0,
    izin: 0,
    alpa: 0
  };

  // Catatan wali kelas (berbeda per semester)
  const studentNote =
    db.catatanWali.find(
      (c) => matchesStudentRecord(c, currentStudent) && c.kelas === selectedKelas && c.semester === activeSemester
    )?.catatanWaliKelas || '';

  // Wali Kelas info (terhubung langsung dari Profil Sekolah & Data Kelas)
  const waliKelasData = db.kelas.find((k) => k.namaKelas === selectedKelas);
  const waliKelasGuru = {
    namaGuru:
      db.sekolah.waliKelas ||
      waliKelasData?.waliKelas ||
      'Selpius Nabyal, S.Pd.Gr',
    nip: db.sekolah.nipWaliKelas || '-'
  };

  // Nilai teks keputusan untuk ditampilkan
  const displayNaikKelas =
    currentDecision?.statusKeputusan === 'Naik Kelas'
      ? currentDecision.naikKelas || getNextClass(selectedKelas)
      : currentDecision?.statusKeputusan === 'Tinggal di Kelas'
      ? '-'
      : currentDecision?.naikKelas || '........................................................';

  const displayTinggalKelas =
    currentDecision?.statusKeputusan === 'Tinggal di Kelas'
      ? currentDecision.tinggalKelas || getCurrentClassText(selectedKelas)
      : currentDecision?.statusKeputusan === 'Naik Kelas'
      ? '-'
      : currentDecision?.tinggalKelas || '........................................................';

  const { hasPart2 } = splitAcademicScoresForPages(studentScores);

  return (
    <div className="space-y-6">
      {/* Control Bar (hidden on print) */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-xs space-y-4 no-print">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className={`p-2.5 rounded-xl ${
              activeSemester === 'Ganjil' || activeSemester === '1'
                ? 'bg-blue-50 text-blue-700'
                : 'bg-emerald-50 text-emerald-700'
            }`}>
              <FileText size={24} />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-800">
                {!forcedSemester
                  ? `Rapor Gabungan Semester 1 & 2 • Lembar Aktif: ${activeSemester === 'Ganjil' || activeSemester === '1' ? 'Semester 1 (Ganjil)' : 'Semester 2 (Genap)'}`
                  : activeSemester === 'Ganjil' || activeSemester === '1'
                  ? 'Pratinjau Rapor Semester 1 (Ganjil) - Sendiri'
                  : 'Pratinjau Rapor Semester 2 (Genap) - Sendiri'}
              </h2>
              <p className="text-xs text-slate-500">
                Format Resmi Kurikulum Merdeka • Kelas {selectedKelas} • {
                  !forcedSemester
                    ? 'Rapor Gabungan Semester Satu dan Semester Dua (Gunakan tombol tab di bawah untuk beralih lembar)'
                    : activeSemester === 'Ganjil' || activeSemester === '1'
                    ? 'Semester 1 (Ganjil - Sendiri)'
                    : 'Semester 2 (Genap + Keputusan Kenaikan Kelas)'
                }
              </p>
            </div>
          </div>

          {/* Student Selector & Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <StudentFilterSelect compact />
            <button
              onClick={handlePrevStudent}
              className="p-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl transition-colors cursor-pointer"
              title="Siswa Sebelumnya"
            >
              <ChevronLeft size={16} />
            </button>

            <select
              value={currentStudent.idSiswa}
              onChange={(e) => setPrintTargetStudentId(e.target.value)}
              className="px-3 py-2 text-xs font-bold bg-slate-50 border border-slate-300 rounded-xl text-slate-800 focus:outline-hidden"
            >
              {siswaKelas.map((s) => (
                <option key={s.idSiswa} value={s.idSiswa}>
                  {s.namaSiswa} ({s.nisn})
                </option>
              ))}
            </select>

            <button
              onClick={handleNextStudent}
              className="p-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl transition-colors cursor-pointer"
              title="Siswa Berikutnya"
            >
              <ChevronRight size={16} />
            </button>

            {/* Tombol Sinkronkan / Hubungkan TP ke Rapor */}
            <button
              type="button"
              onClick={() => {
                const count = syncAllScoresWithTPDescriptions(activeSemester);
                showSaveNotification(`Berhasil memperbarui ${count} data deskripsi dari TP ke rapor!`);
              }}
              className="w-8 h-8 inline-flex items-center justify-center bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg shadow-xs transition-colors cursor-pointer"
              title="Hubungkan TP ke Rapor"
              aria-label="Hubungkan TP ke Rapor"
            >
              <Zap size={15} />
            </button>

            {/* Tombol Input Keputusan (Hanya Relevan untuk Semester 2) */}
            {(activeSemester === 'Genap' || activeSemester === '2') && (
              <button
                type="button"
                onClick={handleOpenDecisionModal}
                className="w-8 h-8 inline-flex items-center justify-center bg-amber-500 hover:bg-amber-600 text-slate-950 rounded-lg shadow-xs transition-colors cursor-pointer"
                title="Input Keputusan Naik Kelas / Tinggal di Kelas"
                aria-label="Input Keputusan"
              >
                <Edit3 size={15} />
              </button>
            )}

            {/* Tombol Preview Rapor Layar Penuh (Pop-up Interaktif) */}
            <button
              type="button"
              onClick={() => openReportPreview(currentStudent.idSiswa)}
              className="w-8 h-8 inline-flex items-center justify-center bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white rounded-lg shadow-xs transition-all cursor-pointer ring-1 ring-purple-300"
              title="Preview Rapor Layar Penuh"
              aria-label="Preview Rapor"
            >
              <Eye size={15} className="text-amber-300" />
            </button>

            {/* Tombol Cetak Otomatis Semua Halaman (Buku Rapor Lengkap 1 s.d. 4) */}
            <button
              type="button"
              onClick={handlePrintAllPages}
              className="w-8 h-8 inline-flex items-center justify-center bg-gradient-to-r from-blue-700 to-indigo-800 hover:from-blue-800 hover:to-indigo-900 text-white rounded-lg shadow-xs transition-all cursor-pointer ring-1 ring-blue-300"
              title="Cetak Semua Halaman (Otomatis)"
              aria-label="Cetak Semua Halaman"
            >
              <Printer size={15} className="text-amber-300" />
            </button>

            {/* Tombol Cetak Lembar Aktif */}
            <button
              type="button"
              onClick={handlePrintCurrentPage}
              className="w-8 h-8 inline-flex items-center justify-center bg-white hover:bg-slate-100 text-slate-700 rounded-lg shadow-xs transition-colors cursor-pointer border border-slate-300"
              title="Cetak Lembar Ini"
              aria-label="Cetak Lembar Ini"
            >
              <FileCheck2 size={15} />
            </button>
          </div>
        </div>

        {/* TAB PILIHAN HALAMAN BUKU RAPOR: KAVER, RAPOR PESERTA DIDIK, IDENTITAS, DAN RAPOR */}
        <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
            <span className="text-xs font-bold text-slate-700 mr-1">Halaman Buku Rapor:</span>
            <button
              type="button"
              onClick={() => setActivePage('kaver')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activePage === 'kaver'
                  ? 'bg-amber-600 text-white border-amber-600 shadow-md ring-2 ring-amber-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>📘 Halaman Kaver</span>
            </button>
            <button
              type="button"
              onClick={() => setActivePage('rapor_peserta_didik')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activePage === 'rapor_peserta_didik'
                  ? 'bg-cyan-700 text-white border-cyan-700 shadow-md ring-2 ring-cyan-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>🏫 Rapor Peserta Didik</span>
            </button>
            <button
              type="button"
              onClick={() => setActivePage('identitas')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activePage === 'identitas'
                  ? 'bg-purple-700 text-white border-purple-700 shadow-md ring-2 ring-purple-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>👤 Identitas Peserta Didik</span>
            </button>
            <button
              type="button"
              onClick={() => setActivePage('rapor')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activePage === 'rapor'
                  ? 'bg-blue-700 text-white border-blue-700 shadow-md ring-2 ring-blue-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>📄 Rapor (Lembar Nilai)</span>
            </button>
            <button
              type="button"
              onClick={() => setActivePage('semua')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activePage === 'semua'
                  ? 'bg-slate-900 text-white border-slate-900 shadow-md ring-2 ring-slate-400'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>📑 Lengkap (Semua Halaman)</span>
            </button>
          </div>

          <div className="text-[11px] font-bold text-slate-700 bg-slate-100 px-3 py-1 rounded-lg border border-slate-200">
            Halaman Aktif:{' '}
            <strong className="text-blue-900 uppercase">
              {activePage === 'kaver'
                ? 'Halaman Kaver (Sampul Depan)'
                : activePage === 'rapor_peserta_didik'
                ? 'Rapor Peserta Didik (Identitas Sekolah)'
                : activePage === 'identitas'
                ? 'Identitas Peserta Didik (Biodata Siswa)'
                : activePage === 'rapor'
                ? 'Rapor (Lembar Nilai & Capaian)'
                : 'Semua Halaman (Buku Rapor Lengkap)'}
            </strong>
          </div>
        </div>

        {/* TAB PEMISAH RAPOR SEMESTER 1 SENDIRI vs SEMESTER 2 SENDIRI */}
        <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-xs font-bold text-slate-700 mr-1">Pemisahan Rapor:</span>
            <button
              type="button"
              onClick={() => setSelectedSemester('Ganjil')}
              className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activeSemester === 'Ganjil' || activeSemester === '1'
                  ? 'bg-blue-600 text-white border-blue-600 shadow-md ring-2 ring-blue-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>📄 RAPOR SEMESTER 1 (GANJIL)</span>
              <span className={`text-[9px] px-1.5 py-0.5 rounded font-bold ${
                activeSemester === 'Ganjil' || activeSemester === '1'
                  ? 'bg-blue-800 text-white'
                  : 'bg-slate-200 text-slate-700'
              }`}>
                SENDIRI
              </span>
            </button>

            <button
              type="button"
              onClick={() => setSelectedSemester('Genap')}
              className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs font-black transition-all cursor-pointer border ${
                activeSemester === 'Genap' || activeSemester === '2'
                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-md ring-2 ring-emerald-300'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
            >
              <span>🎓 RAPOR SEMESTER 2 (GENAP)</span>
              <span className={`text-[9px] px-1.5 py-0.5 rounded font-bold ${
                activeSemester === 'Genap' || activeSemester === '2'
                  ? 'bg-emerald-800 text-white'
                  : 'bg-slate-200 text-slate-700'
              }`}>
                + KENAIKAN KELAS
              </span>
            </button>
          </div>

          <div className="text-[11px] text-slate-500 font-medium">
            {activeSemester === 'Ganjil' || activeSemester === '1' ? (
              <span className="text-blue-700 bg-blue-50 px-2.5 py-1 rounded-lg border border-blue-200 font-bold">
                ✓ Rapor Semester 1 (Ganjil)
              </span>
            ) : (
              <span className="text-emerald-700 bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200 font-bold">
                ✓ Rapor Semester 2 (Genap)
              </span>
            )}
          </div>
        </div>

        {/* ROW FILTER MATA PELAJARAN PILIHAN & WAJIB */}
        <div className="pt-3 border-t border-slate-100 space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5 mr-1">
                <Filter size={14} className="text-indigo-600" />
                Filter Mata Pelajaran di Rapor:
              </span>
              <button
                type="button"
                onClick={() => setMapelCategoryFilter('semua')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border cursor-pointer ${
                  mapelCategoryFilter === 'semua'
                    ? 'bg-indigo-600 text-white border-indigo-600 shadow-xs ring-2 ring-indigo-300'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Semua Mapel ({db.mataPelajaran.length})
              </button>
              <button
                type="button"
                onClick={() => setMapelCategoryFilter('wajib')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border cursor-pointer ${
                  mapelCategoryFilter === 'wajib'
                    ? 'bg-indigo-600 text-white border-indigo-600 shadow-xs ring-2 ring-indigo-300'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Hanya Mapel Wajib ({db.mataPelajaran.filter((m) => m.kelompok === 'Wajib').length})
              </button>
              <button
                type="button"
                onClick={() => setMapelCategoryFilter('pilihan')}
                className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border cursor-pointer ${
                  mapelCategoryFilter === 'pilihan'
                    ? 'bg-amber-600 text-white border-amber-600 shadow-xs ring-2 ring-amber-300'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                Khusus Mapel Pilihan ({mapelPilihanList.length})
              </button>
            </div>

            <div className="text-[11px] text-slate-600 font-semibold bg-slate-100 px-3 py-1 rounded-lg border border-slate-200">
              Menampilkan <strong className="text-indigo-900">{filteredMataPelajaran.length}</strong> dari {db.mataPelajaran.length} mata pelajaran
            </div>
          </div>

          {/* CHECKLIST MATA PELAJARAN PILIHAN YANG AKTIF / DIAMBIL */}
          {mapelPilihanList.length > 0 && mapelCategoryFilter !== 'wajib' && (
            <div className="p-2.5 bg-amber-50/70 rounded-xl border border-amber-200 text-xs flex flex-wrap items-center justify-between gap-2.5">
              <div className="flex flex-wrap items-center gap-2">
                <span className="font-bold text-amber-950 flex items-center gap-1.5 text-[11px]">
                  <BookOpen size={13} className="text-amber-700" />
                  Mata Pelajaran Pilihan yang Diambil Siswa (Centang untuk Tampilkan di Rapor):
                </span>
                {mapelPilihanList.map((m) => {
                  const isChecked = activePilihanMapels.includes(m.mataPelajaran);
                  return (
                    <button
                      key={m.mataPelajaran}
                      type="button"
                      onClick={() => togglePilihanMapel(m.mataPelajaran)}
                      className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-bold transition-all border cursor-pointer ${
                        isChecked
                          ? 'bg-white text-amber-950 border-amber-400 shadow-xs ring-1 ring-amber-300'
                          : 'bg-white/60 text-slate-400 border-slate-300 line-through'
                      }`}
                      title={isChecked ? 'Klik untuk sembunyikan dari rapor' : 'Klik untuk tampilkan di rapor'}
                    >
                      {isChecked ? (
                        <CheckSquare size={13} className="text-amber-600" />
                      ) : (
                        <Square size={13} className="text-slate-400" />
                      )}
                      <span>{m.mataPelajaran}</span>
                      <span className="text-[9px] px-1 rounded bg-amber-100 text-amber-800 font-semibold">
                        {m.kelompok}
                      </span>
                    </button>
                  );
                })}
              </div>

              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={selectAllPilihan}
                  title="Pilih Semua Mata Pelajaran Pilihan"
                  aria-label="Pilih Semua"
                  className="w-7 h-7 inline-flex items-center justify-center bg-white hover:bg-amber-100 text-amber-800 rounded-lg border border-amber-300 cursor-pointer"
                >
                  <CheckSquare size={13} />
                </button>
                <button
                  type="button"
                  onClick={clearAllPilihan}
                  title="Hapus Semua Pilihan"
                  aria-label="Hapus Pilihan"
                  className="w-7 h-7 inline-flex items-center justify-center bg-white hover:bg-slate-100 text-slate-600 rounded-lg border border-slate-300 cursor-pointer"
                >
                  <Square size={13} />
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* BARIS TOMBOL AKSI LENGKAP RAPOR AKHIR SEMESTER DI SAMPING / DI ATAS INPUT */}
      <StandardActionToolbar
        title="Aksi Rapor Akhir Semester"
        badge={currentStudent?.namaSiswa || 'Siswa Terpilih'}
        onSave={() => showSaveNotification('Lembar rapor siswa berhasil disimpan!')}
        saveLabel="Simpan Rapor"
        onUpdate={() => showSaveNotification('Tampilan Rapor Akhir Semester diperbarui!')}
        updateLabel="Update Rapor"
        onEdit={handleOpenDecisionModal}
        editLabel="Input Keputusan"
        onCancel={() => {}}
        cancelLabel="Batal"
        onDelete={() => {}}
        deleteLabel="Bersihkan Filter"
        onImport={handleImportFile}
        importLabel="Impor Data"
        onUpload={handleImportFile}
        uploadLabel="Upload File"
        onExport={() => {
          const wb = XLSX.utils.book_new();
          const data: (string | number)[][] = [
            [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
            [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
            [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
            ['RAPOR (LAPORAN HASIL BELAJAR PESERTA DIDIK)'],
            [`Nama Siswa: ${currentStudent?.namaSiswa || '-'} • NISN: ${currentStudent?.nisn || '-'}`],
            [`Kelas: ${selectedKelas} • Semester: ${activeSemester}`],
            [],
            ['NO', 'MATA PELAJARAN', 'KKM', 'FORMATIF', 'NILAI PTS', 'NILAI AKHIR', 'CAPAIAN KOMPETENSI', 'DESKRIPSI CAPAIAN PEMBELAJARAN'],
            ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)', '(8)'],
            ...studentScores.map((sc, i) => [
              i + 1,
              sc.mapel.mataPelajaran,
              sc.score.kkm,
              sc.score.formatif ?? '',
              sc.score.nilaiPts ?? '',
              sc.score.nilaiAkhir > 0 ? sc.score.nilaiAkhir : '',
              sc.score.capaianKompetensi,
              sc.score.deskripsiRapor
            ]),
            ['', 'JUMLAH NILAI', '', '', '', totalNilai, '', ''],
            ['', 'RATA-RATA NILAI', '', '', '', Number(rataRataNilai), '', '']
          ];
          const ws = XLSX.utils.aoa_to_sheet(data);
          ws['!cols'] = [{ wch: 6 }, { wch: 32 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 22 }, { wch: 60 }];
          XLSX.utils.book_append_sheet(wb, ws, 'Rapor Siswa');
          XLSX.writeFile(wb, `rapor-${currentStudent?.namaSiswa || 'siswa'}-kelas-${selectedKelas}.xlsx`);
        }}
        exportLabel="Expor Excel (.xlsx)"
        onClear={() => {}}
        clearLabel="Bersihkan"
        onPrint={handlePrint}
        onTemplate={handleDownloadReportTemplate}
        templateLabel="Template Rapor"
      />

      <style>{`
        @media print {
          @page {
            size: ${finalPageCssSize};
            margin: ${paperMargins.top}mm ${paperMargins.right}mm ${paperMargins.bottom}mm ${paperMargins.left}mm;
          }
          html, body, #root, #root > div, main {
            overflow: visible !important;
            height: auto !important;
            max-height: none !important;
            position: static !important;
            width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
            background: #ffffff !important;
          }
          body {
            background-color: #fff !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .no-print, header, nav, aside, button:not(.allow-print) {
            display: none !important;
          }
          .print-page-break {
            page-break-after: always !important;
            break-after: page !important;
            clear: both !important;
            display: block !important;
            margin-bottom: 0 !important;
            padding: 0 !important;
            border: none !important;
            box-shadow: none !important;
            background: #ffffff !important;
          }
          .print-avoid-break, .report-cover-page, table, tr, td, th {
            page-break-inside: avoid !important;
            break-inside: avoid !important;
          }
          .printable-area {
            box-shadow: none !important;
            border: none !important;
            padding: 0 !important;
            margin: 0 !important;
            width: 100% !important;
            background: #ffffff !important;
          }
        }
      `}</style>

      {/* PENGATURAN BENTUK HURUF & UKURAN HURUF CETAK RAPOR */}
      <PrintTypographyControl className="mb-4" />

      {/* FLOATING STICKY QUICK SCROLL BAR UNTUK SEMUA HALAMAN RAPOR */}
      {scrollBarMode === 'hidden' ? (
        <div className="no-print flex justify-end mb-3">
          <button
            type="button"
            onClick={() => setScrollBarMode('full')}
            className="w-8 h-8 bg-slate-900 hover:bg-slate-800 text-white rounded-lg shadow-md border border-slate-700 cursor-pointer inline-flex items-center justify-center transition-all ring-1 ring-blue-400"
            title="Tampilkan Bilah Navigasi Halaman & Cetak"
            aria-label="Tampilkan Navigasi Halaman"
          >
            <Eye size={14} className="text-amber-400" />
          </button>
        </div>
      ) : (
        <div className="no-print sticky top-2 z-20 bg-white/95 backdrop-blur-md py-1.5 px-3 border border-slate-300 rounded-xl shadow-md mb-4 flex flex-wrap items-center justify-between gap-2 transition-all">
          <div className="flex items-center gap-2">
            <span className="font-extrabold text-blue-900 text-xs">
              Hal:
            </span>
          </div>

          <div className="ml-auto flex items-center justify-end flex-wrap gap-1.5">
            {scrollBarMode === 'full' && (
              <div className="flex flex-wrap items-center gap-1">
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-1-kaver')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-amber-50 hover:bg-amber-100 text-amber-950 border border-amber-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 1: Kaver"
                  aria-label="Halaman 1"
                >
                  1
                </button>
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-2-sekolah')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-cyan-50 hover:bg-cyan-100 text-cyan-950 border border-cyan-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 2: Profil Sekolah"
                  aria-label="Halaman 2"
                >
                  2
                </button>
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-3-siswa')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-purple-50 hover:bg-purple-100 text-purple-950 border border-purple-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 3: Identitas Siswa"
                  aria-label="Halaman 3"
                >
                  3
                </button>
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-4-rapor')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-blue-50 hover:bg-blue-100 text-blue-950 border border-blue-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 4: Lembar Nilai & Capaian"
                  aria-label="Halaman 4"
                >
                  4
                </button>
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-5-nilai')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-sky-50 hover:bg-sky-100 text-sky-950 border border-sky-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 5: Lanjutan Nilai & Muatan Lokal"
                  aria-label="Halaman 5"
                >
                  5
                </button>
                <button
                  type="button"
                  onClick={() => scrollToReportPage('hal-5-keputusan')}
                  className="w-7 h-7 inline-flex items-center justify-center bg-emerald-50 hover:bg-emerald-100 text-emerald-950 border border-emerald-300 rounded-lg text-[11px] font-black cursor-pointer transition-colors"
                  title="Lompat ke Hal 6: Ekstrakurikuler, Ketidakhadiran, Keputusan & TTD"
                  aria-label="Halaman 6"
                >
                  6
                </button>
              </div>
            )}

            {/* Quick Scroll Next / Prev */}
            <div className="inline-flex items-center bg-slate-100 rounded-lg p-0.5 border border-slate-200 gap-0.5">
              <button
                type="button"
                onClick={scrollToPrevReportPage}
                className="w-7 h-7 inline-flex items-center justify-center hover:bg-white text-slate-700 rounded-md transition-all cursor-pointer"
                title="Scroll ke Halaman Sebelumnya"
                aria-label="Halaman Sebelumnya"
              >
                <ArrowUp size={13} />
              </button>
              <button
                type="button"
                onClick={scrollToNextReportPage}
                className="w-7 h-7 inline-flex items-center justify-center hover:bg-white text-slate-700 rounded-md transition-all cursor-pointer"
                title="Scroll ke Halaman Berikutnya"
                aria-label="Halaman Berikutnya"
              >
                <ArrowDown size={13} />
              </button>
            </div>

            {/* Kontrol Cepat Besar/Kecil Tulisan & Logo */}
            <div className="inline-flex items-center bg-blue-50 border border-blue-200 rounded-lg p-0.5 gap-0.5">
              <button
                type="button"
                onClick={() => adjustFontSize('down')}
                className="w-6 h-6 inline-flex items-center justify-center bg-white hover:bg-slate-100 text-slate-800 rounded text-[10px] font-black border border-slate-200 cursor-pointer"
                title="Perkecil tulisan dalam rapor"
              >
                A-
              </button>
              <span className="px-1 text-[10px] font-black text-blue-900">{printFontSize}</span>
              <button
                type="button"
                onClick={() => adjustFontSize('up')}
                className="w-6 h-6 inline-flex items-center justify-center bg-blue-600 hover:bg-blue-700 text-white rounded text-[10px] font-black cursor-pointer"
                title="Perbesar tulisan dalam rapor"
              >
                A+
              </button>
            </div>

            <div className="inline-flex items-center bg-emerald-50 border border-emerald-200 rounded-lg p-0.5 gap-0.5">
              <button
                type="button"
                onClick={() => {
                  const next = Math.max(50, printLogoScale - 10);
                  setPrintLogoScale(next);
                  showSaveNotification(`Ukuran logo rapor diubah ke ${next}%`);
                }}
                className="px-1.5 h-6 inline-flex items-center justify-center bg-white hover:bg-slate-100 text-slate-800 rounded text-[10px] font-black border border-slate-200 cursor-pointer"
                title="Perkecil logo dalam rapor"
              >
                <ImageIcon size={11} className="mr-0.5" />-
              </button>
              <span className="px-1 text-[10px] font-black text-emerald-900">{printLogoScale}%</span>
              <button
                type="button"
                onClick={() => {
                  const next = Math.min(200, printLogoScale + 10);
                  setPrintLogoScale(next);
                  showSaveNotification(`Ukuran logo rapor diubah ke ${next}%`);
                }}
                className="px-1.5 h-6 inline-flex items-center justify-center bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-black cursor-pointer"
                title="Perbesar logo dalam rapor"
              >
                <ImageIcon size={11} className="mr-0.5" />+
              </button>
            </div>

            {/* Kontrol Cepat Besar/Kecil Kertas & Batas Kiri/Kanan/Atas/Bawah */}
            <div className="inline-flex items-center bg-amber-50 border border-amber-300 rounded-lg p-0.5 gap-0.5" title="Perbesar / Perkecil Skala Kertas">
              <button
                type="button"
                onClick={() => adjustPaperScale(-5)}
                className="px-1.5 h-6 inline-flex items-center justify-center bg-white hover:bg-slate-100 text-slate-800 rounded text-[10px] font-black border border-slate-200 cursor-pointer"
                title="Perkecil ukuran kertas (-5%)"
              >
                Kertas-
              </button>
              <span className="px-1 text-[10px] font-black text-amber-950">{printScale}%</span>
              <button
                type="button"
                onClick={() => adjustPaperScale(5)}
                className="px-1.5 h-6 inline-flex items-center justify-center bg-amber-600 hover:bg-amber-700 text-white rounded text-[10px] font-black cursor-pointer"
                title="Perbesar ukuran kertas (+5%)"
              >
                Kertas+
              </button>
            </div>

            <div className="inline-flex items-center bg-indigo-50 border border-indigo-200 rounded-lg p-0.5 gap-1">
              {(
                [
                  { side: 'left', short: 'Ki', title: 'Kiri' },
                  { side: 'right', short: 'Ka', title: 'Kanan' },
                  { side: 'top', short: 'At', title: 'Atas' },
                  { side: 'bottom', short: 'Bw', title: 'Bawah' }
                ] as const
              ).map(({ side, short, title }) => (
                <div key={side} className="inline-flex items-center gap-0.5 bg-white border border-indigo-200 rounded px-0.5">
                  <span className="text-[9px] font-black text-indigo-900 px-0.5" title={`Batas ${title}`}>
                    {short}
                  </span>
                  <button
                    type="button"
                    onClick={() => adjustPaperMargin(side, -2)}
                    className="w-4 h-5 inline-flex items-center justify-center bg-slate-100 hover:bg-slate-200 text-slate-800 rounded text-[10px] font-black cursor-pointer"
                    title={`Perkecil batas ${title} (-2mm)`}
                  >
                    -
                  </button>
                  <span className="text-[9px] font-mono font-bold text-slate-800 min-w-[16px] text-center">
                    {paperMargins[side]}
                  </span>
                  <button
                    type="button"
                    onClick={() => adjustPaperMargin(side, 2)}
                    className="w-4 h-5 inline-flex items-center justify-center bg-indigo-600 hover:bg-indigo-700 text-white rounded text-[10px] font-black cursor-pointer"
                    title={`Perbesar batas ${title} (+2mm)`}
                  >
                    +
                  </button>
                </div>
              ))}
            </div>

            {/* Tombol Cetak Semua Halaman */}
            <button
              type="button"
              onClick={handlePrintAllPages}
              className="w-7 h-7 inline-flex items-center justify-center bg-blue-700 hover:bg-blue-800 text-white rounded-lg cursor-pointer shadow-xs transition-all"
              title="Cetak Semua Halaman"
              aria-label="Cetak Semua Halaman"
            >
              <Printer size={14} className="text-amber-300" />
            </button>

            {/* Kontrol Bilah: Rampingkan / Perluas & Hilangkan */}
            <div className="inline-flex items-center rounded-lg bg-slate-100 p-0.5 border border-slate-300 shadow-2xs gap-0.5">
              {scrollBarMode === 'compact' ? (
                <button
                  type="button"
                  onClick={() => setScrollBarMode('full')}
                  className="w-7 h-7 inline-flex items-center justify-center rounded-md text-blue-700 hover:bg-blue-50 transition-colors cursor-pointer"
                  title="Perluas tombol scroll halaman"
                  aria-label="Perluas"
                >
                  <Maximize2 size={13} />
                </button>
              ) : (
                <button
                  type="button"
                  onClick={() => setScrollBarMode('compact')}
                  className="w-7 h-7 inline-flex items-center justify-center rounded-md text-slate-700 hover:bg-slate-200 transition-colors cursor-pointer"
                  title="Rampingkan bilah scroll halaman"
                  aria-label="Rampingkan"
                >
                  <Minimize2 size={13} />
                </button>
              )}
              <button
                type="button"
                onClick={() => setScrollBarMode('hidden')}
                className="w-7 h-7 inline-flex items-center justify-center rounded-md text-rose-700 hover:bg-rose-50 transition-colors cursor-pointer"
                title="Sembunyikan bilah scroll ini"
                aria-label="Hilangkan"
              >
                <EyeOff size={13} />
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="bg-[#9ca3af]/45 p-2.5 sm:p-5 rounded-2xl print:bg-white print:p-0 flex flex-col xl:flex-row items-start justify-center gap-4">
        <div
          className="w-full xl:flex-1 max-w-4xl mx-auto border-[5px] border-[#0022dd] bg-white shadow-2xl print:border-none print:shadow-none print:max-w-none scroll-smooth"
          id="academic-report-scroll-container"
        >
          {/* 1. HALAMAN KAVER RAPOR */}
          {(activePage === 'kaver' || activePage === 'semua') && (
            <div id="hal-1-kaver" className={`scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none ${activePage === 'semua' ? 'print-page-break' : ''}`}>
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 1
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto overflow-visible print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                <ReportCoverPage sekolah={db.sekolah} student={currentStudent} paperSize={paperSize} paperOrientation={paperOrientation} fontFamily={printFontFamily} fontSize={printFontSize} />
              </div>
            </div>
          )}

          {/* 2. HALAMAN RAPOR PESERTA DIDIK (IDENTITAS SEKOLAH) */}
          {(activePage === 'rapor_peserta_didik' || activePage === 'semua') && (
            <div id="hal-2-sekolah" className={`scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none ${activePage === 'semua' ? 'print-page-break' : ''}`}>
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 2
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto overflow-visible print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                <ReportSchoolIdentityPage sekolah={db.sekolah} paperSize={paperSize} paperOrientation={paperOrientation} fontFamily={printFontFamily} fontSize={printFontSize} />
              </div>
            </div>
          )}

          {/* 3. HALAMAN IDENTITAS PESERTA DIDIK (BIODATA LENGKAP SISWA) */}
          {(activePage === 'identitas' || activePage === 'semua') && (
            <div id="hal-3-siswa" className={`scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none ${activePage === 'semua' ? 'print-page-break' : ''}`}>
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 3
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto overflow-visible print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                <ReportStudentIdentityPage sekolah={db.sekolah} student={currentStudent} paperSize={paperSize} paperOrientation={paperOrientation} fontFamily={printFontFamily} fontSize={printFontSize} />
              </div>
            </div>
          )}

          {/* 4. HALAMAN RAPOR (LEMBAR NILAI BAGIAN 1) */}
          {(activePage === 'rapor' || activePage === 'semua') && (
            <div id="hal-4-rapor" className={`scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none ${activePage === 'semua' || activePage === 'rapor' ? 'print-page-break' : ''}`}>
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 4
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto printable-area text-slate-900 leading-normal print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                <ReportAcademicHeaderAndScores
                  student={currentStudent}
                  sekolah={db.sekolah}
                  selectedKelas={selectedKelas}
                  selectedFase={selectedFase}
                  semester={activeSemester}
                  studentScores={studentScores}
                  pagePart={hasPart2 ? 'part1' : 'all'}
                />
              </div>
            </div>
          )}

          {/* 5. LANJUTAN TABEL NILAI (BAGIAN 2 + MUATAN LOKAL) */}
          {(activePage === 'rapor' || activePage === 'semua') && hasPart2 && (
            <div id="hal-5-nilai" className="scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none print-page-break">
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 5
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto printable-area text-slate-900 leading-normal print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                <ReportAcademicHeaderAndScores
                  student={currentStudent}
                  sekolah={db.sekolah}
                  selectedKelas={selectedKelas}
                  selectedFase={selectedFase}
                  semester={activeSemester}
                  studentScores={studentScores}
                  pagePart="part2"
                />
              </div>
            </div>
          )}

          {/* 6. HALAMAN EKSTRAKURIKULER, KETIDAKHADIRAN, KEPUTUSAN & TTD */}
          {(activePage === 'rapor' || activePage === 'semua') && (
            <div id="hal-5-keputusan" className="scroll-mt-20 border-b-4 border-dashed border-[#0022dd] print:border-none print-page-break">
              <div className="no-print text-center text-[11px] font-serif text-slate-400 pt-1 select-none">
                Page 6
              </div>
              <div
                className={`report-typography-scope report-paper-sheet bg-white ${sheetMaxWClass} mx-auto printable-area text-slate-900 leading-normal print:border-none print:shadow-none print:p-0 print:m-0 print:bg-white print:max-w-none print:w-full`}
                style={{ fontFamily: reportFontCss, fontSize: printFontSize, lineHeight: printLineHeight }}
              >
                {/* EKSTRAKURIKULER (5 BARIS) */}
                <ReportExtracurricularTable studentEkskul={studentEkskul} minRows={5} />

                {/* KETIDAKHADIRAN, KEPUTUSAN (KHUSUS SEM 2), CATATAN WALI & TTD */}
                <ReportBottomSection
                  semester={activeSemester}
                  student={currentStudent}
                  selectedKelas={selectedKelas}
                  sekolah={db.sekolah}
                  waliKelasGuru={waliKelasGuru}
                  catatanWaliKelas={studentNote}
                  kehadiran={studentAttendance}
                  keputusan={currentDecision}
                  pengesahanOleh={pengesahanOleh}
                  onEditKeputusan={activeSemester === 'Genap' || activeSemester === '2' ? handleOpenDecisionModal : undefined}
                />
              </div>
            </div>
          )}
        </div>

        {/* PANEL KONTROL SAMPING KANAN */}
        <aside className="w-full xl:w-[230px] shrink-0 space-y-4 no-print print:hidden sticky top-16">
          <div className="bg-white border border-slate-400 p-2.5 text-[11px] text-black shadow-md space-y-1.5">
            <div className="font-bold text-slate-900">Petunjuk :</div>
            <ol className="list-decimal pl-4 space-y-0.5 text-[10.5px] leading-snug">
              <li>
                Pilih <strong className="text-blue-700">SEMESTER 1</strong> atau <strong className="text-emerald-700">SEMESTER 2</strong> di atas.
              </li>
              <li>
                Klik ikon <strong className="text-blue-700">Cetak</strong> untuk mencetak otomatis Halaman 1 s.d. 6.
              </li>
              <li>
                Atur ukuran teks &amp; logo menggunakan tombol <strong className="text-blue-700">A- / A+</strong> dan ikon logo.
              </li>
            </ol>
          </div>

          <div className="space-y-2.5">
            <div>
              <label className="block bg-black text-white font-serif font-bold text-[11.5px] px-2 py-0.5">
                Ganti Tempat dan Tanggal
              </label>
              <input
                type="text"
                value={tempatTanggalInput}
                onChange={(e) => handleTempatTanggalChange(e.target.value)}
                placeholder="Contoh: Balil, 23 Desember 2026"
                className="w-full bg-white border-2 border-[#ff0000] px-2 py-1 text-[11.5px] font-semibold text-black focus:outline-none"
              />
            </div>

            <div>
              <label className="block bg-black text-white font-serif font-bold text-[11.5px] px-2 py-0.5">
                Pengesahan Oleh
              </label>
              <input
                type="text"
                list="pengesahan-view-options"
                value={pengesahanOleh}
                onChange={(e) => setPengesahanOleh(e.target.value)}
                placeholder="Kepala Sekolah"
                className="w-full bg-white border-2 border-[#ff0000] px-2 py-1 text-[11.5px] font-semibold text-black focus:outline-none"
              />
              <datalist id="pengesahan-view-options">
                <option value="Kepala Sekolah" />
                <option value="Plt. Kepala Sekolah" />
                <option value="Komite Sekolah" />
              </datalist>
            </div>

            {/* KONTROL BESAR / KECIL KERTAS & BATAS KIRI, KANAN, ATAS, BAWAH */}
            <div className="bg-white border-2 border-[#0022dd] p-2.5 shadow-md space-y-2 text-[11px] text-black">
              <div className="flex items-center justify-between border-b border-slate-300 pb-1">
                <span className="font-serif font-bold text-[11.5px] text-slate-900">
                  Atur Kertas &amp; Batas (Print)
                </span>
                <button
                  type="button"
                  onClick={resetPaperLayout}
                  className="text-[10px] font-bold text-blue-700 hover:underline cursor-pointer"
                  title="Kembalikan margin & ukuran kertas standar"
                >
                  Reset
                </button>
              </div>

              <div className="flex items-center justify-between gap-1 bg-amber-50 border border-amber-200 px-2 py-1 rounded">
                <span className="font-bold text-[10.5px] text-amber-950">Ukuran Kertas</span>
                <div className="inline-flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => adjustPaperScale(-5)}
                    className="w-6 h-6 inline-flex items-center justify-center bg-white hover:bg-slate-100 border border-slate-400 rounded font-black text-xs cursor-pointer"
                    title="Perkecil kertas (-5%)"
                  >
                    -
                  </button>
                  <span className="w-10 text-center font-mono font-black text-[11px] text-amber-950">
                    {printScale}%
                  </span>
                  <button
                    type="button"
                    onClick={() => adjustPaperScale(5)}
                    className="w-6 h-6 inline-flex items-center justify-center bg-amber-600 hover:bg-amber-700 text-white rounded font-black text-xs cursor-pointer"
                    title="Perbesar kertas (+5%)"
                  >
                    +
                  </button>
                </div>
              </div>

              <div className="space-y-1.5">
                {(
                  [
                    { side: 'left', label: 'Samping Kiri' },
                    { side: 'right', label: 'Samping Kanan' },
                    { side: 'top', label: 'Bagian Atas' },
                    { side: 'bottom', label: 'Bagian Bawah' }
                  ] as const
                ).map(({ side, label }) => (
                  <div key={side} className="flex items-center justify-between gap-1 bg-slate-50 border border-slate-200 px-2 py-1 rounded">
                    <span className="font-semibold text-[10.5px] text-slate-800">{label}</span>
                    <div className="inline-flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => adjustPaperMargin(side, -2)}
                        className="w-6 h-5 inline-flex items-center justify-center bg-white hover:bg-slate-200 border border-slate-400 rounded font-black text-xs cursor-pointer"
                        title={`Perkecil ${label} (-2mm)`}
                      >
                        -
                      </button>
                      <span className="w-10 text-center font-mono font-bold text-[10.5px] text-indigo-900">
                        {paperMargins[side]}mm
                      </span>
                      <button
                        type="button"
                        onClick={() => adjustPaperMargin(side, 2)}
                        className="w-6 h-5 inline-flex items-center justify-center bg-blue-600 hover:bg-blue-700 text-white rounded font-black text-xs cursor-pointer"
                        title={`Perbesar ${label} (+2mm)`}
                      >
                        +
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </aside>
      </div>

      {/* POP-UP MODAL INPUT KEPUTUSAN */}
      {openDecisionModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200 animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-slate-900 text-white p-4 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Award size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm">
                  Input Keputusan Rapor Kenaikan Kelas
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setOpenDecisionModal(false)}
                className="text-white/70 hover:text-white cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleSaveDecision} className="p-5 space-y-4 text-xs">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <span className="text-[11px] text-slate-500 font-medium">Siswa Penerima:</span>
                <div className="font-bold text-sm text-slate-900 mt-0.5">
                  {currentStudent.namaSiswa}
                </div>
                <div className="text-[10px] text-slate-500 font-mono">
                  NISN: {currentStudent.nisn} • Kelas {selectedKelas} • Semester {selectedSemester}
                </div>
              </div>

              {/* Pilihan Cepat Keputusan */}
              <div>
                <label className="block font-bold text-slate-700 mb-1.5">
                  Pilih Penetapan Keputusan:
                </label>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => handleSelectStatus('Naik Kelas')}
                    className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all ${
                      decisionForm.status === 'Naik Kelas'
                        ? 'bg-emerald-50 border-emerald-500 text-emerald-900 font-bold shadow-2xs'
                        : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="font-bold text-xs flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                      <span>Naik Kelas</span>
                    </div>
                    <div className="text-[10px] text-slate-500 mt-0.5">
                      Ke Kelas {getNextClass(selectedKelas)}
                    </div>
                  </button>

                  <button
                    type="button"
                    onClick={() => handleSelectStatus('Tinggal di Kelas')}
                    className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all ${
                      decisionForm.status === 'Tinggal di Kelas'
                        ? 'bg-rose-50 border-rose-500 text-rose-900 font-bold shadow-2xs'
                        : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="font-bold text-xs flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-rose-500"></span>
                      <span>Tinggal di Kelas</span>
                    </div>
                    <div className="text-[10px] text-slate-500 mt-0.5">
                      Tetap di Kelas {getCurrentClassText(selectedKelas)}
                    </div>
                  </button>

                  {selectedKelas === 'VI' && (
                    <button
                      type="button"
                      onClick={() => handleSelectStatus('Lulus')}
                      className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all col-span-2 ${
                        decisionForm.status === 'Lulus'
                          ? 'bg-blue-50 border-blue-500 text-blue-900 font-bold shadow-2xs'
                          : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <div className="font-bold text-xs flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-blue-500"></span>
                        <span>Lulus dari Satuan Pendidikan</span>
                      </div>
                      <div className="text-[10px] text-slate-500 mt-0.5">
                        Tamatan Kelas VI SD Negeri Balil
                      </div>
                    </button>
                  )}

                  <button
                    type="button"
                    onClick={() => handleSelectStatus('Belum Ditetapkan')}
                    className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all col-span-2 ${
                      decisionForm.status === 'Belum Ditetapkan'
                        ? 'bg-amber-50 border-amber-500 text-amber-900 font-bold shadow-2xs'
                        : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="font-bold text-xs flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-slate-400"></span>
                      <span>Format Titik-Titik (Diisi Manual)</span>
                    </div>
                    <div className="text-[10px] text-slate-500 mt-0.5">
                      Dikosongkan bergaris titik-titik untuk ditandatangani manual
                    </div>
                  </button>
                </div>
              </div>

              {/* Rincian Kolom Teks */}
              <div className="space-y-3 pt-1">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Teks &ldquo;Naik Kelas :&rdquo;
                  </label>
                  <input
                    type="text"
                    value={decisionForm.naikKelas}
                    onChange={(e) =>
                      setDecisionForm((prev) => ({ ...prev, naikKelas: e.target.value }))
                    }
                    placeholder="Contoh: V (Lima) atau -"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-bold text-slate-900 bg-white"
                  />
                  <span className="text-[10px] text-slate-400 mt-0.5 block">
                    Isi dengan nama kelas tujuan atau tanda &apos;-&apos; jika tinggal kelas.
                  </span>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Teks &ldquo;Tinggal di Kelas :&rdquo;
                  </label>
                  <input
                    type="text"
                    value={decisionForm.tinggalKelas}
                    onChange={(e) =>
                      setDecisionForm((prev) => ({ ...prev, tinggalKelas: e.target.value }))
                    }
                    placeholder="Contoh: - atau IV (Empat)"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-bold text-slate-900 bg-white"
                  />
                  <span className="text-[10px] text-slate-400 mt-0.5 block">
                    Isi dengan &apos;-&apos; jika siswa naik kelas, atau nama kelas saat ini jika tinggal kelas.
                  </span>
                </div>
              </div>

              {/* Opsi Terapkan ke Seluruh Kelas */}
              <div className="p-3 bg-indigo-50/70 border border-indigo-200 rounded-xl">
                <label className="flex items-start gap-2.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={decisionForm.applyToAll}
                    onChange={(e) =>
                      setDecisionForm((prev) => ({ ...prev, applyToAll: e.target.checked }))
                    }
                    className="mt-0.5 rounded text-indigo-600 focus:ring-indigo-500"
                  />
                  <div>
                    <span className="font-bold text-indigo-950 block">
                      Terapkan keputusan ini untuk SEMUA ({siswaKelas.length}) siswa di Kelas {selectedKelas}
                    </span>
                    <span className="text-[10px] text-indigo-700 block mt-0.5">
                      Hemat waktu: Semua siswa di kelas ini akan langsung otomatis mendapatkan keputusan kenaikan kelas yang sama.
                    </span>
                  </div>
                </label>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-200">
                <button
                  type="button"
                  onClick={() => setOpenDecisionModal(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold rounded-xl shadow-md cursor-pointer flex items-center gap-1.5"
                >
                  <Save size={14} />
                  <span>Simpan Keputusan</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
