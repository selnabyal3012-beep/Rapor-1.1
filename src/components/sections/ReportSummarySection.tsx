import React, { useMemo, useState, useEffect, useRef } from 'react';
import {
  Home,
  ChevronUp,
  ChevronDown,
  Printer,
  Download,
  FileSpreadsheet,
  SlidersHorizontal,
  X,
  Upload
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp } from '../../context/AppContext';
import { DEFAULT_KKM_SETTINGS } from '../../data/defaultData';
import { AcademicScore, Student } from '../../types';
import { matchesStudentScore } from '../../utils/studentSyncHelper';
import { parseAndImportTemplate } from '../../utils/templateImporter';

interface ChartSubjectCol {
  id: string;
  label: string;
  group?: 'seni' | 'mulok';
  matchNames: string[];
  isAgama?: boolean;
  defaultKkm: number;
  visible: boolean;
}

export const ReportSummarySection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    setSelectedStudentFilter,
    setPrintTargetStudentId,
    setActiveMenu,
    updateSekolah,
    showSaveNotification
  } = appContext;

  const kkmSettings = db.kkmSettings || DEFAULT_KKM_SETTINGS;
  const importInputRef = useRef<HTMLInputElement | null>(null);

  const [seniVisible, setSeniVisible] = useState<
    Record<'seniMusik' | 'seniTari' | 'seniRupa' | 'seniTeater', boolean>
  >({
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

  const [showSettingsModal, setShowSettingsModal] = useState<boolean>(false);
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

  const siswaKelas = useMemo(
    () => db.siswa.filter((s) => s.kelas === selectedKelas && s.statusSiswa === 'Aktif'),
    [db.siswa, selectedKelas]
  );

  const [selectedIndexOneBased, setSelectedIndexOneBased] = useState<number>(1);

  useEffect(() => {
    if (filteredStudent) {
      const idx = siswaKelas.findIndex((s) => s.idSiswa === filteredStudent.idSiswa);
      if (idx >= 0) {
        setSelectedIndexOneBased(idx + 1);
        return;
      }
    }
    if (selectedIndexOneBased > Math.max(1, siswaKelas.length)) {
      setSelectedIndexOneBased(1);
    }
  }, [filteredStudent, siswaKelas]);

  const safeIndex =
    siswaKelas.length > 0
      ? Math.max(1, Math.min(siswaKelas.length, selectedIndexOneBased))
      : 1;
  const currentStudent: Student | undefined = siswaKelas[safeIndex - 1];

  const handleStepStudent = (nextIndex: number) => {
    if (siswaKelas.length === 0) return;
    const bounded = Math.max(1, Math.min(siswaKelas.length, nextIndex));
    setSelectedIndexOneBased(bounded);
    const target = siswaKelas[bounded - 1];
    if (target) {
      setPrintTargetStudentId(target.idSiswa);
    }
  };

  const handleTempatTanggalChange = (val: string) => {
    setTempatTanggalInput(val);
    updateSekolah({
      ...db.sekolah,
      tempatTanggalRapor: val
    });
  };

  const hasIpas = useMemo(
    () =>
      db.mataPelajaran.some(
        (m) =>
          m.mataPelajaran.toLowerCase().includes('ipas') ||
          m.mataPelajaran.toLowerCase().includes('ilmu pengetahuan alam dan sosial')
      ),
    [db.mataPelajaran]
  );

  const allColumns = useMemo<ChartSubjectCol[]>(() => {
    const mulokNames = kkmSettings.muatanLokalNames || [];
    return [
      {
        id: 'agama',
        label: 'Pendidikan Agama dan Budi Pekerti',
        matchNames: ['pendidikan agama dan budi pekerti'],
        isAgama: true,
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'pancasila',
        label: 'Pendidikan Pancasila',
        matchNames: ['pendidikan pancasila', 'ppkn'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'bindo',
        label: 'Bahasa Indonesia',
        matchNames: ['bahasa indonesia'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'mtk',
        label: 'Matematika',
        matchNames: ['matematika'],
        defaultKkm: 75,
        visible: true
      },
      ...(hasIpas
        ? [
            {
              id: 'ipas',
              label: 'Ilmu Pengetahuan Alam dan Sosial',
              matchNames: [
                'ilmu pengetahuan alam dan sosial (ipas)',
                'ilmu pengetahuan alam dan sosial',
                'ipas'
              ],
              defaultKkm: 75,
              visible: true
            }
          ]
        : []),
      {
        id: 'seni_musik',
        label: 'Seni Musik',
        group: 'seni',
        matchNames: ['seni musik'],
        defaultKkm: 75,
        visible: seniVisible.seniMusik
      },
      {
        id: 'seni_tari',
        label: 'Seni Tari',
        group: 'seni',
        matchNames: ['seni tari'],
        defaultKkm: 75,
        visible: seniVisible.seniTari
      },
      {
        id: 'seni_rupa',
        label: 'Seni Rupa',
        group: 'seni',
        matchNames: ['seni rupa', 'seni dan budaya (seni rupa & musik)'],
        defaultKkm: 75,
        visible: seniVisible.seniRupa
      },
      {
        id: 'seni_teater',
        label: 'Seni Teater',
        group: 'seni',
        matchNames: ['seni teater'],
        defaultKkm: 75,
        visible: seniVisible.seniTeater
      },
      {
        id: 'pjok',
        label: 'Pendidikan Jasmani, Olahraga dan Kesehatan',
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
        matchNames: ['bahasa inggris'],
        defaultKkm: 75,
        visible: true
      },
      {
        id: 'mulok_1',
        label: mulokNames[0]?.trim() || 'Muatan Lokal 1',
        group: 'mulok',
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
        label: mulokNames[1]?.trim() || 'Muatan Lokal 2',
        group: 'mulok',
        matchNames: [(mulokNames[1] || '').trim().toLowerCase(), 'muatan lokal 2', 'mulok 2'].filter(
          Boolean
        ),
        defaultKkm: 75,
        visible: mulokVisible[1]
      },
      {
        id: 'mulok_3',
        label: mulokNames[2]?.trim() || 'Muatan Lokal 3',
        group: 'mulok',
        matchNames: [(mulokNames[2] || '').trim().toLowerCase(), 'muatan lokal 3', 'mulok 3'].filter(
          Boolean
        ),
        defaultKkm: 75,
        visible: mulokVisible[2]
      },
      {
        id: 'mulok_4',
        label: mulokNames[3]?.trim() || 'Muatan Lokal 4',
        group: 'mulok',
        matchNames: [(mulokNames[3] || '').trim().toLowerCase(), 'muatan lokal 4', 'mulok 4'].filter(
          Boolean
        ),
        defaultKkm: 75,
        visible: mulokVisible[3]
      },
      {
        id: 'mulok_5',
        label: mulokNames[4]?.trim() || 'Muatan Lokal 5',
        group: 'mulok',
        matchNames: [(mulokNames[4] || '').trim().toLowerCase(), 'muatan lokal 5', 'mulok 5'].filter(
          Boolean
        ),
        defaultKkm: 75,
        visible: mulokVisible[4]
      }
    ];
  }, [hasIpas, kkmSettings.muatanLokalNames, mulokVisible, seniVisible]);

  const visibleColumns = useMemo(() => allColumns.filter((c) => c.visible), [allColumns]);

  const getColumnKkm = (col: ChartSubjectCol): number => {
    const found = db.mataPelajaran.find((m) => {
      const mName = m.mataPelajaran.trim().toLowerCase();
      if (col.isAgama && mName.includes('agama')) return true;
      return col.matchNames.some((mn) => mn === mName);
    });
    return found?.kkm ?? col.defaultKkm;
  };

  const findStudentColScore = (
    student: Student,
    col: ChartSubjectCol
  ): AcademicScore | undefined => {
    const studentScores = db.nilaiAkademik.filter(
      (n) => matchesStudentScore(n, student) && n.semester === selectedSemester
    );

    if (col.isAgama) {
      const agamaName = (student.agama || '').trim().toLowerCase();
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

  // Hitung pasangan data grafik per mata pelajaran:
  // 1) CAPAIAN SISWA (nilaiAkhir siswa terpilih)
  // 2) DAYA SERAP (%) (persentase siswa kelas yang tuntas >= KKM)
  const chartData = useMemo(() => {
    return visibleColumns.map((col) => {
      const kkm = getColumnKkm(col);
      const classValues = siswaKelas
        .map((s) => {
          const rec = findStudentColScore(s, col);
          return rec && rec.nilaiAkhir > 0 ? rec.nilaiAkhir : null;
        })
        .filter((v): v is number => v !== null);

      const tuntasCount = classValues.filter((v) => v >= kkm).length;
      const dayaSerap =
        classValues.length > 0 ? Math.round((tuntasCount / classValues.length) * 100) : 0;

      const currentRec = currentStudent ? findStudentColScore(currentStudent, col) : undefined;
      const capaianSiswa =
        currentRec && currentRec.nilaiAkhir > 0 ? Math.round(currentRec.nilaiAkhir) : 0;

      return {
        id: col.id,
        label: col.label,
        kkm,
        capaianSiswa,
        dayaSerap
      };
    });
  }, [visibleColumns, siswaKelas, currentStudent, db.nilaiAkademik, selectedSemester, db.mataPelajaran]);

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const rows: (string | number)[][] = [
      [`PERBANDINGAN CAPAIAN PRESTASI SISWA DENGAN DAYA SERAP KELAS ${selectedKelas}`],
      [`TAHUN PELAJARAN ${db.sekolah.tahunPelajaran || ''}`],
      [`NAMA : ${currentStudent?.namaSiswa || '-'}`],
      [`NISN/NIS : ${currentStudent?.nisn || '-'} / ${currentStudent?.nis || '-'}`],
      [],
      ['No', 'Mata Pelajaran', 'KKM', 'Capaian Siswa', 'Daya Serap (%)'],
      ...chartData.map((d, idx) => [idx + 1, d.label, d.kkm, d.capaianSiswa, d.dayaSerap])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Grafik Rekap Rapor');
    XLSX.writeFile(
      wb,
      `grafik-rekap-rapor-${currentStudent?.namaSiswa || 'kelas-' + selectedKelas}.xlsx`
    );
    showSaveNotification('Grafik Rekap Rapor berhasil diekspor ke Excel!');
  };

  const handlePrint = () => {
    window.print();
  };

  // Tinggi area batang chart (px untuk skor 100)
  const MAX_BAR_HEIGHT = 210;

  return (
    <div className="bg-[#00e5ff] min-h-[calc(100vh-5rem)] p-2 sm:p-4 rounded-xl print:bg-white print:p-0">
      <style>{`
        @media print {
          @page {
            size: A4 landscape;
            margin: 6mm;
          }
          .grafik-printable-area {
            width: 100% !important;
            max-width: 100% !important;
            min-height: 190mm !important;
            box-shadow: none !important;
            margin: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
        }
      `}</style>

      <div className="flex flex-col xl:flex-row items-stretch justify-between gap-4">
        {/* KIRI: HALAMAN PUTIH GRAFIK PERBANDINGAN CAPAIAN & DAYA SERAP */}
        <div className="w-full xl:flex-1 overflow-x-auto pb-2 print:overflow-visible print:pb-0">
          <div
            className="w-[297mm] min-h-[200mm] bg-white text-black shadow-[0_8px_28px_rgba(0,0,0,0.2)] mx-auto px-8 sm:px-12 pt-5 pb-8 flex flex-col justify-between printable-area grafik-printable-area"
            style={{
              fontFamily: '"Tinos", "Times New Roman", Times, serif',
              WebkitPrintColorAdjust: 'exact',
              printColorAdjust: 'exact'
            }}
          >
            {/* JUDUL & IDENTITAS SISWA */}
            <div>
              <div className="text-center space-y-3">
                <h1 className="text-[17px] sm:text-[19px] font-bold uppercase tracking-wide text-black">
                  PERBANDINGAN CAPAIAN PRESTASI SISWA DENGAN DAYA SERAP KELAS {selectedKelas}
                </h1>
                <h2 className="text-[16px] sm:text-[18px] font-bold uppercase tracking-wide text-black">
                  TAHUN PELAJARAN {db.sekolah.tahunPelajaran || ''}
                </h2>
              </div>

              {/* NAMA & NISN/NIS */}
              <div className="mt-3 text-[13px] sm:text-[14px] font-bold text-black space-y-1">
                <div className="flex items-center">
                  <span className="w-28 inline-block uppercase">NAMA</span>
                  <span className="mr-2">:</span>
                  <span className="uppercase">{currentStudent?.namaSiswa || ''}</span>
                </div>
                <div className="flex items-center">
                  <span className="w-28 inline-block uppercase">NISN/NIS</span>
                  <span className="mr-2">:</span>
                  <span>
                    {currentStudent?.nisn || ''} / {currentStudent?.nis || ''}
                  </span>
                </div>
              </div>
            </div>

            {/* GRAFIK KOLOM 3D CAPAIAN SISWA VS DAYA SERAP (%) */}
            <div className="my-4 px-2 sm:px-6 flex flex-col items-center">
              {/* Legenda */}
              <div
                className="flex items-center justify-center gap-6 mb-5 text-[11px] font-semibold text-black"
                style={{ fontFamily: 'Arial, Helvetica, sans-serif' }}
              >
                <div className="flex items-center gap-1.5">
                  <span className="inline-block w-2.5 h-2.5 bg-[#00e600] border border-black" />
                  <span>CAPAIAN SISWA</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="inline-block w-2.5 h-2.5 bg-[#ffff00] border border-black" />
                  <span>DAYA SERAP (%)</span>
                </div>
              </div>

              {/* Area Grafik 3D */}
              <div className="w-full max-w-[920px]">
                {/* Baris Batang 3D */}
                <div
                  className="relative flex items-end justify-around px-5 pt-6"
                  style={{ height: `${MAX_BAR_HEIGHT + 30}px` }}
                >
                  {chartData.map((col) => {
                    const hStudent = Math.round((Math.min(100, Math.max(0, col.capaianSiswa)) / 100) * MAX_BAR_HEIGHT);
                    const hSerap = Math.round((Math.min(100, Math.max(0, col.dayaSerap)) / 100) * MAX_BAR_HEIGHT);

                    return (
                      <div
                        key={col.id}
                        className="flex items-end justify-center gap-[3px] relative z-10 px-1"
                        style={{ minWidth: '40px' }}
                        title={`${col.label}\nCapaian Siswa: ${col.capaianSiswa}\nDaya Serap Kelas: ${col.dayaSerap}%`}
                      >
                        {/* Kolom 1: CAPAIAN SISWA (Hijau 3D) */}
                        <div className="flex flex-col items-center">
                          <span
                            className="text-[9.5px] font-bold text-black mb-1 leading-none"
                            style={{ fontFamily: 'Arial, Helvetica, sans-serif' }}
                          >
                            {col.capaianSiswa}
                          </span>
                          <div className="relative" style={{ width: '16px', height: `${Math.max(4, hStudent + 4)}px` }}>
                            {/* Top 3D Diamond Cap */}
                            <div
                              style={{
                                position: 'absolute',
                                top: 0,
                                left: 0,
                                width: '16px',
                                height: '5px',
                                backgroundColor: '#52ff52',
                                border: '1px solid #0f172a',
                                transform: 'skewX(-45deg)',
                                transformOrigin: 'bottom left',
                                zIndex: 3
                              }}
                            />
                            {/* Front Face */}
                            {hStudent > 0 && (
                              <>
                                <div
                                  style={{
                                    position: 'absolute',
                                    top: '4px',
                                    left: 0,
                                    width: '16px',
                                    height: `${hStudent}px`,
                                    backgroundColor: '#00e600',
                                    border: '1px solid #0f172a',
                                    zIndex: 2
                                  }}
                                />
                                {/* Right 3D Side Face */}
                                <div
                                  style={{
                                    position: 'absolute',
                                    top: '2px',
                                    left: '15px',
                                    width: '5px',
                                    height: `${hStudent}px`,
                                    backgroundColor: '#009900',
                                    border: '1px solid #0f172a',
                                    transform: 'skewY(-45deg)',
                                    transformOrigin: 'top left',
                                    zIndex: 1
                                  }}
                                />
                              </>
                            )}
                          </div>
                        </div>

                        {/* Kolom 2: DAYA SERAP (%) (Kuning 3D) */}
                        <div className="flex flex-col items-center">
                          <span
                            className="text-[9.5px] font-bold text-black mb-1 leading-none"
                            style={{ fontFamily: 'Arial, Helvetica, sans-serif' }}
                          >
                            {col.dayaSerap}
                          </span>
                          <div className="relative" style={{ width: '16px', height: `${Math.max(4, hSerap + 4)}px` }}>
                            {/* Top 3D Diamond Cap */}
                            <div
                              style={{
                                position: 'absolute',
                                top: 0,
                                left: 0,
                                width: '16px',
                                height: '5px',
                                backgroundColor: '#ffff80',
                                border: '1px solid #0f172a',
                                transform: 'skewX(-45deg)',
                                transformOrigin: 'bottom left',
                                zIndex: 3
                              }}
                            />
                            {/* Front Face */}
                            {hSerap > 0 && (
                              <>
                                <div
                                  style={{
                                    position: 'absolute',
                                    top: '4px',
                                    left: 0,
                                    width: '16px',
                                    height: `${hSerap}px`,
                                    backgroundColor: '#ffff00',
                                    border: '1px solid #0f172a',
                                    zIndex: 2
                                  }}
                                />
                                {/* Right 3D Side Face */}
                                <div
                                  style={{
                                    position: 'absolute',
                                    top: '2px',
                                    left: '15px',
                                    width: '5px',
                                    height: `${hSerap}px`,
                                    backgroundColor: '#ca8a04',
                                    border: '1px solid #0f172a',
                                    transform: 'skewY(-45deg)',
                                    transformOrigin: 'top left',
                                    zIndex: 1
                                  }}
                                />
                              </>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Lantai Dasar 3D Perspektif Biru */}
                <div className="relative w-full -mt-1 z-0">
                  <svg viewBox="0 0 1000 18" className="w-full h-[14px]" preserveAspectRatio="none">
                    <defs>
                      <linearGradient id="floorBlue3d" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#6c8cd5" />
                        <stop offset="55%" stopColor="#3b5998" />
                        <stop offset="100%" stopColor="#1e3a8a" />
                      </linearGradient>
                    </defs>
                    <polygon
                      points="18,1 1000,1 984,12 2,12"
                      fill="url(#floorBlue3d)"
                      stroke="#1e293b"
                      strokeWidth="1.2"
                    />
                    <rect
                      x="2"
                      y="12"
                      width="982"
                      height="5"
                      fill="#1e3a8a"
                      stroke="#0f172a"
                      strokeWidth="1"
                    />
                  </svg>
                </div>

                {/* Label Mata Pelajaran Vertikal di Bawah Lantai 3D */}
                <div
                  className="flex items-start justify-around px-5 pt-1.5"
                  style={{ fontFamily: 'Arial, Helvetica, sans-serif' }}
                >
                  {chartData.map((col) => (
                    <div
                      key={col.id}
                      className="flex justify-center"
                      style={{ minWidth: '40px', maxWidth: '52px' }}
                    >
                      <div
                        className="text-[9px] text-black leading-[1.15] text-right"
                        style={{
                          writingMode: 'vertical-rl',
                          transform: 'rotate(180deg)',
                          maxHeight: '115px'
                        }}
                      >
                        {col.label}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* FOOTER TANDA TANGAN KIRI & KANAN */}
            <div className="grid grid-cols-2 items-end gap-8 px-6 sm:px-14 pt-6 text-[13px] sm:text-[14px] text-black">
              {/* Kiri: Mengetahui Kepala Sekolah */}
              <div className="text-left leading-snug">
                <div>Mengetahui,</div>
                <div>{pengesahanOleh || 'Kepala Sekolah'}</div>
                <div className="h-14 sm:h-16" />
                {db.sekolah.namaKepalaSekolah && (
                  <div className="font-bold underline">{db.sekolah.namaKepalaSekolah}</div>
                )}
                <div>NIP. {db.sekolah.nipKepalaSekolah || ''}</div>
              </div>

              {/* Kanan: Tempat Tanggal + Guru Kelas */}
              <div className="text-left leading-snug pl-14 sm:pl-28">
                <div className="min-h-[20px]">
                  {tempatTanggalInput.trim() !== '' ? tempatTanggalInput : '\u00A0'}
                </div>
                <div>Guru Kelas,</div>
                <div className="h-14 sm:h-16" />
                {db.sekolah.waliKelas && (
                  <div className="font-bold underline">{db.sekolah.waliKelas}</div>
                )}
                <div>NIP. {db.sekolah.nipWaliKelas || ''}</div>
              </div>
            </div>
          </div>
        </div>

        {/* KANAN: PANEL KONTROL HITAM & INPUT TANDA TANGAN (NO-PRINT) */}
        <aside className="w-full xl:w-[225px] shrink-0 flex flex-col justify-between gap-6 no-print print:hidden">
          {/* Kotak Hitam Berbingkai Merah */}
          <div className="bg-black border-[2.5px] border-[#ff0000] p-3 space-y-3 shadow-xl">
            {/* Row 1: Home + Kotak Angka [1] + Spinner ▲▼ */}
            <div className="flex items-center justify-center gap-2">
              <button
                type="button"
                onClick={() => setActiveMenu('dashboard')}
                title="Kembali ke Dashboard"
                className="w-10 h-10 rounded-full bg-[#1d4ed8] hover:bg-[#2563eb] border-2 border-[#38bdf8] flex items-center justify-center text-black shadow-md cursor-pointer shrink-0 transition-transform active:scale-95"
              >
                <Home size={20} strokeWidth={2.5} />
              </button>

              {/* Kotak Indeks Siswa */}
              <div className="flex items-stretch">
                <div
                  className="w-14 h-9 bg-white border-2 border-[#ff0000] flex items-center justify-center font-bold text-black text-base select-none"
                  title="Nomor Urut Siswa di Kelas"
                >
                  {safeIndex}
                </div>
                <div className="flex flex-col w-6 h-9 border border-slate-500 bg-slate-200 ml-1">
                  <button
                    type="button"
                    onClick={() => handleStepStudent(safeIndex + 1)}
                    disabled={safeIndex >= siswaKelas.length}
                    title="Siswa Berikutnya"
                    className="flex-1 flex items-center justify-center hover:bg-slate-300 disabled:opacity-40 text-black border-b border-slate-400 cursor-pointer"
                  >
                    <ChevronUp size={13} strokeWidth={3} />
                  </button>
                  <button
                    type="button"
                    onClick={() => handleStepStudent(safeIndex - 1)}
                    disabled={safeIndex <= 1}
                    title="Siswa Sebelumnya"
                    className="flex-1 flex items-center justify-center hover:bg-slate-300 disabled:opacity-40 text-black cursor-pointer"
                  >
                    <ChevronDown size={13} strokeWidth={3} />
                  </button>
                </div>
              </div>
            </div>

            {/* Row 2: Tombol Cyan ATUR HALAMAN */}
            <div className="flex flex-col items-center gap-2">
              <button
                type="button"
                onClick={() => setShowSettingsModal(true)}
                className="px-5 py-1.5 bg-[#00e5ff] hover:bg-[#40ecff] text-black font-black text-[11px] leading-[1.1] uppercase border border-white/90 shadow-inner cursor-pointer text-center"
              >
                ATUR
                <br />
                HALAMAN
              </button>

              <div className="w-full grid grid-cols-2 gap-1.5 pt-1">
                <button
                  type="button"
                  onClick={handlePrint}
                  className="flex items-center justify-center gap-1 py-1.5 px-2 bg-[#0066ff] hover:bg-[#1d4ed8] text-white font-bold text-[10.5px] border border-sky-300 cursor-pointer"
                >
                  <Printer size={12} />
                  <span>Cetak</span>
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
          </div>

          {/* Bawah Kanan: Ganti Tempat dan Tanggal + Pengesahan Oleh */}
          <div className="space-y-3 mt-auto pt-4">
            <div>
              <label className="block bg-black text-white font-serif font-bold text-[12px] px-2 py-0.5">
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
              <label className="block bg-black text-white font-serif font-bold text-[12px] px-2 py-0.5">
                Pengesahan Oleh
              </label>
              <input
                type="text"
                list="pengesahan-grafik-options"
                value={pengesahanOleh}
                onChange={(e) => setPengesahanOleh(e.target.value)}
                placeholder="Kepala Sekolah"
                className="w-full bg-white border-2 border-[#ff0000] px-2 py-1 text-[12px] font-semibold text-black focus:outline-none"
              />
              <datalist id="pengesahan-grafik-options">
                <option value="Kepala Sekolah" />
                <option value="Plt. Kepala Sekolah" />
                <option value="Komite Sekolah" />
              </datalist>
            </div>
          </div>
        </aside>
      </div>

      {/* MODAL ATUR HALAMAN, PILIH SISWA, TABEL RINGKAS & EKSPOR */}
      {showSettingsModal && (
        <div className="fixed inset-0 z-50 bg-black/60 flex items-center justify-center p-4 print:hidden">
          <div className="bg-white rounded-xl max-w-2xl w-full shadow-2xl border border-slate-200 overflow-hidden max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between px-5 py-3.5 bg-[#005a9e] text-white">
              <div className="flex items-center gap-2">
                <SlidersHorizontal size={18} />
                <h3 className="font-bold text-sm uppercase tracking-wide">
                  Atur Halaman Grafik Rekap Nilai Rapor
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

            <div className="p-5 overflow-y-auto space-y-4 text-xs text-slate-800">
              {/* Pilih Siswa */}
              <div>
                <label className="block font-bold text-slate-700 mb-1">
                  Pilih Siswa Kelas {selectedKelas} (Semester {selectedSemester})
                </label>
                <select
                  value={currentStudent?.idSiswa || ''}
                  onChange={(e) => {
                    const idx = siswaKelas.findIndex((s) => s.idSiswa === e.target.value);
                    if (idx >= 0) {
                      handleStepStudent(idx + 1);
                      setSelectedStudentFilter(e.target.value);
                    }
                  }}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg font-bold bg-white"
                >
                  {siswaKelas.map((s, idx) => (
                    <option key={s.idSiswa} value={s.idSiswa}>
                      {idx + 1}. {s.namaSiswa} (NISN: {s.nisn || '-'})
                    </option>
                  ))}
                </select>
              </div>

              {/* Filter Kolom Seni & Mulok */}
              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 space-y-2">
                <div className="font-bold text-slate-700 uppercase text-[11px]">
                  Filter Mata Pelajaran Seni & Muatan Lokal di Grafik
                </div>
                <div className="flex flex-wrap gap-3">
                  {(
                    [
                      ['seniMusik', 'Seni Musik'],
                      ['seniTari', 'Seni Tari'],
                      ['seniRupa', 'Seni Rupa'],
                      ['seniTeater', 'Seni Teater']
                    ] as const
                  ).map(([key, lbl]) => (
                    <label key={key} className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={seniVisible[key]}
                        onChange={(e) =>
                          setSeniVisible((prev) => ({ ...prev, [key]: e.target.checked }))
                        }
                      />
                      <span>{lbl}</span>
                    </label>
                  ))}
                  {[0, 1, 2, 3, 4].map((idx) => (
                    <label key={idx} className="flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={mulokVisible[idx]}
                        onChange={(e) =>
                          setMulokVisible((prev) =>
                            prev.map((v, i) => (i === idx ? e.target.checked : v))
                          )
                        }
                      />
                      <span>Mulok {idx + 1}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Ringkasan Nilai Grafik */}
              <div className="border border-slate-200 rounded-lg overflow-hidden">
                <table className="w-full text-xs text-left">
                  <thead className="bg-slate-800 text-white text-[10px] uppercase">
                    <tr>
                      <th className="px-3 py-2">No</th>
                      <th className="px-3 py-2">Mata Pelajaran</th>
                      <th className="px-3 py-2 text-center">KKM</th>
                      <th className="px-3 py-2 text-center">Capaian Siswa</th>
                      <th className="px-3 py-2 text-center">Daya Serap (%)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {chartData.map((d, idx) => (
                      <tr key={d.id} className="hover:bg-slate-50">
                        <td className="px-3 py-1.5 font-mono text-slate-500">{idx + 1}</td>
                        <td className="px-3 py-1.5 font-medium">{d.label}</td>
                        <td className="px-3 py-1.5 text-center font-mono">{d.kkm}</td>
                        <td className="px-3 py-1.5 text-center font-bold text-emerald-700 font-mono">
                          {d.capaianSiswa}
                        </td>
                        <td className="px-3 py-1.5 text-center font-bold text-amber-700 font-mono">
                          {d.dayaSerap}%
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Aksi Impor / Ekspor / Cetak */}
              <div className="flex flex-wrap items-center justify-between gap-2 border-t pt-3">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => importInputRef.current?.click()}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-emerald-300 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-semibold cursor-pointer"
                  >
                    <Upload size={13} />
                    <span>Impor Excel</span>
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
                        const res = await parseAndImportTemplate(file, undefined, appContext);
                        showSaveNotification(
                          res.success ? res.message : `Gagal Impor: ${res.message}`
                        );
                      } catch (err: unknown) {
                        const msg = err instanceof Error ? err.message : 'Format tidak valid';
                        showSaveNotification('Gagal mengimpor file: ' + msg);
                      }
                      e.target.value = '';
                    }}
                  />
                  <button
                    type="button"
                    onClick={handleExportExcel}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-sky-300 bg-sky-50 hover:bg-sky-100 text-sky-900 font-semibold cursor-pointer"
                  >
                    <FileSpreadsheet size={13} />
                    <span>Ekspor Excel</span>
                  </button>
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
                    <span>Cetak Grafik A4</span>
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
