import React, { useMemo, useState } from 'react';
import {
  Home,
  HelpCircle,
  Trash2,
  FileSpreadsheet,
  Hand,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp } from '../../context/AppContext';
import { ClassLevel, Student, StudentDecision } from '../../types';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { matchesStudentRecord } from '../../utils/studentSyncHelper';

const MAX_SHEET_ROWS = 60;

function getNextClassText(current: ClassLevel): string {
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
      return 'Lulus (Tamat SD/MI)';
    case 'VII':
      return 'VIII (Delapan)';
    case 'VIII':
      return 'IX (Sembilan)';
    case 'IX':
      return 'Lulus (Tamat SMP/MTs)';
    case 'X':
      return 'XI (Sebelas)';
    case 'XI':
      return 'XII (Dua Belas)';
    case 'XII':
      return 'Lulus (Tamat SMA/SMK)';
    case 'XIII':
      return 'Lulus (Tamat SMK 4 Tahun)';
    default:
      return '-';
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

function isGraduatingLevel(kelas: ClassLevel): boolean {
  return kelas === 'VI' || kelas === 'IX' || kelas === 'XII' || kelas === 'XIII';
}

export const KeteranganNaikKelasSection: React.FC = () => {
  const {
    db,
    selectedKelas,
    filteredStudent,
    saveKeputusanSiswa,
    bulkSaveKeputusanSiswa,
    showSaveNotification,
    setActiveMenu,
    openClearDataModal
  } = useApp();

  const [showHelp, setShowHelp] = useState(false);

  const siswaKelas = useMemo(
    () => db.siswa.filter((s) => s.kelas === selectedKelas),
    [db.siswa, selectedKelas]
  );

  const displayedStudents = useMemo(
    () => (filteredStudent ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa) : siswaKelas),
    [filteredStudent, siswaKelas]
  );

  const totalRowSlots = filteredStudent
    ? Math.max(displayedStudents.length, 1)
    : Math.max(MAX_SHEET_ROWS, displayedStudents.length);

  // Keterangan Naik/Tidak Naik khusus Semester 2 ('Genap' / '2')
  const getStudentDecision = (student: Student): StudentDecision | undefined => {
    return (db.keputusanSiswa || []).find(
      (k) =>
        matchesStudentRecord(k, student) &&
        k.kelas === selectedKelas &&
        (k.semester === 'Genap' || k.semester === '2')
    );
  };

  const getOptionValue = (decision?: StudentDecision): '' | 'Ya' | 'Tidak' => {
    if (!decision) return '';
    if (decision.statusKeputusan === 'Naik Kelas' || decision.statusKeputusan === 'Lulus') {
      return 'Ya';
    }
    if (decision.statusKeputusan === 'Tinggal di Kelas') {
      return 'Tidak';
    }
    return '';
  };

  const buildDecisionRecord = (student: Student, choice: '' | 'Ya' | 'Tidak'): StudentDecision => {
    if (choice === 'Ya') {
      const graduating = isGraduatingLevel(selectedKelas);
      return {
        id: `DEC_${student.idSiswa}_${selectedKelas}_Genap`,
        idSiswa: student.idSiswa,
        nisn: student.nisn,
        namaSiswa: student.namaSiswa,
        kelas: selectedKelas,
        semester: 'Genap',
        tahunPelajaran: db.sekolah.tahunPelajaran,
        statusKeputusan: graduating ? 'Lulus' : 'Naik Kelas',
        naikKelas: getNextClassText(selectedKelas),
        tinggalKelas: '-'
      };
    }

    if (choice === 'Tidak') {
      return {
        id: `DEC_${student.idSiswa}_${selectedKelas}_Genap`,
        idSiswa: student.idSiswa,
        nisn: student.nisn,
        namaSiswa: student.namaSiswa,
        kelas: selectedKelas,
        semester: 'Genap',
        tahunPelajaran: db.sekolah.tahunPelajaran,
        statusKeputusan: 'Tinggal di Kelas',
        naikKelas: '-',
        tinggalKelas: getCurrentClassText(selectedKelas)
      };
    }

    return {
      id: `DEC_${student.idSiswa}_${selectedKelas}_Genap`,
      idSiswa: student.idSiswa,
      nisn: student.nisn,
      namaSiswa: student.namaSiswa,
      kelas: selectedKelas,
      semester: 'Genap',
      tahunPelajaran: db.sekolah.tahunPelajaran,
      statusKeputusan: 'Belum Ditetapkan',
      naikKelas: '........................................................',
      tinggalKelas: '........................................................'
    };
  };

  const handleSelectChange = (student: Student, choice: '' | 'Ya' | 'Tidak') => {
    saveKeputusanSiswa(buildDecisionRecord(student, choice));
  };

  const handleAllNaikKelas = () => {
    if (siswaKelas.length === 0) {
      showSaveNotification(`Belum ada data siswa di Kelas ${selectedKelas}.`);
      return;
    }
    const batch = siswaKelas.map((s) => buildDecisionRecord(s, 'Ya'));
    bulkSaveKeputusanSiswa(batch);
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['KETERANGAN NAIK / TIDAK NAIK KELAS (SEMESTER 2)'],
      [`Kelas: ${selectedKelas} • Tahun Pelajaran: ${db.sekolah.tahunPelajaran}`],
      [],
      ['NO', 'NISN', 'NAMA SISWA', 'NAIK KELAS', 'STATUS KEPUTUSAN RAPOR'],
      ...siswaKelas.map((s, idx) => {
        const dec = getStudentDecision(s);
        const val = getOptionValue(dec);
        return [
          idx + 1,
          s.nisn,
          s.namaSiswa,
          val || '-',
          dec?.statusKeputusan || 'Belum Ditetapkan'
        ];
      })
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 6 }, { wch: 16 }, { wch: 30 }, { wch: 14 }, { wch: 24 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Naik Kelas');
    XLSX.writeFile(wb, `keterangan-naik-kelas-${selectedKelas}-semester-2.xlsx`);
    showSaveNotification('Data Keterangan Naik/Tidak Naik Kelas berhasil diekspor ke Excel!');
  };

  const summaryCounts = useMemo(() => {
    let ya = 0;
    let tidak = 0;
    siswaKelas.forEach((s) => {
      const val = getOptionValue(getStudentDecision(s));
      if (val === 'Ya') ya++;
      else if (val === 'Tidak') tidak++;
    });
    return { ya, tidak, belum: Math.max(0, siswaKelas.length - ya - tidak) };
  }, [siswaKelas, db.keputusanSiswa, selectedKelas]);

  return (
    <div className="bg-[#00e5ff] min-h-screen p-2 sm:p-4 rounded-xl shadow-inner">
      {/* TOP BAR */}
      <div className="flex flex-wrap items-start justify-between gap-3 mb-2">
        <div className="pl-2 sm:pl-12">
          <h1 className="text-lg sm:text-xl font-extrabold uppercase tracking-wide text-black leading-tight">
            KETERANGAN NAIK/TIDAK NAIK
          </h1>
          <p className="text-sm sm:text-base font-extrabold uppercase text-red-600 tracking-wide">
            HANYA UNTUK SEMESTER 2
          </p>
        </div>

        {/* BLACK CONTROL PANEL MATCHING SCREENSHOT */}
        <div className="bg-black border-2 border-red-600 px-3 py-2 flex flex-wrap items-center gap-2.5 shadow-lg">
          <button
            onClick={() => setActiveMenu('dashboard')}
            className="w-9 h-9 rounded-full bg-sky-600 hover:bg-sky-500 text-white border border-sky-300 flex items-center justify-center cursor-pointer transition-transform active:scale-95"
            title="Kembali ke Dashboard Utama"
          >
            <Home size={18} />
          </button>

          <button
            onClick={() => setShowHelp((prev) => !prev)}
            className="w-9 h-9 rounded-full bg-purple-700 hover:bg-purple-600 text-white border border-purple-300 flex items-center justify-center cursor-pointer"
            title="Panduan Pengisian Keterangan Naik/Tidak Naik Kelas"
          >
            <HelpCircle size={18} />
          </button>

          <button
            onClick={handleAllNaikKelas}
            className="px-3 py-1.5 bg-[#00ff00] hover:bg-lime-400 text-slate-950 font-extrabold text-[11px] border border-black rounded shadow flex items-center gap-1.5 cursor-pointer leading-tight"
            title="Tetapkan seluruh siswa kelas ini menjadi Naik Kelas (Ya)"
          >
            <Hand size={15} className="text-slate-900 shrink-0" />
            <span className="text-center">
              Semua Siswa
              <br />
              Naik Kelas
            </span>
          </button>

          <button
            onClick={() => openClearDataModal()}
            className="w-9 h-9 rounded-full bg-red-600 hover:bg-red-500 text-white border border-red-300 flex items-center justify-center cursor-pointer"
            title="Bersihkan Data Keputusan Naik/Tidak Naik Kelas"
          >
            <Trash2 size={16} />
          </button>
        </div>
      </div>

      {/* HELP DRAWER */}
      {showHelp && (
        <div className="mb-3 bg-white border-2 border-black p-3 rounded shadow-lg text-xs text-slate-800">
          <div className="flex items-center justify-between border-b border-slate-200 pb-1 mb-1.5">
            <h3 className="font-bold text-slate-900">
              Panduan Keterangan Naik / Tidak Naik Kelas (Kelas {selectedKelas} • Khusus Semester 2)
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
              Pilih <strong>Ya</strong> pada kolom <strong>NAIK KELAS</strong> jika siswa dinyatakan{' '}
              {isGraduatingLevel(selectedKelas) ? 'Lulus' : `naik ke kelas ${getNextClassText(selectedKelas)}`}, atau pilih{' '}
              <strong>Tidak</strong> jika siswa tinggal di kelas {getCurrentClassText(selectedKelas)}.
            </li>
            <li>
              Klik tombol hijau <strong>Semua Siswa Naik Kelas</strong> untuk mengisi pilihan <strong>Ya</strong> secara otomatis ke seluruh siswa di kelas {selectedKelas}.
            </li>
            <li>
              Setiap pilihan langsung tersimpan untuk <strong>Semester 2 (Genap)</strong> dan otomatis muncul pada kolom Keputusan Kenaikan Kelas di menu <em>Cetak Rapor Semester 2</em> &amp; <em>Buku Induk</em>.
            </li>
          </ul>
        </div>
      )}

      {/* MAIN CONTENT: TABLE + CLOUD CALLOUT + UTILITIES */}
      <div className="flex flex-col lg:flex-row items-start gap-6">
        {/* LEFT SPREADSHEET TABLE */}
        <div className="bg-white border-2 border-black shadow-xl overflow-x-auto w-full lg:w-auto lg:min-w-[540px]">
          <table className="w-full border-collapse text-xs text-black">
            <thead>
              <tr>
                <th className="border border-black bg-[#fff2cc] text-black font-bold text-center px-2 py-2.5 w-11 sticky left-0 z-10">
                  NO
                </th>
                <th className="border border-black bg-[#92d050] text-black font-bold text-center px-3 py-2.5 min-w-[240px]">
                  NAMA
                </th>
                <th className="border border-black bg-[#ff66ff] text-black font-extrabold text-center px-3 py-2.5 min-w-[140px]">
                  NAIK KELAS
                </th>
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: totalRowSlots }).map((_, rowIdx) => {
                const student = displayedStudents[rowIdx];
                const rowNum = filteredStudent ? 1 : rowIdx + 1;
                const decision = student ? getStudentDecision(student) : undefined;
                const optionValue = getOptionValue(decision);

                return (
                  <tr
                    key={student ? student.idSiswa : `empty-naik-${rowIdx}`}
                    className="hover:bg-cyan-50/60"
                  >
                    <td className="border border-black text-center font-medium py-1 px-1.5 bg-white sticky left-0 z-10 h-[28px]">
                      {rowNum}
                    </td>
                    <td className="border border-black px-2.5 py-1 font-semibold text-left bg-white whitespace-nowrap">
                      {student ? student.namaSiswa : ''}
                    </td>
                    <td className="border border-black p-0 bg-white text-center">
                      {student ? (
                        <select
                          value={optionValue}
                          onChange={(e) =>
                            handleSelectChange(student, e.target.value as '' | 'Ya' | 'Tidak')
                          }
                          aria-label={`Status Naik Kelas untuk ${student.namaSiswa}`}
                          className={`w-full h-full py-1 px-2 text-center font-extrabold text-xs bg-transparent cursor-pointer focus:bg-yellow-50 focus:outline-none ${
                            optionValue === 'Ya'
                              ? 'text-emerald-800'
                              : optionValue === 'Tidak'
                              ? 'text-red-600'
                              : 'text-slate-500'
                          }`}
                        >
                          <option value="">-- Pilih --</option>
                          <option value="Ya">Ya</option>
                          <option value="Tidak">Tidak</option>
                        </select>
                      ) : (
                        <div className="h-[26px]" />
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* RIGHT CLOUD INSTRUCTION CALLOUT & TOOLS */}
        <div className="flex flex-col items-start gap-4 lg:pt-8 max-w-xs w-full">
          {/* Cloud Thought-Bubble matching screenshot */}
          <div className="relative flex items-start gap-2">
            <div className="hidden lg:flex flex-col items-center gap-1.5 mt-1 -ml-4">
              <span className="w-2.5 h-2.5 rounded-full bg-white border-2 border-[#3b82f6] block" />
              <span className="w-4 h-4 rounded-full bg-white border-2 border-[#3b82f6] block ml-3" />
            </div>
            <div className="bg-white border-2 border-[#3b82f6] rounded-[42px] px-6 py-5 shadow-md text-center text-xs font-semibold text-slate-900 leading-relaxed">
              Pilih pada list yang telah disediakan!
              <br />
              Pilih <span className="font-extrabold text-emerald-700">"Ya"</span> jika naik kelas
              <br />
              dan Pilih <span className="font-extrabold text-red-600">"Tidak"</span> Jika tidak naik kelas.
            </div>
          </div>

          {/* Filter, Summary Stats & Excel Export */}
          <div className="bg-white/95 border-2 border-black rounded-lg p-3 shadow-md w-full space-y-2.5 text-xs">
            <div className="font-bold text-slate-900 border-b border-slate-200 pb-1">
              Ringkasan Kelas {selectedKelas} (Semester 2)
            </div>
            <div className="grid grid-cols-3 gap-1.5 text-center font-bold">
              <div className="bg-emerald-50 border border-emerald-300 rounded py-1 px-1 text-emerald-800">
                <div className="text-sm">{summaryCounts.ya}</div>
                <div className="text-[10px]">Ya (Naik)</div>
              </div>
              <div className="bg-rose-50 border border-rose-300 rounded py-1 px-1 text-rose-700">
                <div className="text-sm">{summaryCounts.tidak}</div>
                <div className="text-[10px]">Tidak</div>
              </div>
              <div className="bg-slate-100 border border-slate-300 rounded py-1 px-1 text-slate-700">
                <div className="text-sm">{summaryCounts.belum}</div>
                <div className="text-[10px]">Belum</div>
              </div>
            </div>

            <div className="pt-1">
              <StudentFilterSelect />
            </div>

            <button
              onClick={handleExportExcel}
              className="w-full py-1.5 px-3 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-[11px] rounded border border-black flex items-center justify-center gap-1.5 cursor-pointer"
            >
              <FileSpreadsheet size={14} />
              <span>Ekspor Excel (.xlsx)</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
