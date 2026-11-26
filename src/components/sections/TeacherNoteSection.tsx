import React, { useState, useMemo, useEffect } from 'react';
import {
  Home,
  HelpCircle,
  Trash2,
  Save,
  FileSpreadsheet,
  Upload,
  Sparkles,
  X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp } from '../../context/AppContext';
import { TeacherNote, Student } from '../../types';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentRecord } from '../../utils/studentSyncHelper';

const MAX_SHEET_ROWS = 60;

const NOTE_PRESETS = [
  'Menunjukkan perkembangan belajar yang sangat baik dan aktif. Terus tingkatkan kemampuan membaca dan ketelitian menulis.',
  'Rajin, sopan, dan memiliki hubungan pertemanan yang sangat positif. Tingkatkan rasa percaya diri untuk aktif berbicara di depan kelas.',
  'Memiliki potensi yang sangat baik. Diharapkan lebih fokus dan disiplin dalam mengulang pelajaran di rumah bersama orang tua.'
];

export const TeacherNoteSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    saveCatatanWali,
    bulkSaveCatatanWali,
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

  const buildNotesMap = (): Record<string, TeacherNote> => {
    const map: Record<string, TeacherNote> = {};
    siswaKelas.forEach((s) => {
      const match = db.catatanWali.find(
        (c) => matchesStudentRecord(c, s) && c.semester === selectedSemester
      );
      if (match) {
        map[s.idSiswa] = {
          ...match,
          idSiswa: s.idSiswa,
          nisn: s.nisn,
          namaSiswa: s.namaSiswa,
          kelas: selectedKelas
        };
      } else {
        map[s.idSiswa] = {
          id: `CW_${s.idSiswa}_${selectedSemester}`,
          idSiswa: s.idSiswa,
          nisn: s.nisn,
          namaSiswa: s.namaSiswa,
          kelas: selectedKelas,
          semester: selectedSemester,
          tahunPelajaran: db.sekolah.tahunPelajaran,
          catatanWaliKelas: ''
        };
      }
    });
    return map;
  };

  const [notes, setNotes] = useState<Record<string, TeacherNote>>(buildNotesMap);

  useEffect(() => {
    setNotes(buildNotesMap());
  }, [db.siswa, db.catatanWali, selectedKelas, selectedSemester]);

  const buildFallbackNote = (s: Student): TeacherNote => ({
    id: `CW_${s.idSiswa}_${selectedSemester}`,
    idSiswa: s.idSiswa,
    nisn: s.nisn,
    namaSiswa: s.namaSiswa,
    kelas: selectedKelas,
    semester: selectedSemester,
    tahunPelajaran: db.sekolah.tahunPelajaran,
    catatanWaliKelas: ''
  });

  const handleTextChange = (student: Student, text: string) => {
    const updated: TeacherNote = {
      ...(notes[student.idSiswa] || buildFallbackNote(student)),
      catatanWaliKelas: text
    };
    setNotes((prev) => ({
      ...prev,
      [student.idSiswa]: updated
    }));
    saveCatatanWali(updated, true);
  };

  const handleFillDefaultNotes = () => {
    const updatedList: TeacherNote[] = [];
    const nextNotes = { ...notes };

    displayedSiswa.forEach((s) => {
      const currentText = (nextNotes[s.idSiswa]?.catatanWaliKelas || '').trim();
      if (!currentText) {
        const noteObj: TeacherNote = {
          ...(nextNotes[s.idSiswa] || buildFallbackNote(s)),
          catatanWaliKelas:
            'Menunjukkan sikap belajar yang baik dan aktif dalam kelas. Terus pertahankan semangat belajarmu!'
        };
        nextNotes[s.idSiswa] = noteObj;
        updatedList.push(noteObj);
      }
    });

    if (updatedList.length === 0) {
      showSaveNotification('Seluruh siswa sudah memiliki Catatan Guru.');
      return;
    }

    setNotes(nextNotes);
    bulkSaveCatatanWali(updatedList);
  };

  const handleSaveAll = () => {
    const allList = siswaKelas.map((s) => notes[s.idSiswa] || buildFallbackNote(s));
    bulkSaveCatatanWali(allList);
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['CATATAN GURU (WALI KELAS)'],
      [`Kelas: ${selectedKelas} • Semester: ${selectedSemester} • Tahun Ajaran: ${db.sekolah.tahunPelajaran}`],
      [],
      ['NO', 'NISN', 'NAMA SISWA', 'KELAS', 'SEMESTER', 'CATATAN GURU'],
      ...siswaKelas.map((s, idx) => [
        idx + 1,
        s.nisn,
        s.namaSiswa,
        selectedKelas,
        selectedSemester,
        notes[s.idSiswa]?.catatanWaliKelas || ''
      ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 18 },
      { wch: 30 },
      { wch: 10 },
      { wch: 12 },
      { wch: 65 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Catatan Guru');
    XLSX.writeFile(wb, `catatan-guru-kelas-${selectedKelas}-${selectedSemester}.xlsx`);
    showSaveNotification('Catatan Guru berhasil diekspor ke Excel (.xlsx)!');
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const result = await parseAndImportTemplate(file, 'template_catatan_wali', appContext);
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
              title="Panduan Pengisian Catatan Guru"
            >
              <HelpCircle size={17} />
            </button>
            <button
              onClick={() => openClearDataModal()}
              className="w-8 h-8 rounded-full bg-red-600 hover:bg-red-500 text-white border border-red-300 flex items-center justify-center cursor-pointer"
              title="Bersihkan Semua Catatan Guru"
            >
              <Trash2 size={15} />
            </button>
          </div>

          <h1 className="text-lg sm:text-xl font-extrabold tracking-[0.2em] uppercase text-black font-serif">
            CATATAN GURU
          </h1>
        </div>

        {/* Right Action Controls */}
        <div className="flex flex-wrap items-center gap-1.5">
          <div className="bg-white rounded border border-black px-1 py-0.5">
            <StudentFilterSelect />
          </div>

          <button
            onClick={handleFillDefaultNotes}
            className="px-2.5 py-1.5 bg-yellow-300 hover:bg-yellow-200 text-slate-950 font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Isi catatan motivasi otomatis untuk siswa yang masih kosong"
          >
            <Sparkles size={13} />
            <span>Isi Otomatis</span>
          </button>

          <button
            onClick={handleSaveAll}
            className="px-2.5 py-1.5 bg-sky-700 hover:bg-sky-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Simpan & Hubungkan semua Catatan Guru ke Rapor"
          >
            <Save size={13} />
            <span>Simpan</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Ekspor Catatan Guru ke Excel (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>Excel</span>
          </button>

          <label
            className="px-2.5 py-1.5 bg-indigo-700 hover:bg-indigo-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Impor Catatan Guru dari Excel"
          >
            <Upload size={13} />
            <span>Impor</span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} className="hidden" />
          </label>
        </div>
      </div>

      {/* HELP PANEL */}
      {showHelp && (
        <div className="mb-2 bg-white border-2 border-black p-3 rounded shadow-lg text-xs text-slate-800">
          <div className="flex items-center justify-between border-b border-slate-200 pb-1 mb-1.5">
            <h3 className="font-bold text-slate-900">
              Panduan Pengisian Catatan Guru / Wali Kelas (Kelas {selectedKelas} • Semester {selectedSemester})
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
              <strong>Pengisian Langsung:</strong> Ketik langsung catatan perkembangan sikap, apresiasi, atau motivasi siswa pada kolom <em>Catatan Guru</em> (tersedia saran template otomatis saat mengetik).
            </li>
            <li>
              <strong>Terhubung ke Rapor:</strong> Semua catatan langsung tampil pada kolom <em>Catatan Wali Kelas</em> di Rapor Akademik, Rapor PTS, Buku Induk, dan Cetak Rapor.
            </li>
            <li>
              <strong>Isi Otomatis &amp; Simpan:</strong> Gunakan tombol <strong>Isi Otomatis</strong> untuk mengisi siswa yang masih kosong, lalu klik <strong>Simpan</strong> untuk menyimpan riwayat data.
            </li>
          </ul>
        </div>
      )}

      {/* PRESET SUGGESTIONS */}
      <datalist id="catatan-guru-presets">
        {NOTE_PRESETS.map((preset) => (
          <option key={preset} value={preset} />
        ))}
      </datalist>

      {/* SPREADSHEET TABLE MATCHING SCREENSHOT */}
      <div className="bg-white border-2 border-black shadow-xl overflow-x-auto max-w-5xl">
        <table className="w-full border-collapse text-xs text-black">
          <thead>
            <tr>
              <th className="border border-black bg-[#00ffff] text-black font-bold text-center px-2 py-2.5 w-12 font-serif">
                No.
              </th>
              <th className="border border-black bg-[#00ff00] text-black font-bold text-center px-3 py-2.5 min-w-[210px] w-[260px] font-serif">
                Nama
              </th>
              <th className="border border-black bg-[#00b0f0] text-black font-bold text-center px-3 py-2.5 min-w-[440px] font-serif">
                Catatan Guru
              </th>
            </tr>
          </thead>

          <tbody>
            {Array.from({ length: totalRowSlots }).map((_, rowIdx) => {
              const student = displayedSiswa[rowIdx];
              const noteText = student ? notes[student.idSiswa]?.catatanWaliKelas || '' : '';
              const rowNum = filteredStudent ? 1 : rowIdx + 1;

              return (
                <tr
                  key={student ? student.idSiswa : `empty-note-row-${rowIdx}`}
                  className="hover:bg-sky-50/40"
                >
                  <td className="border border-black bg-white text-center font-medium py-1 px-2 h-[27px]">
                    {rowNum}
                  </td>
                  <td className="border border-black bg-white px-2.5 py-1 font-semibold text-left whitespace-nowrap">
                    {student ? student.namaSiswa : ''}
                  </td>
                  <td className="border border-black bg-white p-0">
                    {student ? (
                      <input
                        type="text"
                        list="catatan-guru-presets"
                        value={noteText}
                        onChange={(e) => handleTextChange(student, e.target.value)}
                        placeholder="Tuliskan catatan guru / wali kelas..."
                        className="w-full h-full px-2.5 py-1 text-xs text-slate-900 bg-transparent focus:bg-yellow-50 focus:outline-none"
                      />
                    ) : (
                      <div className="h-[25px]" />
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
