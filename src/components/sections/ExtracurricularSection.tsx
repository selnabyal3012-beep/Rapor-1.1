import React, { useMemo, useState } from 'react';
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
import { Extracurricular, Student } from '../../types';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentRecord } from '../../utils/studentSyncHelper';

type SlotNum = 1 | 2 | 3 | 4 | 5;

const SLOT_DEFS: Array<{
  slot: SlotNum;
  groupTitle: string;
  groupBg: string;
  nameLabel: string;
  nameBg: string;
  nameText: string;
  descBg: string;
}> = [
  {
    slot: 1,
    groupTitle: 'EKSTRAKURIKULER WAJIB',
    groupBg: '#ffff00',
    nameLabel: 'EKSTRAKURIKULER YANG DIIKUTI',
    nameBg: '#00ff00',
    nameText: '#000000',
    descBg: '#ff9999'
  },
  {
    slot: 2,
    groupTitle: 'EKSTRAKURIKULER PILIHAN 1',
    groupBg: '#9bc2e6',
    nameLabel: 'Ekskul 2',
    nameBg: '#00b0f0',
    nameText: '#000000',
    descBg: '#00b050'
  },
  {
    slot: 3,
    groupTitle: 'EKSTRAKURIKULER PILIHAN 2',
    groupBg: '#ff7c80',
    nameLabel: 'Ekskul 3',
    nameBg: '#ffc000',
    nameText: '#000000',
    descBg: '#b4c6e7'
  },
  {
    slot: 4,
    groupTitle: 'EKSTRAKURIKULER PILIHAN 3',
    groupBg: '#b4a7d6',
    nameLabel: 'Ekskul 4',
    nameBg: '#ff0000',
    nameText: '#ffffff',
    descBg: '#b4c6e7'
  },
  {
    slot: 5,
    groupTitle: 'EKSTRAKURIKULER PILIHAN 4',
    groupBg: '#00ffff',
    nameLabel: 'Ekskul 5',
    nameBg: '#ff0000',
    nameText: '#ffffff',
    descBg: '#b4c6e7'
  }
];

const EKSKUL_SUGGESTIONS = [
  'Pramuka',
  'Pramuka (Wajib)',
  'Olahraga / Sepak Bola',
  'Tari Daerah',
  'Noken Kriya',
  'Musik Tradisional',
  'Paduan Suara',
  'Dokter Kecil / UKS',
  'Kerohanian',
  'Literasi & Membaca',
  'Seni Lukis'
];

const MAX_SHEET_ROWS = 60;

export const ExtracurricularSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    saveEkskul,
    bulkSaveEkskul,
    deleteEkskul,
    showSaveNotification,
    setActiveMenu,
    openClearDataModal
  } = appContext;

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

  // Map each student's existing Extracurricular rows into slots 1..5
  const studentSlotsMap = useMemo(() => {
    const map = new Map<string, Partial<Record<SlotNum, Extracurricular>>>();

    siswaKelas.forEach((student) => {
      const studentRecords = db.ekstrakurikuler.filter(
        (e) =>
          matchesStudentRecord(e, student) &&
          e.kelas === selectedKelas &&
          e.semester === selectedSemester
      );

      const slots: Partial<Record<SlotNum, Extracurricular>> = {};
      const unslotted: Extracurricular[] = [];

      studentRecords.forEach((rec) => {
        const inferredSlot =
          rec.slotIndex ||
          (rec.id.includes('_slot1_')
            ? 1
            : rec.id.includes('_slot2_')
            ? 2
            : rec.id.includes('_slot3_')
            ? 3
            : rec.id.includes('_slot4_')
            ? 4
            : rec.id.includes('_slot5_')
            ? 5
            : undefined);
        if (inferredSlot && !slots[inferredSlot]) {
          slots[inferredSlot] = rec;
        } else {
          unslotted.push(rec);
        }
      });

      const slotOrder: SlotNum[] = [1, 2, 3, 4, 5];
      unslotted.forEach((rec) => {
        const freeSlot = slotOrder.find((s) => !slots[s]);
        if (freeSlot) {
          slots[freeSlot] = rec;
        }
      });

      map.set(student.idSiswa, slots);
    });

    return map;
  }, [siswaKelas, db.ekstrakurikuler, selectedKelas, selectedSemester]);

  const handleCellChange = (
    student: Student,
    slot: SlotNum,
    field: 'jenisEkstrakurikuler' | 'keterangan',
    value: string
  ) => {
    const existing = studentSlotsMap.get(student.idSiswa)?.[slot];
    const nextName = field === 'jenisEkstrakurikuler' ? value : existing?.jenisEkstrakurikuler || '';
    const nextDesc = field === 'keterangan' ? value : existing?.keterangan || '';

    if (!nextName.trim() && !nextDesc.trim()) {
      if (existing?.id) {
        deleteEkskul(existing.id);
      }
      return;
    }

    const recordId = existing?.id || `EKS_${student.idSiswa}_slot${slot}_${selectedSemester}`;
    const updated: Extracurricular = {
      id: recordId,
      idSiswa: student.idSiswa,
      nisn: student.nisn,
      namaSiswa: student.namaSiswa,
      kelas: selectedKelas,
      semester: selectedSemester,
      tahunPelajaran: db.sekolah.tahunPelajaran,
      slotIndex: slot,
      jenisEkstrakurikuler: nextName,
      predikat: existing?.predikat || 'Baik',
      keterangan: nextDesc
    };

    saveEkskul(updated, true);
  };

  const handleFillDefaultWajib = () => {
    const toSave: Extracurricular[] = [];
    displayedStudents.forEach((student) => {
      const slot1 = studentSlotsMap.get(student.idSiswa)?.[1];
      if (!slot1 || !slot1.jenisEkstrakurikuler.trim()) {
        toSave.push({
          id: slot1?.id || `EKS_${student.idSiswa}_slot1_${selectedSemester}`,
          idSiswa: student.idSiswa,
          nisn: student.nisn,
          namaSiswa: student.namaSiswa,
          kelas: selectedKelas,
          semester: selectedSemester,
          tahunPelajaran: db.sekolah.tahunPelajaran,
          slotIndex: 1,
          jenisEkstrakurikuler: 'Pramuka (Wajib)',
          predikat: 'Baik',
          keterangan: 'Aktif dan disiplin mengikuti kegiatan kepramukaan mingguan.'
        });
      }
    });
    if (toSave.length === 0) {
      showSaveNotification('Seluruh siswa sudah memiliki data Ekstrakurikuler Wajib.');
      return;
    }
    bulkSaveEkskul(toSave);
  };

  const handleSaveAll = () => {
    const allCurrent: Extracurricular[] = [];
    siswaKelas.forEach((student) => {
      const slots = studentSlotsMap.get(student.idSiswa);
      if (!slots) return;
      ([1, 2, 3, 4, 5] as SlotNum[]).forEach((slot) => {
        const rec = slots[slot];
        if (rec && (rec.jenisEkstrakurikuler.trim() || rec.keterangan.trim())) {
          allCurrent.push({ ...rec, slotIndex: slot });
        }
      });
    });
    if (allCurrent.length > 0) {
      bulkSaveEkskul(allCurrent);
    } else {
      showSaveNotification('Data Ekstrakurikuler berhasil disimpan.');
    }
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['DATA KEGIATAN EKSTRAKURIKULER'],
      [`Kelas: ${selectedKelas} • Semester: ${selectedSemester} • Tahun Ajaran: ${db.sekolah.tahunPelajaran}`],
      [],
      [
        'NO',
        'NAMA SISWA',
        'EKSTRAKURIKULER WAJIB',
        'DESKRIPSI WAJIB',
        'EKSTRAKURIKULER PILIHAN 1',
        'DESKRIPSI PILIHAN 1',
        'EKSTRAKURIKULER PILIHAN 2',
        'DESKRIPSI PILIHAN 2',
        'EKSTRAKURIKULER PILIHAN 3',
        'DESKRIPSI PILIHAN 3',
        'EKSTRAKURIKULER PILIHAN 4',
        'DESKRIPSI PILIHAN 4'
      ],
      ...siswaKelas.map((s, idx) => {
        const slots = studentSlotsMap.get(s.idSiswa) || {};
        return [
          idx + 1,
          s.namaSiswa,
          slots[1]?.jenisEkstrakurikuler || '',
          slots[1]?.keterangan || '',
          slots[2]?.jenisEkstrakurikuler || '',
          slots[2]?.keterangan || '',
          slots[3]?.jenisEkstrakurikuler || '',
          slots[3]?.keterangan || '',
          slots[4]?.jenisEkstrakurikuler || '',
          slots[4]?.keterangan || '',
          slots[5]?.jenisEkstrakurikuler || '',
          slots[5]?.keterangan || ''
        ];
      })
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 5 },
      { wch: 28 },
      { wch: 22 },
      { wch: 32 },
      { wch: 20 },
      { wch: 30 },
      { wch: 20 },
      { wch: 30 },
      { wch: 20 },
      { wch: 30 },
      { wch: 20 },
      { wch: 30 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Ekstrakurikuler');
    XLSX.writeFile(wb, `ekstrakurikuler-kelas-${selectedKelas}-${selectedSemester}.xlsx`);
    showSaveNotification('Data Ekstrakurikuler berhasil diekspor ke Excel (.xlsx)!');
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const result = await parseAndImportTemplate(file, 'template_ekskul', appContext);
      showSaveNotification(
        result.success ? result.message : `Gagal impor: ${result.message}`
      );
    } catch (err: any) {
      showSaveNotification(`Gagal mengimpor file: ${err?.message || 'Format tidak sesuai'}`);
    }
    e.target.value = '';
  };

  return (
    <div className="bg-[#00ff00] min-h-screen p-2 sm:p-3 rounded-xl shadow-inner">
      {/* TOP TOOLBAR STRIP */}
      <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
        <div className="flex flex-wrap items-center gap-4">
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
              title="Panduan Pengisian Ekstrakurikuler"
            >
              <HelpCircle size={17} />
            </button>
            <button
              onClick={() => openClearDataModal()}
              className="w-8 h-8 rounded-full bg-red-600 hover:bg-red-500 text-white border border-red-300 flex items-center justify-center cursor-pointer"
              title="Bersihkan Data Ekstrakurikuler"
            >
              <Trash2 size={15} />
            </button>
          </div>

          <h1 className="text-lg sm:text-xl font-extrabold tracking-wide uppercase text-black font-serif">
            EKSTRAKURIKULER
          </h1>
        </div>

        {/* Right Action Controls */}
        <div className="flex flex-wrap items-center gap-1.5">
          <div className="bg-white rounded border border-black px-1 py-0.5">
            <StudentFilterSelect />
          </div>

          <button
            onClick={handleFillDefaultWajib}
            className="px-2.5 py-1.5 bg-yellow-300 hover:bg-yellow-200 text-slate-950 font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Isi otomatis Ekskul Wajib (Pramuka) untuk siswa yang masih kosong"
          >
            <Sparkles size={13} />
            <span>Isi Pramuka</span>
          </button>

          <button
            onClick={handleSaveAll}
            className="px-2.5 py-1.5 bg-sky-600 hover:bg-sky-500 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Simpan & Hubungkan semua data Ekstrakurikuler ke Rapor"
          >
            <Save size={13} />
            <span>Simpan</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Ekspor Tabel Ekstrakurikuler ke Excel (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>Excel</span>
          </button>

          <label
            className="px-2.5 py-1.5 bg-indigo-700 hover:bg-indigo-600 text-white font-bold text-[11px] border border-black rounded flex items-center gap-1 cursor-pointer shadow-xs"
            title="Impor Data Ekstrakurikuler dari Excel"
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
              Panduan Pengisian Kegiatan Ekstrakurikuler (Kelas {selectedKelas} • Semester {selectedSemester})
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
              <strong>5 Slot Ekstrakurikuler:</strong> Tersedia 1 kolom <em>Ekstrakurikuler Wajib</em> (kuning) dan 4 kolom <em>Ekstrakurikuler Pilihan (1–4)</em>. Ketik langsung nama kegiatan dan deskripsinya pada sel tabel.
            </li>
            <li>
              <strong>Tersambung Otomatis ke Rapor:</strong> Seluruh isian yang memiliki nama atau deskripsi otomatis muncul pada <em>Rapor Akademik</em>, <em>Rapor PTS</em>, <em>Buku Induk</em>, dan <em>Cetak Rapor</em> siswa.
            </li>
            <li>
              <strong>Simpan Otomatis:</strong> Setiap ketikan tersimpan langsung; klik tombol <strong>Simpan</strong> kapan saja untuk membuat titik simpan riwayat (Undo/Redo) dan sinkronisasi penuh.
            </li>
          </ul>
        </div>
      )}

      {/* SUGGESTION DATALIST */}
      <datalist id="ekskul-suggestions">
        {EKSKUL_SUGGESTIONS.map((item) => (
          <option key={item} value={item} />
        ))}
      </datalist>

      {/* SPREADSHEET TABLE */}
      <div className="bg-white border-2 border-black shadow-xl overflow-x-auto">
        <table className="w-full border-collapse text-[11px] text-black">
          <thead>
            {/* Row 1 */}
            <tr>
              <th
                rowSpan={3}
                className="border border-black bg-[#8db4e2] text-black font-bold italic text-center px-1.5 py-1 w-9 sticky left-0 z-20"
              >
                No.
              </th>
              <th
                rowSpan={3}
                className="border border-black bg-[#8db4e2] text-black font-bold italic text-center px-2 py-1 min-w-[175px] sticky left-9 z-20"
              >
                Nama
              </th>
              <th
                colSpan={10}
                className="border border-black bg-[#00ffff] text-black font-extrabold uppercase tracking-wider text-center py-1 text-xs"
              >
                EKSTRAKURIKULER
              </th>
            </tr>

            {/* Row 2 */}
            <tr>
              {SLOT_DEFS.map((def) => (
                <th
                  key={`grp-${def.slot}`}
                  colSpan={2}
                  style={{ backgroundColor: def.groupBg }}
                  className="border border-black text-black font-extrabold uppercase text-center py-1 px-2 text-[10px] whitespace-nowrap"
                >
                  {def.groupTitle}
                </th>
              ))}
            </tr>

            {/* Row 3 */}
            <tr>
              {SLOT_DEFS.map((def) => (
                <React.Fragment key={`sub-${def.slot}`}>
                  <th
                    style={{ backgroundColor: def.nameBg, color: def.nameText }}
                    className={`border border-black font-extrabold text-center px-1.5 py-1 min-w-[115px] leading-tight ${
                      def.slot === 1 ? 'text-[8px] uppercase' : 'text-[10px]'
                    }`}
                  >
                    {def.nameLabel}
                  </th>
                  <th
                    style={{ backgroundColor: def.descBg }}
                    className="border border-black text-black font-bold text-center px-2 py-1 min-w-[190px] text-[11px]"
                  >
                    Deskripsi
                  </th>
                </React.Fragment>
              ))}
            </tr>
          </thead>

          <tbody>
            {Array.from({ length: totalRowSlots }).map((_, rowIdx) => {
              const student = displayedStudents[rowIdx];
              const slots = student ? studentSlotsMap.get(student.idSiswa) || {} : {};
              const rowNum = filteredStudent ? 1 : rowIdx + 1;

              return (
                <tr
                  key={student ? student.idSiswa : `empty-ekskul-row-${rowIdx}`}
                  className="hover:bg-sky-50/40"
                >
                  <td className="border border-black bg-[#d9d9d9] text-center font-medium py-1 px-1 sticky left-0 z-10 h-[26px]">
                    {rowNum}
                  </td>
                  <td className="border border-black bg-[#d9d9d9] px-2 py-1 font-semibold text-left sticky left-9 z-10 whitespace-nowrap">
                    {student ? student.namaSiswa : ''}
                  </td>

                  {SLOT_DEFS.map((def) => {
                    const entry = slots[def.slot];
                    return (
                      <React.Fragment key={`cell-${rowIdx}-${def.slot}`}>
                        <td className="border border-black p-0 bg-white">
                          {student ? (
                            <input
                              type="text"
                              list="ekskul-suggestions"
                              value={entry?.jenisEkstrakurikuler || ''}
                              onChange={(e) =>
                                handleCellChange(
                                  student,
                                  def.slot,
                                  'jenisEkstrakurikuler',
                                  e.target.value
                                )
                              }
                              placeholder={def.slot === 1 ? 'Pramuka...' : ''}
                              className="w-full h-full px-1.5 py-1 text-[11px] font-semibold text-slate-900 bg-transparent focus:bg-yellow-50 focus:outline-none"
                            />
                          ) : (
                            <div className="h-[24px]" />
                          )}
                        </td>
                        <td className="border border-black p-0 bg-white">
                          {student ? (
                            <input
                              type="text"
                              value={entry?.keterangan || ''}
                              onChange={(e) =>
                                handleCellChange(student, def.slot, 'keterangan', e.target.value)
                              }
                              placeholder={
                                entry?.jenisEkstrakurikuler ? 'Tulis deskripsi kegiatan...' : ''
                              }
                              className="w-full h-full px-1.5 py-1 text-[11px] text-slate-800 bg-transparent focus:bg-yellow-50 focus:outline-none"
                            />
                          ) : (
                            <div className="h-[24px]" />
                          )}
                        </td>
                      </React.Fragment>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
};
