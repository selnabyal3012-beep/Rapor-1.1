import React from 'react';
import {
  ClassLevel,
  Extracurricular,
  SchoolIdentity,
  SemesterType,
  Student,
  Subject
} from '../../types';
import { getNamaPelajaranAgama } from '../../utils/tpDescriptionHelper';

export interface AcademicScoreRowItem {
  mapel: Subject;
  score: {
    kkm?: number;
    formatif?: number;
    nilaiPts?: number;
    nilaiAkhir: number;
    predikat?: string;
    capaianKompetensi?: string;
    deskripsiRapor: string;
    deskripsiTercapai: string;
    deskripsiPerluPeningkatan: string;
  };
  linkedTpKodes?: string[];
}

export function splitAcademicScoresForPages(studentScores: AcademicScoreRowItem[]) {
  const nonMulok = studentScores.filter((item) => item.mapel.kelompok !== 'Muatan Lokal');
  const mulokScores = studentScores.filter((item) => item.mapel.kelompok === 'Muatan Lokal');

  const isSeniItem = (item: AcademicScoreRowItem) => {
    const name = item.mapel.mataPelajaran.toLowerCase();
    return item.mapel.kelompok === 'Pilihan' || name.startsWith('seni');
  };

  // Pisahkan bagian Page 4 (maksimal 7 mapel non-mulok termasuk maksimal 2 Seni Pilihan)
  // dan bagian Page 5 (sisa mapel non-mulok + semua Muatan Lokal)
  const wajibBeforeSeni: AcademicScoreRowItem[] = [];
  const seniScores: AcademicScoreRowItem[] = [];
  const wajibAfterSeni: AcademicScoreRowItem[] = [];

  let passedSeni = false;
  for (const item of nonMulok) {
    if (isSeniItem(item)) {
      passedSeni = true;
      seniScores.push(item);
    } else if (!passedSeni) {
      wajibBeforeSeni.push(item);
    } else {
      wajibAfterSeni.push(item);
    }
  }

  const page4SeniCount = Math.min(2, seniScores.length);
  const part1Main = wajibBeforeSeni;
  const part1Seni = seniScores.slice(0, page4SeniCount);
  const part2Seni = seniScores.slice(page4SeniCount);
  const part2WajibAfter = wajibAfterSeni;

  const hasPart2 = part2Seni.length > 0 || part2WajibAfter.length > 0 || mulokScores.length > 0;

  return {
    part1Main,
    part1Seni,
    part2Seni,
    part2WajibAfter,
    mulokScores,
    startNoPart2: part1Main.length + 1,
    seniStartNoPart2: page4SeniCount + 1,
    hasPart2
  };
}

interface ReportAcademicHeaderAndScoresProps {
  student: Student;
  sekolah: SchoolIdentity;
  selectedKelas: ClassLevel;
  selectedFase: string;
  semester: SemesterType;
  studentScores: AcademicScoreRowItem[];
  pagePart?: 'all' | 'part1' | 'part2';
}

export const ReportAcademicHeaderAndScores: React.FC<ReportAcademicHeaderAndScoresProps> = ({
  student,
  sekolah,
  selectedKelas,
  selectedFase,
  semester,
  studentScores,
  pagePart = 'all'
}) => {
  const isSem2 = semester === 'Genap' || semester === '2';
  const cleanNisn =
    student.nisn && !student.nisn.startsWith('BLM_') && student.nisn !== '-'
      ? student.nisn
      : '-';
  const cleanNis = student.nis && student.nis !== '-' ? student.nis : '';
  const nisDisplay = cleanNis ? `${cleanNisn} / ${cleanNis}` : cleanNisn;

  const {
    part1Main,
    part1Seni,
    part2Seni,
    part2WajibAfter,
    mulokScores,
    startNoPart2,
    seniStartNoPart2
  } = splitAcademicScoresForPages(studentScores);

  const renderSubjectRows = (item: AcademicScoreRowItem, displayNo: number | string) => {
    const mapelTitle = getNamaPelajaranAgama(item.mapel.mataPelajaran, student.agama);
    const hasScore = item.score.nilaiAkhir > 0;

    let topDesc = item.score.deskripsiTercapai?.trim() || '';
    let bottomDesc = item.score.deskripsiPerluPeningkatan?.trim() || '';

    if (!topDesc && item.score.deskripsiRapor && item.score.deskripsiRapor.trim() !== '-') {
      const parts = item.score.deskripsiRapor.split(/\n+|\.\s+(?=Perlu)/i);
      topDesc = parts[0]?.trim() || '';
      if (!bottomDesc && parts.length > 1) {
        bottomDesc = parts.slice(1).join(' ').trim();
      }
    }

    return (
      <React.Fragment key={item.mapel.kode || item.mapel.mataPelajaran}>
        {/* Sub-baris 1: Capaian yang Sudah Dikuasai */}
        <tr className="bg-white">
          <td
            rowSpan={2}
            className="border border-black py-1.5 px-1.5 text-center align-middle font-normal text-black w-[2.6em]"
          >
            {displayNo}
          </td>
          <td
            rowSpan={2}
            className="border border-black py-1.5 px-2.5 text-left align-middle font-normal text-black w-[28%]"
          >
            {mapelTitle}
          </td>
          <td
            rowSpan={2}
            className="border border-black py-1.5 px-2 text-center align-middle font-bold text-black w-[5em]"
          >
            {hasScore ? item.score.nilaiAkhir : ''}
          </td>
          <td className="report-desc-cell border border-black py-1.5 px-2.5 align-top leading-snug text-black min-h-[26px]">
            {topDesc || <span className="inline-block h-4" />}
          </td>
        </tr>
        {/* Sub-baris 2: Kompetensi yang Perlu Peningkatan */}
        <tr className="bg-white">
          <td className="report-desc-cell border border-black py-1.5 px-2.5 align-top leading-snug text-black min-h-[26px]">
            {bottomDesc || <span className="inline-block h-4" />}
          </td>
        </tr>
      </React.Fragment>
    );
  };

  const renderTableHeader = () => (
    <thead>
      <tr className="bg-white text-black font-bold">
        <th className="border border-black py-1.5 px-1.5 text-center w-[2.6em]">No</th>
        <th className="border border-black py-1.5 px-2.5 text-center w-[28%]">Mata Pelajaran</th>
        <th className="border border-black py-1.5 px-1.5 text-center w-[5em]">
          Nilai
          <br />
          Akhir
        </th>
        <th className="border border-black py-1.5 px-2.5 text-center">Capaian Kompetensi</th>
      </tr>
    </thead>
  );

  return (
    <div className="text-black">
      {/* JUDUL & IDENTITAS HANYA DITAMPILKAN DI HALAMAN 4 (PART 1 / ALL) */}
      {pagePart !== 'part2' && (
        <>
          <div className="text-center mb-4">
            <h2 className="text-sm sm:text-base font-bold tracking-wide uppercase text-black">
              LAPORAN HASIL BELAJAR
            </h2>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-1 text-xs mb-4">
            <div className="space-y-1">
              <div className="flex">
                <span className="w-[10.5em] shrink-0">Nama Peserta Didik</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-bold uppercase text-black flex-1">{student.namaSiswa}</span>
              </div>
              <div className="flex">
                <span className="w-[10.5em] shrink-0">NISN/NIS</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">{nisDisplay}</span>
              </div>
              <div className="flex">
                <span className="w-[10.5em] shrink-0">Sekolah</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">{sekolah.namaSekolah || '-'}</span>
              </div>
              <div className="flex">
                <span className="w-[10.5em] shrink-0">Alamat</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">
                  {[sekolah.alamat, sekolah.distrikKecamatan, sekolah.kabupaten]
                    .filter((v) => v && v.trim() !== '' && v !== '-')
                    .join(', ') || '-'}
                </span>
              </div>
            </div>

            <div className="space-y-1">
              <div className="flex">
                <span className="w-[9.2em] shrink-0">Kelas</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">
                  {selectedKelas}
                  {sekolah.rombel && sekolah.kelas === selectedKelas ? ` (${sekolah.rombel})` : ''}
                </span>
              </div>
              <div className="flex">
                <span className="w-[9.2em] shrink-0">Fase</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">{selectedFase}</span>
              </div>
              <div className="flex">
                <span className="w-[9.2em] shrink-0">Semester</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">
                  {!isSem2 ? '1 (Ganjil)' : '2 (Genap)'}
                </span>
              </div>
              <div className="flex">
                <span className="w-[9.2em] shrink-0">Tahun Pelajaran</span>
                <span className="w-[1.1em] text-center shrink-0">:</span>
                <span className="font-normal text-black flex-1">{sekolah.tahunPelajaran || '-'}</span>
              </div>
            </div>
          </div>
        </>
      )}

      {/* TABEL NILAI 4 KOLOM KURIKULUM MERDEKA */}
      <table className="w-full text-[11px] border-collapse border border-black bg-white">
        {renderTableHeader()}
        <tbody>
          {studentScores.length === 0 ? (
            <tr>
              <td colSpan={4} className="border border-black py-6 text-center text-slate-500 italic">
                Tidak ada mata pelajaran yang aktif pada filter ini.
              </td>
            </tr>
          ) : (
            <>
              {/* HALAMAN 4 (PART 1 atau ALL) */}
              {(pagePart === 'all' || pagePart === 'part1') && (
                <>
                  {part1Main.map((item, idx) => renderSubjectRows(item, idx + 1))}

                  {part1Seni.length > 0 && (
                    <>
                      <tr className="bg-white">
                        <td className="border border-black py-1 px-1.5 text-center font-bold">
                          {part1Main.length + 1}
                        </td>
                        <td
                          colSpan={3}
                          className="border border-black py-1 px-2.5 text-left font-bold text-black"
                        >
                          Seni Pilihan :
                        </td>
                      </tr>
                      {part1Seni.map((item, idx) => renderSubjectRows(item, idx + 1))}
                    </>
                  )}
                </>
              )}

              {/* HALAMAN 5 (PART 2 atau ALL) */}
              {(pagePart === 'all' || pagePart === 'part2') && (
                <>
                  {part2Seni.map((item, idx) =>
                    renderSubjectRows(
                      item,
                      pagePart === 'part2' ? seniStartNoPart2 + idx : part1Seni.length + idx + 1
                    )
                  )}

                  {part2WajibAfter.map((item, idx) =>
                    renderSubjectRows(item, startNoPart2 + (part1Seni.length + part2Seni.length > 0 ? 1 : 0) + idx)
                  )}

                  {mulokScores.length > 0 && (
                    <>
                      <tr className="bg-white">
                        <td className="border border-black py-1 px-1.5 text-center font-bold">
                          {startNoPart2 +
                            (part1Seni.length + part2Seni.length > 0 ? 1 : 0) +
                            part2WajibAfter.length}
                        </td>
                        <td
                          colSpan={3}
                          className="border border-black py-1 px-2.5 text-left font-bold text-black"
                        >
                          Muatan Lokal
                        </td>
                      </tr>
                      {mulokScores.map((item, idx) => renderSubjectRows(item, idx + 1))}
                    </>
                  )}
                </>
              )}
            </>
          )}
        </tbody>
      </table>
    </div>
  );
};

interface ReportExtracurricularTableProps {
  studentEkskul: Extracurricular[];
  minRows?: number;
}

export const ReportExtracurricularTable: React.FC<ReportExtracurricularTableProps> = ({
  studentEkskul,
  minRows = 5
}) => {
  const rowCount = Math.max(minRows, studentEkskul.length);
  const rows = Array.from({ length: rowCount }, (_, i) => studentEkskul[i] || null);

  return (
    <div className="mb-4 text-black">
      <table className="w-full text-[11px] border-collapse border border-black bg-white">
        <thead>
          <tr className="bg-white text-black font-bold">
            <th className="border border-black py-1.5 px-2 text-center w-[2.8em]">No</th>
            <th className="border border-black py-1.5 px-2.5 text-center w-[34%]">
              Ekstrakurikuler
            </th>
            <th className="border border-black py-1.5 px-2.5 text-center">Keterangan</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((item, idx) => (
            <tr key={item ? item.id : `empty-ekskul-${idx}`}>
              <td className="border border-black py-1.5 px-2 text-center">{idx + 1}</td>
              <td className="border border-black py-1.5 px-2.5 font-normal">
                {item ? item.jenisEkstrakurikuler : <span className="inline-block h-3.5" />}
              </td>
              <td className="border border-black py-1.5 px-2.5 font-normal">
                {item ? (
                  <span>
                    {item.predikat
                      ? `${item.predikat}${item.keterangan ? ` — ${item.keterangan}` : ''}`
                      : item.keterangan || '-'}
                  </span>
                ) : (
                  <span className="inline-block h-3.5" />
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
