from pathlib import Path

import pandas as pd

from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render

from rest_framework import filters, viewsets
from rest_framework.decorators import api_view
from rest_framework.response import Response

from .forms import UploadExcelForm
from .models import ImportedFile, Kelas, Siswa
from .serializers import KelasSerializer, SiswaSerializer


# ==========================================================
# HELPER
# ==========================================================

def format_nama_kelas(kelas):
    """
    Menghasilkan nama kelas yang konsisten.

    Contoh:
    tingkat = XII
    jurusan = IPA
    nama = A

    Hasil:
    XII IPA A
    """

    if not kelas:
        return "-"

    tingkat = (kelas.tingkat or "").strip()
    jurusan = (kelas.jurusan or "").strip()
    nama = (kelas.nama or "").strip()

    # Jika nama sudah lengkap
    if nama.upper().startswith(("X ", "XI ", "XII ")):
        return nama

    bagian = []

    if tingkat:
        bagian.append(tingkat)

    if jurusan:
        bagian.append(jurusan)

    if nama:
        bagian.append(nama)

    return " ".join(bagian)


def _parse_kelas(kelas_name):
    """
    Mengubah nama kelas menjadi:

    nama
    tingkat
    jurusan

    Contoh:
    XII IPA A
    ↓
    nama = A
    tingkat = XII
    jurusan = IPA
    """

    if not kelas_name:
        return None, None, None

    parts = str(kelas_name).strip().split()

    if len(parts) >= 3:
        tingkat = parts[0]
        jurusan = parts[1]
        nama = " ".join(parts[2:])

        return nama, tingkat, jurusan

    if len(parts) == 2:
        tingkat = parts[0]
        nama = parts[1]

        return nama, tingkat, None

    return parts[0], None, None


def _normalize_header(value):
    """
    Normalisasi nama kolom Excel.
    """

    if value is None:
        return ""

    value = str(value)
    value = value.replace("\u00a0", " ")
    value = " ".join(value.split())

    return value.strip().lower()


def _pick_column(df_columns, candidates):
    """
    Mencari nama kolom Excel berdasarkan beberapa alias.
    """

    normalized_columns = {
        _normalize_header(column): column
        for column in df_columns
    }

    for candidate in candidates:

        candidate_normalized = _normalize_header(candidate)

        if candidate_normalized in normalized_columns:
            return normalized_columns[candidate_normalized]

    return None


def _clean_cell(value):
    """
    Membersihkan data dari Excel.
    """

    if value is None:
        return None

    if isinstance(value, float) and pd.isna(value):
        return None

    if isinstance(value, str):

        value = " ".join(value.split()).strip()

        return value if value else None

    return value


def _normalize_jenis_kelamin(value):
    """
    Normalisasi jenis kelamin menjadi L atau P.
    """

    value = _clean_cell(value)

    if value is None:
        return None

    value = str(value).strip().lower()

    laki_laki = {
        "l",
        "la",
        "lk",
        "laki",
        "laki laki",
        "laki-laki",
        "pria",
        "male",
        "cowok",
        "1",
    }

    perempuan = {
        "p",
        "pa",
        "pr",
        "perempuan",
        "wanita",
        "female",
        "cewek",
        "2",
    }

    if value in laki_laki:
        return "L"

    if value in perempuan:
        return "P"

    return None


def _int_val(row, column):
    """
    Mengubah nilai menjadi integer.
    """

    value = _clean_cell(row.get(column)) if column else None

    if value is None:
        return 0

    try:
        return int(value)

    except (TypeError, ValueError):
        return 0


# ==========================================================
# IMPORT EXCEL
# ==========================================================

@login_required(login_url="login")
def import_excel(request):

    if request.method == "POST":

        form = UploadExcelForm(
            request.POST,
            request.FILES
        )

        if form.is_valid():

            excel_file = request.FILES["file"]

            # Validasi ekstensi
            if not excel_file.name.lower().endswith(
                (".xlsx", ".xls")
            ):

                messages.error(
                    request,
                    "File harus berupa Excel (.xlsx atau .xls)."
                )

                return redirect("import_excel")

            # Baca Excel
            df = pd.read_excel(excel_file)

            df.columns = (
                df.columns
                .astype(str)
                .str.strip()
            )

            # --------------------------------------------------
            # CARI KOLOM
            # --------------------------------------------------

            col_no = _pick_column(
                df.columns,
                ["no", "nomor"]
            )

            col_nis = _pick_column(
                df.columns,
                ["nis", "nim"]
            )

            col_nama = _pick_column(
                df.columns,
                [
                    "nama",
                    "nama siswa",
                    "nama_siswa",
                    "nama_lengkap"
                ]
            )

            col_jk = _pick_column(
                df.columns,
                [
                    "jenis kelamin",
                    "jenis_kelamin",
                    "jk",
                    "gender"
                ]
            )

            col_kelas = _pick_column(
                df.columns,
                [
                    "kelas",
                    "kelas siswa"
                ]
            )

            col_alamat = _pick_column(
                df.columns,
                [
                    "alamat",
                    "address"
                ]
            )

            col_hadir = _pick_column(
                df.columns,
                ["hadir"]
            )

            col_izin = _pick_column(
                df.columns,
                ["izin"]
            )

            col_sakit = _pick_column(
                df.columns,
                ["sakit"]
            )

            col_alfa = _pick_column(
                df.columns,
                ["alfa"]
            )

            col_tugas = _pick_column(
                df.columns,
                ["tugas"]
            )

            col_terlambat = _pick_column(
                df.columns,
                [
                    "terlambat tugas",
                    "terlambat_tugas"
                ]
            )

            col_uts = _pick_column(
                df.columns,
                ["uts"]
            )

            col_uas = _pick_column(
                df.columns,
                ["uas"]
            )

            col_nilai = _pick_column(
                df.columns,
                [
                    "nilai akhir",
                    "nilai_akhir"
                ]
            )

            col_status = _pick_column(
                df.columns,
                [
                    "status kelulusan",
                    "status_kelulusan"
                ]
            )

            # --------------------------------------------------
            # VALIDASI KOLOM WAJIB
            # --------------------------------------------------

            required_columns = []

            if not col_nis:
                required_columns.append("NIS")

            if not col_nama:
                required_columns.append("Nama Siswa")

            if not col_jk:
                required_columns.append("Jenis Kelamin")

            if not col_kelas:
                required_columns.append("Kelas")

            if required_columns:

                messages.error(
                    request,
                    "Kolom tidak ditemukan: "
                    + ", ".join(required_columns)
                )

                return redirect("import_excel")

            berhasil = 0

            # --------------------------------------------------
            # TRANSACTION
            # --------------------------------------------------

            with transaction.atomic():

                for _, row in df.iterrows():

                    # NIS
                    nis = _clean_cell(
                        row.get(col_nis)
                    )

                    if nis is None:
                        continue

                    if isinstance(nis, float):
                        nis = str(int(nis))
                    else:
                        nis = str(nis)

                    # Hindari duplikasi
                    if Siswa.objects.filter(
                        nis=nis
                    ).exists():

                        continue

                    # Nama
                    nama = _clean_cell(
                        row.get(col_nama)
                    )

                    if not nama:
                        continue

                    # Jenis kelamin
                    jenis_kelamin = _normalize_jenis_kelamin(
                        row.get(col_jk)
                    )

                    if not jenis_kelamin:
                        continue

                    # Alamat
                    alamat = (
                        _clean_cell(row.get(col_alamat))
                        if col_alamat
                        else None
                    )

                    # Nomor
                    no = _clean_cell(
                        row.get(col_no)
                    ) if col_no else None

                    if no is not None:

                        try:
                            no = int(no)

                        except (TypeError, ValueError):
                            no = None

                    # --------------------------------------------------
                    # KELAS
                    # --------------------------------------------------

                    kelas_name = _clean_cell(
                        row.get(col_kelas)
                    )

                    if not kelas_name:
                        continue

                    nama_kelas, tingkat, jurusan = _parse_kelas(
                        kelas_name
                    )

                    kelas_obj = Kelas.objects.filter(
                        nama=nama_kelas,
                        tingkat=tingkat,
                        jurusan=jurusan
                    ).first()

                    if not kelas_obj:

                        kelas_obj = Kelas.objects.create(
                            nama=nama_kelas,
                            tingkat=tingkat,
                            jurusan=jurusan
                        )

                    # --------------------------------------------------
                    # NILAI AKHIR
                    # --------------------------------------------------

                    nilai_akhir = (
                        _clean_cell(row.get(col_nilai))
                        if col_nilai
                        else None
                    )

                    if nilai_akhir is not None:

                        try:
                            nilai_akhir = float(
                                nilai_akhir
                            )

                        except (TypeError, ValueError):
                            nilai_akhir = None

                    # Status
                    status = (
                        _clean_cell(row.get(col_status))
                        if col_status
                        else None
                    )

                    # --------------------------------------------------
                    # SIMPAN SISWA
                    # --------------------------------------------------

                    Siswa.objects.create(

                        no=no,

                        nis=nis,

                        nama=nama,

                        jenis_kelamin=jenis_kelamin,

                        alamat=alamat,

                        kelas=kelas_obj,

                        hadir=_int_val(
                            row,
                            col_hadir
                        ),

                        izin=_int_val(
                            row,
                            col_izin
                        ),

                        sakit=_int_val(
                            row,
                            col_sakit
                        ),

                        alfa=_int_val(
                            row,
                            col_alfa
                        ),

                        tugas=_int_val(
                            row,
                            col_tugas
                        ),

                        terlambat_tugas=_int_val(
                            row,
                            col_terlambat
                        ),

                        uts=_int_val(
                            row,
                            col_uts
                        ),

                        uas=_int_val(
                            row,
                            col_uas
                        ),

                        nilai_akhir=nilai_akhir,

                        status_kelulusan=status,
                    )

                    berhasil += 1

            # --------------------------------------------------
            # SIMPAN FILE
            # --------------------------------------------------

            excel_file.seek(0)

            imported = ImportedFile(
                file=excel_file,
                nama_asli=excel_file.name,
                jumlah_data=berhasil,
            )

            imported.save()

            messages.success(
                request,
                f"{berhasil} data berhasil diimport."
            )

            return redirect("import_excel")

    else:

        form = UploadExcelForm()

    return render(
        request,
        "import_excel.html",
        {
            "form": form
        }
    )


# ==========================================================
# HALAMAN
# ==========================================================

@login_required(login_url="login")
def dashboard(request):

    return render(
        request,
        "dashboard.html"
    )


@login_required(login_url="login")
def data_siswa(request):

    return render(
        request,
        "data_siswa.html"
    )


@login_required(login_url="login")
def status_kelulusan(request):

    return render(
        request,
        "Status_kelulusan.html"
    )


@login_required(login_url="login")
def penyimpanan_file(request):

    files = ImportedFile.objects.all().order_by(
        "-tanggal_upload"
    )

    return render(
        request,
        "penyimpanan_file.html",
        {
            "files": files
        }
    )


@login_required(login_url="login")
def hapus_file(request, pk):

    file_obj = get_object_or_404(
        ImportedFile,
        pk=pk
    )

    if request.method == "POST":

        if file_obj.file:
            file_obj.file.delete(
                save=False
            )

        file_obj.delete()

        messages.success(
            request,
            f"File '{file_obj.nama_asli}' berhasil dihapus."
        )

    return redirect(
        "penyimpanan_file"
    )


# ==========================================================
# API DASHBOARD
# DATABASE → API
# ==========================================================

@api_view(["GET"])
def dashboard_data_api(request):

    nama_kelas = request.GET.get(
        "kelas"
    )

    # ------------------------------------------------------
    # AMBIL DATA DATABASE
    # ------------------------------------------------------

    siswa = Siswa.objects.select_related(
        "kelas"
    ).all()

    data = list(
        siswa.values(

            "id",

            "nama",

            "nis",

            "jenis_kelamin",

            "hadir",

            "izin",

            "sakit",

            "alfa",

            "tugas",

            "terlambat_tugas",

            "uts",

            "uas",

            "nilai_akhir",

            "status_kelulusan",

            "kelas_id",

            "kelas__nama",

            "kelas__tingkat",

            "kelas__jurusan",
        )
    )

    df = pd.DataFrame(data)

    # ------------------------------------------------------
    # JIKA DATA KOSONG
    # ------------------------------------------------------

    if df.empty:

        return Response(
            {
                "error": "Data siswa belum tersedia."
            },
            status=404
        )

    # ------------------------------------------------------
    # FORMAT NAMA KELAS
    # ------------------------------------------------------

    df["nama_kelas"] = (
        df["kelas__tingkat"]
        .fillna("")
        .astype(str)

        + " "

        + df["kelas__jurusan"]
        .fillna("")
        .astype(str)

        + " "

        + df["kelas__nama"]
        .fillna("")
        .astype(str)
    )

    df["nama_kelas"] = (
        df["nama_kelas"]
        .str.replace(
            r"\s+",
            " ",
            regex=True
        )
        .str.strip()
    )

    # ------------------------------------------------------
    # DAFTAR KELAS
    # ------------------------------------------------------

    daftar_kelas = sorted(
        df["nama_kelas"]
        .dropna()
        .unique()
        .tolist()
    )

    # ------------------------------------------------------
    # FILTER KELAS
    # ------------------------------------------------------

    if nama_kelas and nama_kelas != "Semua":

        df = df[
            df["nama_kelas"] == nama_kelas
        ].copy()

        if df.empty:

            return Response(
                {
                    "error": (
                        f"Kelas '{nama_kelas}' "
                        "tidak ditemukan."
                    ),

                    "daftar_kelas": daftar_kelas,
                },
                status=404
            )

    else:

        nama_kelas = "Semua Kelas"

    # ------------------------------------------------------
    # STATISTIK UTAMA
    # ------------------------------------------------------

    total = len(df)

    lulus = int(
        df["status_kelulusan"]
        .eq("Lulus")
        .sum()
    )

    tidak_lulus = int(
        df["status_kelulusan"]
        .eq("Tidak Lulus")
        .sum()
    )

    rata_rata = (
        df["nilai_akhir"]
        .mean()
        if total
        else 0
    )

    # ------------------------------------------------------
    # PERSENTASE
    # ------------------------------------------------------

    persentase_lulus = (
        lulus / total * 100
        if total
        else 0
    )

    persentase_tidak_lulus = (
        tidak_lulus / total * 100
        if total
        else 0
    )

    # ------------------------------------------------------
    # TOP 10 SISWA
    # ------------------------------------------------------

    top10 = (
        df.sort_values(
            by="nilai_akhir",
            ascending=False
        )
        .head(10)
        [
            [
                "nama",
                "nis",
                "nama_kelas",
                "nilai_akhir",
                "status_kelulusan",
            ]
        ]
        .to_dict("records")
    )

    # ------------------------------------------------------
    # GRAFIK NILAI
    # ------------------------------------------------------

    grafik_nilai = (
        df.sort_values(
            by="nilai_akhir",
            ascending=False
        )
        [
            [
                "nama",
                "nilai_akhir",
            ]
        ]
        .head(10)
        .to_dict("records")
    )

    # ------------------------------------------------------
    # KEHADIRAN
    # ------------------------------------------------------

    kehadiran = {

        "hadir": int(
            df["hadir"].sum()
        ),

        "izin": int(
            df["izin"].sum()
        ),

        "sakit": int(
            df["sakit"].sum()
        ),

        "alfa": int(
            df["alfa"].sum()
        ),
    }

    # ------------------------------------------------------
    # KETERLAMBATAN
    # ------------------------------------------------------

    keterlambatan = int(
        df["terlambat_tugas"].sum()
    )

    # ------------------------------------------------------
    # RESPONSE
    # ------------------------------------------------------

    return Response({

        "kelas_dipilih": nama_kelas,

        "total": total,

        "lulus": lulus,

        "tidak_lulus": tidak_lulus,

        "persentase_lulus": round(
            persentase_lulus,
            2
        ),

        "persentase_tidak_lulus": round(
            persentase_tidak_lulus,
            2
        ),

        "rata_rata": round(
            rata_rata,
            2
        ),

        "daftar_kelas": daftar_kelas,

        "top10": top10,

        "grafik_nilai": grafik_nilai,

        "kehadiran": kehadiran,

        "keterlambatan": keterlambatan,
    })


# ==========================================================
# API ANALISIS KELAS
# ==========================================================

@api_view(["GET"])
def analisis_kelas_api(request):

    nama_kelas = request.GET.get(
        "kelas"
    )

    if not nama_kelas:

        return Response(
            {
                "error": "Parameter kelas wajib diisi."
            },
            status=400
        )

    siswa = Siswa.objects.select_related(
        "kelas"
    ).all()

    data = list(
        siswa.values(
            "id",
            "nama",
            "nis",
            "jenis_kelamin",
            "hadir",
            "izin",
            "sakit",
            "alfa",
            "tugas",
            "terlambat_tugas",
            "uts",
            "uas",
            "nilai_akhir",
            "status_kelulusan",
            "kelas_id",
            "kelas__nama",
            "kelas__tingkat",
            "kelas__jurusan",
        )
    )

    df = pd.DataFrame(data)

    if df.empty:

        return Response(
            {
                "error": "Data siswa tidak ditemukan."
            },
            status=404
        )

    # Nama kelas
    df["nama_kelas"] = (
        df["kelas__tingkat"]
        .fillna("")
        .astype(str)

        + " "

        + df["kelas__jurusan"]
        .fillna("")
        .astype(str)

        + " "

        + df["kelas__nama"]
        .fillna("")
        .astype(str)
    )

    df["nama_kelas"] = (
        df["nama_kelas"]
        .str.replace(
            r"\s+",
            " ",
            regex=True
        )
        .str.strip()
    )

    # Filter
    df_kelas = df[
        df["nama_kelas"] == nama_kelas
    ].copy()

    if df_kelas.empty:

        return Response(
            {
                "error": (
                    f"Kelas '{nama_kelas}' "
                    "tidak ditemukan."
                )
            },
            status=404
        )

    total = len(df_kelas)

    lulus = int(
        df_kelas["status_kelulusan"]
        .eq("Lulus")
        .sum()
    )

    tidak_lulus = int(
        df_kelas["status_kelulusan"]
        .eq("Tidak Lulus")
        .sum()
    )

    rata_rata = (
        df_kelas["nilai_akhir"]
        .mean()
    )

    top10 = (
        df_kelas
        .sort_values(
            by="nilai_akhir",
            ascending=False
        )
        .head(10)
        [
            [
                "nama",
                "nis",
                "nilai_akhir",
                "status_kelulusan"
            ]
        ]
        .to_dict("records")
    )

    return Response({

        "kelas": nama_kelas,

        "total_siswa": total,

        "lulus": lulus,

        "tidak_lulus": tidak_lulus,

        "persentase_lulus": round(
            lulus / total * 100,
            2
        ),

        "persentase_tidak_lulus": round(
            tidak_lulus / total * 100,
            2
        ),

        "rata_rata_nilai": round(
            rata_rata,
            2
        ),

        "kehadiran": {

            "hadir": int(
                df_kelas["hadir"].sum()
            ),

            "izin": int(
                df_kelas["izin"].sum()
            ),

            "sakit": int(
                df_kelas["sakit"].sum()
            ),

            "alfa": int(
                df_kelas["alfa"].sum()
            ),
        },

        "keterlambatan_tugas": int(
            df_kelas[
                "terlambat_tugas"
            ].sum()
        ),

        "top10": top10,
    })


# ==========================================================
# SISWA API
# ==========================================================

class SiswaViewSet(viewsets.ModelViewSet):

    queryset = Siswa.objects.all()

    serializer_class = SiswaSerializer

    filter_backends = [
        filters.SearchFilter,
        filters.OrderingFilter,
    ]

    search_fields = [
        "nama",
        "nis",
    ]

    ordering_fields = [
        "nama",
        "nis",
        "id",
    ]

    ordering = [
        "nama"
    ]


# ==========================================================
# KELAS API
# ==========================================================

class KelasViewSet(viewsets.ModelViewSet):

    queryset = Kelas.objects.all().order_by(
        "id"
    )

    serializer_class = KelasSerializer