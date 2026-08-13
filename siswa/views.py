from pathlib import Path
import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render

from rest_framework import filters, viewsets
from rest_framework.decorators import api_view
from rest_framework.response import Response

from .forms import UploadExcelForm
from .models import ImportedFile, Kelas, Siswa
from .serializers import KelasSerializer, SiswaSerializer
from django.db.models import Avg, Count, Max, Min, Q, Sum



def dashboard_data_api(request):
    """
    API Dashboard Data Siswa

    GET:
        /api/dashboard-data/
        /api/dashboard-data/?kelas=XII%20IPA%20A
    """

    kelas = request.GET.get("kelas", "").strip()

    # ==============================
    # QUERY SISWA
    # ==============================

    siswa_qs = Siswa.objects.select_related("kelas").all()

    # Filter kelas jika dipilih
    if kelas and kelas != "Semua Kelas":
        nama, tingkat, jurusan = _parse_kelas(kelas)
        siswa_qs = siswa_qs.filter(
            kelas__nama=nama,
            kelas__tingkat=tingkat,
            kelas__jurusan=jurusan,
        )

    # ==============================
    # DATA DASAR
    # ==============================

    total = siswa_qs.count()

    lulus = siswa_qs.filter(
        status_kelulusan__iexact="Lulus"
    ).count()

    tidak_lulus = siswa_qs.filter(
        status_kelulusan__iexact="Tidak Lulus"
    ).count()

    # ==============================
    # NILAI
    # ==============================

    nilai_qs = siswa_qs.exclude(
        nilai_akhir__isnull=True
    )

    rata_rata = nilai_qs.aggregate(
        rata=Avg("nilai_akhir")
    )["rata"]

    nilai_tertinggi = nilai_qs.aggregate(
        nilai=Max("nilai_akhir")
    )["nilai"]

    nilai_terendah = nilai_qs.aggregate(
        nilai=Min("nilai_akhir")
    )["nilai"]

    # ==============================
    # PERSENTASE KELULUSAN
    # ==============================

    if total > 0:
        persentase_lulus = round(
            (lulus / total) * 100,
            2
        )

        persentase_tidak_lulus = round(
            (tidak_lulus / total) * 100,
            2
        )
    else:
        persentase_lulus = 0
        persentase_tidak_lulus = 0

    # ==============================
    # KEHADIRAN
    # ==============================

    kehadiran = siswa_qs.aggregate(
        hadir=Avg("hadir"),
        izin=Avg("izin"),
        sakit=Avg("sakit"),
        alfa=Avg("alfa"),
    )

    total_hadir = sum(
        (siswa.hadir or 0)
        for siswa in siswa_qs
    )

    total_izin = sum(
        (siswa.izin or 0)
        for siswa in siswa_qs
    )

    total_sakit = sum(
        (siswa.sakit or 0)
        for siswa in siswa_qs
    )

    total_alfa = sum(
        (siswa.alfa or 0)
        for siswa in siswa_qs
    )

    total_kehadiran = (
        total_hadir +
        total_izin +
        total_sakit +
        total_alfa
    )

    if total_kehadiran > 0:
        rata_kehadiran = round(
            (total_hadir / total_kehadiran) * 100,
            2
        )
    else:
        rata_kehadiran = 0

    # ==============================
    # KETERLAMBATAN
    # ==============================

    keterlambatan = sum(
        (siswa.terlambat_tugas or 0)
        for siswa in siswa_qs
    )

    # ==============================
    # DAFTAR KELAS
    # ==============================

    daftar_kelas = list(
        Kelas.objects.filter(
            siswa__isnull=False
        ).distinct().order_by("tingkat", "jurusan", "nama")
    )

    daftar_kelas = [
        format_nama_kelas(k)
        for k in daftar_kelas
    ]

    # ==============================
    # TOP 10
    # ==============================

    top10_qs = siswa_qs.order_by(
        "-nilai_akhir"
    )[:10]

    top10 = []

    for siswa in top10_qs:
        top10.append({
            "nama": siswa.nama,
            "nis": siswa.nis,
            "nama_kelas": (
                format_nama_kelas(siswa.kelas)
                if siswa.kelas
                else "-"
            ),
            "nilai_akhir": round(
                siswa.nilai_akhir or 0,
                2
            ),
            "status_kelulusan": (
                siswa.status_kelulusan
                or "-"
            ),
        })

    # ==============================
    # GRAFIK NILAI
    # ==============================

    grafik_nilai = [
        {
            "nama": item["nama"],
            "nilai_akhir": item["nilai_akhir"],
        }
        for item in top10
    ]

    # ==============================
    # NILAI PER KELAS
    # ==============================

    nilai_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        data_kelas = (
            Siswa.objects
            .filter(
                kelas__nama=nama,
                kelas__tingkat=tingkat,
                kelas__jurusan=jurusan,
            )
            .filter(
                nilai_akhir__isnull=False
            )
        )

        rata = (
            data_kelas.aggregate(
                rata=Avg("nilai_akhir")
            )["rata"] or 0
        )

        nilai_per_kelas.append({

            "Kelas": nama_kelas,

            "Rata_Rata": round(
                float(rata),
                2
            )

        })

    # ==============================
    # KEHADIRAN PER KELAS
    # ==============================

    kehadiran_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        data_kelas = Siswa.objects.filter(
            kelas__nama=nama,
            kelas__tingkat=tingkat,
            kelas__jurusan=jurusan,
        )

        hadir_kelas = (
            data_kelas.aggregate(
                total=Sum("hadir")
            )["total"] or 0
        )

        izin_kelas = (
            data_kelas.aggregate(
                total=Sum("izin")
            )["total"] or 0
        )

        sakit_kelas = (
            data_kelas.aggregate(
                total=Sum("sakit")
            )["total"] or 0
        )

        alfa_kelas = (
            data_kelas.aggregate(
                total=Sum("alfa")
            )["total"] or 0
        )

        total_kelas = (
            hadir_kelas +
            izin_kelas +
            sakit_kelas +
            alfa_kelas
        )

        if total_kelas > 0:

            persentase_hadir = round(
                (
                    hadir_kelas /
                    total_kelas
                ) * 100,
                2
            )

        else:

            persentase_hadir = 0


        kehadiran_per_kelas.append({

            "Kelas": nama_kelas,

            "Rata_Rata_Hadir":
                persentase_hadir

        })

    # ==============================
    # DISTRIBUSI NILAI
    # ==============================

    distribusi_nilai = [

        {
            "label": "90 - 100",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=90,
                nilai_akhir__lte=100
            ).count()
        },

        {
            "label": "80 - 89",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=80,
                nilai_akhir__lt=90
            ).count()
        },

        {
            "label": "70 - 79",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=70,
                nilai_akhir__lt=80
            ).count()
        },

        {
            "label": "60 - 69",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=60,
                nilai_akhir__lt=70
            ).count()
        },

        {
            "label": "< 60",

            "jumlah": siswa_qs.filter(
                nilai_akhir__lt=60
            ).count()
        },

    ]

    # ==============================
    # KETERLAMBATAN PER KELAS
    # ==============================

    keterlambatan_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        jumlah = (
            Siswa.objects
            .filter(
                kelas__nama=nama,
                kelas__tingkat=tingkat,
                kelas__jurusan=jurusan,
            )
            .aggregate(
                total=Sum("terlambat_tugas")
            )["total"] or 0
        )

        keterlambatan_per_kelas.append({

            "Kelas": nama_kelas,

            "Jumlah_Terlambat":
                jumlah

        })

    # ==============================
    # RESPONSE
    # ==============================

    data = {
        "kelas_dipilih": kelas if kelas else "Semua Kelas",

        "total": total,
        "lulus": lulus,
        "tidak_lulus": tidak_lulus,

        "persentase_lulus": persentase_lulus,
        "persentase_tidak_lulus": persentase_tidak_lulus,

        "rata_rata": round(
            rata_rata or 0,
            2
        ),

        "nilai_tertinggi": round(
            nilai_tertinggi or 0,
            2
        ),

        "nilai_terendah": round(
            nilai_terendah or 0,
            2
        ),

        "rata_kehadiran": rata_kehadiran,

        "daftar_kelas": daftar_kelas,

        "top10": top10,

        "grafik_nilai": grafik_nilai,

        "kehadiran": {
            "hadir": total_hadir,
            "izin": total_izin,
            "sakit": total_sakit,
            "alfa": total_alfa,
        },

        "keterlambatan": keterlambatan,

        "nilai_per_kelas": nilai_per_kelas,

        "kehadiran_per_kelas": kehadiran_per_kelas,

        "distribusi_nilai": distribusi_nilai,

        "keterlambatan_per_kelas": keterlambatan_per_kelas,
    }

    return JsonResponse(data)



def siswa_list_api(request):

    kelas = request.GET.get("kelas", "").strip()
    status = request.GET.get("status", "").strip()

    siswa_qs = Siswa.objects.select_related("kelas").all()

    if kelas and kelas != "Semua":
        nama, tingkat, jurusan = _parse_kelas(kelas)
        siswa_qs = siswa_qs.filter(
            kelas__nama=nama,
            kelas__tingkat=tingkat,
            kelas__jurusan=jurusan,
        )

    if status and status != "Semua":
        siswa_qs = siswa_qs.filter(status_kelulusan__iexact=status)

    siswa_qs = siswa_qs.order_by("nama")

    siswa_detail = []

    for siswa in siswa_qs:

        siswa_detail.append({

            "NIS": siswa.nis,

            "Nama_Siswa": siswa.nama,

            "Kelas": (
                format_nama_kelas(siswa.kelas)
                if siswa.kelas
                else "-"
            ),

            "Jenis_Kelamin": siswa.jenis_kelamin,

            "Nilai_Akhir": (
                round(
                    float(siswa.nilai_akhir or 0),
                    2
                )
            ),

            "Status_Kelulusan": (
                siswa.status_kelulusan
                or "-"
            ),

        })

    return JsonResponse({

        "siswa_detail":
            siswa_detail,

    })


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
def dashboard_data(request):

    # =====================================================
    # FILTER KELAS
    # =====================================================

    kelas = request.GET.get("kelas", "").strip()

    if kelas and kelas != "Semua Kelas":
        nama, tingkat, jurusan = _parse_kelas(kelas)
        siswa_qs = siswa_qs.filter(
            kelas__nama=nama,
            kelas__tingkat=tingkat,
            kelas__jurusan=jurusan,
        )
        kelas_dipilih = kelas
    else:
        siswa_qs = Siswa.objects.all()
        kelas_dipilih = "Semua Kelas"


    # =====================================================
    # STATISTIK SISWA
    # =====================================================

    total = siswa_qs.count()

    lulus = siswa_qs.filter(
        status_kelulusan="Lulus"
    ).count()

    tidak_lulus = siswa_qs.filter(
        status_kelulusan="Tidak Lulus"
    ).count()


    # =====================================================
    # NILAI
    # =====================================================

    nilai_qs = siswa_qs.filter(
        nilai_akhir__isnull=False
    )

    rata_rata = (
        nilai_qs.aggregate(
            rata=Avg("nilai_akhir")
        )["rata"] or 0
    )

    nilai_tertinggi = (
        nilai_qs.order_by("-nilai_akhir")
        .values_list("nilai_akhir", flat=True)
        .first()
        or 0
    )

    nilai_terendah = (
        nilai_qs.order_by("nilai_akhir")
        .values_list("nilai_akhir", flat=True)
        .first()
        or 0
    )


    # =====================================================
    # PERSENTASE KELULUSAN
    # =====================================================

    if total > 0:

        persentase_lulus = round(
            (lulus / total) * 100,
            2
        )

        persentase_tidak_lulus = round(
            (tidak_lulus / total) * 100,
            2
        )

    else:

        persentase_lulus = 0
        persentase_tidak_lulus = 0


    # =====================================================
    # KEHADIRAN
    # =====================================================

    hadir = siswa_qs.aggregate(
        total=Sum("hadir")
    )["total"] or 0

    izin = siswa_qs.aggregate(
        total=Sum("izin")
    )["total"] or 0

    sakit = siswa_qs.aggregate(
        total=Sum("sakit")
    )["total"] or 0

    alfa = siswa_qs.aggregate(
        total=Sum("alfa")
    )["total"] or 0


    total_kehadiran = (
        hadir +
        izin +
        sakit +
        alfa
    )

    if total_kehadiran > 0:

        rata_kehadiran = round(
            (hadir / total_kehadiran) * 100,
            2
        )

    else:

        rata_kehadiran = 0


    # =====================================================
    # KETERLAMBATAN TUGAS
    # =====================================================

    keterlambatan = siswa_qs.aggregate(
        total=Sum("terlambat_tugas")
    )["total"] or 0


    # =====================================================
    # DAFTAR KELAS
    # =====================================================

    daftar_kelas = list(
        Kelas.objects.filter(
            siswa__isnull=False
        ).distinct().order_by("tingkat", "jurusan", "nama")
    )

    daftar_kelas = [
        format_nama_kelas(k)
        for k in daftar_kelas
    ]


    # =====================================================
    # TOP 10
    # =====================================================

    top10_qs = (
        siswa_qs
        .filter(
            nilai_akhir__isnull=False
        )
        .order_by(
            "-nilai_akhir"
        )[:10]
    )


    top10 = []

    for siswa in top10_qs:

        top10.append({

            "nama": siswa.nama,

            "nis": siswa.nis,

            "nama_kelas": (
                format_nama_kelas(siswa.kelas)
                if siswa.kelas
                else "-"
            ),

            "nilai_akhir": round(
                float(siswa.nilai_akhir or 0),
                2
            ),

            "status_kelulusan": (
                siswa.status_kelulusan
                or "-"
            ),

        })


    # =====================================================
    # GRAFIK NILAI
    # =====================================================

    grafik_nilai = [

        {
            "nama": siswa["nama"],
            "nilai_akhir": siswa["nilai_akhir"]
        }

        for siswa in top10

    ]


    # =====================================================
    # NILAI PER KELAS
    # =====================================================

    nilai_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        data_kelas = (
            Siswa.objects
            .filter(
                kelas__nama=nama,
                kelas__tingkat=tingkat,
                kelas__jurusan=jurusan,
            )
            .filter(
                nilai_akhir__isnull=False
            )
        )

        rata = (
            data_kelas.aggregate(
                rata=Avg("nilai_akhir")
            )["rata"] or 0
        )

        nilai_per_kelas.append({

            "Kelas": nama_kelas,

            "Rata_Rata": round(
                float(rata),
                2
            )

        })


    # =====================================================
    # KEHADIRAN PER KELAS
    # =====================================================

    kehadiran_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        data_kelas = Siswa.objects.filter(
            kelas__nama=nama,
            kelas__tingkat=tingkat,
            kelas__jurusan=jurusan,
        )

        hadir_kelas = (
            data_kelas.aggregate(
                total=Sum("hadir")
            )["total"] or 0
        )

        izin_kelas = (
            data_kelas.aggregate(
                total=Sum("izin")
            )["total"] or 0
        )

        sakit_kelas = (
            data_kelas.aggregate(
                total=Sum("sakit")
            )["total"] or 0
        )

        alfa_kelas = (
            data_kelas.aggregate(
                total=Sum("alfa")
            )["total"] or 0
        )

        total_kelas = (
            hadir_kelas +
            izin_kelas +
            sakit_kelas +
            alfa_kelas
        )

        if total_kelas > 0:

            persentase_hadir = round(
                (
                    hadir_kelas /
                    total_kelas
                ) * 100,
                2
            )

        else:

            persentase_hadir = 0


        kehadiran_per_kelas.append({

            "Kelas": nama_kelas,

            "Rata_Rata_Hadir":
                persentase_hadir

        })


    # =====================================================
    # DISTRIBUSI NILAI
    # =====================================================

    distribusi_nilai = [

        {
            "label": "90 - 100",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=90,
                nilai_akhir__lte=100
            ).count()
        },

        {
            "label": "80 - 89",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=80,
                nilai_akhir__lt=90
            ).count()
        },

        {
            "label": "70 - 79",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=70,
                nilai_akhir__lt=80
            ).count()
        },

        {
            "label": "60 - 69",

            "jumlah": siswa_qs.filter(
                nilai_akhir__gte=60,
                nilai_akhir__lt=70
            ).count()
        },

        {
            "label": "< 60",

            "jumlah": siswa_qs.filter(
                nilai_akhir__lt=60
            ).count()
        },

    ]


    # =====================================================
    # KETERLAMBATAN PER KELAS
    # =====================================================

    keterlambatan_per_kelas = []

    for nama_kelas in daftar_kelas:

        nama, tingkat, jurusan = _parse_kelas(nama_kelas)

        jumlah = (
            Siswa.objects
            .filter(
                kelas__nama=nama,
                kelas__tingkat=tingkat,
                kelas__jurusan=jurusan,
            )
            .aggregate(
                total=Sum("terlambat_tugas")
            )["total"] or 0
        )

        keterlambatan_per_kelas.append({

            "Kelas": nama_kelas,

            "Jumlah_Terlambat":
                jumlah

        })


    # =====================================================
    # RESPONSE
    # =====================================================

    return Response({

        "kelas_dipilih":
            kelas_dipilih,

        "total":
            total,

        "lulus":
            lulus,

        "tidak_lulus":
            tidak_lulus,

        "persentase_lulus":
            persentase_lulus,

        "persentase_tidak_lulus":
            persentase_tidak_lulus,

        "rata_rata":
            round(
                float(rata_rata),
                2
            ),

        "nilai_tertinggi":
            round(
                float(nilai_tertinggi),
                2
            ),

        "nilai_terendah":
            round(
                float(nilai_terendah),
                2
            ),

        "rata_kehadiran":
            rata_kehadiran,

        "kehadiran": {

            "hadir":
                hadir,

            "izin":
                izin,

            "sakit":
                sakit,

            "alfa":
                alfa,

        },

        "keterlambatan":
            keterlambatan,

        "daftar_kelas":
            daftar_kelas,

        "top10":
            top10,

        "grafik_nilai":
            grafik_nilai,

        "nilai_per_kelas":
            nilai_per_kelas,

        "kehadiran_per_kelas":
            kehadiran_per_kelas,

        "distribusi_nilai":
            distribusi_nilai,

        "keterlambatan_per_kelas":
            keterlambatan_per_kelas,

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