from pathlib import Path
import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from rest_framework import viewsets, filters
from rest_framework.decorators import api_view
from rest_framework.response import Response
from .forms import UploadExcelForm
from .models import ImportedFile, Kelas, Siswa
from .serializers import SiswaSerializer, KelasSerializer
from django.http import JsonResponse
from django.contrib.auth.decorators import login_required


def format_nama_kelas(kelas):
    nama = (kelas.nama or "").strip()
    tingkat = (kelas.tingkat or "").strip()
    jurusan = (kelas.jurusan or "").strip()

    # Jika nama sudah diawali tingkat,
    # jangan tambahkan tingkat dan jurusan lagi.
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


@api_view(["GET"])
def dashboard_data(request):
    def nama_kelas_tampilan(kelas):
        if not kelas:
            return "-"
        if kelas.tingkat and kelas.jurusan and kelas.nama:
            if kelas.nama.startswith(kelas.tingkat):
                return kelas.nama
            return f"{kelas.tingkat} {kelas.jurusan} {kelas.nama}"
        return kelas.nama

    siswa = Siswa.objects.select_related("kelas").all()
    kelas_filter = request.GET.get("kelas")

    if kelas_filter and kelas_filter != "Semua":
        bagian = kelas_filter.split()

        if len(bagian) >= 3:
            tingkat = bagian[0]
            jurusan = bagian[1]
            nama = " ".join(bagian[2:])

            siswa = siswa.filter(
                kelas__tingkat=tingkat,
                kelas__jurusan=jurusan,
                kelas__nama=nama,
            )
        else:
            siswa = siswa.filter(kelas__nama__iexact=kelas_filter)

    total = siswa.count()
    lulus = siswa.filter(status_kelulusan="Lulus").count()
    tidak_lulus = siswa.filter(status_kelulusan="Tidak Lulus").count()

    nilai = [s.nilai_akhir for s in siswa if s.nilai_akhir is not None]
    rata_rata = round(sum(nilai) / len(nilai), 2) if nilai else 0

    if total:
        persentase_lulus = round(lulus / total * 100, 2)
        persentase_tidak_lulus = round(tidak_lulus / total * 100, 2)
    else:
        persentase_lulus = 0
        persentase_tidak_lulus = 0

    top10 = siswa.order_by("-nilai_akhir")[:10]

    top10_data = [
        {
            "nama": s.nama,
            "nis": s.nis,
            "kelas": nama_kelas_tampilan(s.kelas),
            "nilai_akhir": s.nilai_akhir,
            "status_kelulusan": s.status_kelulusan,
        }
        for s in top10
    ]

    grafik_nilai = [
        {
            "nama": s.nama,
            "nilai": float(s.nilai_akhir or 0)
        }
        for s in siswa.order_by("-nilai_akhir")
    ]

    total_hadir = sum(s.hadir for s in siswa)
    total_izin = sum(s.izin for s in siswa)
    total_sakit = sum(s.sakit for s in siswa)
    total_alfa = sum(s.alfa for s in siswa)
    total_terlambat = sum(s.terlambat_tugas for s in siswa)

    kelas_queryset = Kelas.objects.all().order_by("tingkat", "jurusan", "nama")

    daftar_kelas = [
    format_nama_kelas(kelas)
    for kelas in Kelas.objects.all()
]

    return Response(
        {
            "kelas_dipilih": kelas_filter or "Semua",
            "total": total,
            "lulus": lulus,
            "tidak_lulus": tidak_lulus,
            "persentase_lulus": persentase_lulus,
            "persentase_tidak_lulus": persentase_tidak_lulus,
            "rata_rata": rata_rata,
            "daftar_kelas": daftar_kelas,
            "top10": top10_data,

            "grafik_nilai": grafik_nilai,

            "kehadiran": {
                "hadir": total_hadir,
                "izin": total_izin,
                "sakit": total_sakit,
                "alfa": total_alfa,
            },
            "keterlambatan": total_terlambat,
        }
    )


def _parse_kelas(kelas_name):
    """Parse string kelas menjadi (nama, tingkat, jurusan)."""
    if not kelas_name:
        return None, None, None
    parts = str(kelas_name).split()
    nama = parts[-1] if len(parts) >= 1 else str(kelas_name)
    jurusan = parts[-2] if len(parts) >= 3 else None
    tingkat = parts[0] if len(parts) >= 3 else None
    return nama, tingkat, jurusan


def _normalize_header(value) -> str:
    """Normalisasi header Excel untuk matching alias kolom."""
    s = "" if value is None else str(value)
    s = s.replace("\u00a0", " ")
    s = " ".join(s.split())
    return s.strip().lower()


def _pick_column(df_columns, candidates):
    """Ambil nama kolom asli dari df yang cocok dengan salah satu kandidat."""
    normalized_cols = {_normalize_header(c): c for c in df_columns}
    for cand in candidates:
        cand_norm = _normalize_header(cand)
        if cand_norm in normalized_cols:
            return normalized_cols[cand_norm]
    return None


def _clean_cell(value):
    """Pandas NaN -> None, string trim & kosong -> None."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, str):
        s = " ".join(value.split()).strip()
        return s if s else None
    return value


def _normalize_jenis_kelamin(value):
    """Normalisasi jenis kelamin agar serializer menerima 'L' atau 'P'."""
    v = _clean_cell(value)
    if v is None:
        return None

    s = str(v).strip().lower()
    s = s.replace("laki-laki", "l")
    s = s.replace("perempuan", "p")
    s = s.replace("pria", "l").replace("wanita", "p")

    if s in {
        "l", "la", "pria", "lk", "laki", "laki laki", "laki-laki", "1", "m",
        "male", "cowok",
    }:
        return "L"
    if s in {
        "p", "pa", "perempuan", "pr", "wanita", "w", "f", "female", "cewek",
        "2",
    }:
        return "P"

    if s.upper() in {"L", "P"}:
        return s.upper()

    return None


def _int_val(row, col):
    v = _clean_cell(row.get(col)) if col else None
    if v is None:
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


@login_required(login_url='login')
def import_excel(request):
    if request.method == "POST":
        form = UploadExcelForm(request.POST, request.FILES)
        if form.is_valid():
            excel_file = request.FILES["file"]

            if not excel_file.name.lower().endswith((".xlsx", ".xls")):
                messages.error(
                    request,
                    "File harus berupa Excel (.xlsx atau .xls).",
                )
                return redirect("import_excel")

            df = pd.read_excel(excel_file)
            df.columns = df.columns.astype(str).str.strip()

            col_no = _pick_column(df.columns, ["no", "nomor"])
            col_nis = _pick_column(df.columns, ["nis", "nim"])
            col_nama = _pick_column(df.columns, ["nama", "nama siswa", "nama_siswa", "nama_lengkap"])
            col_jk = _pick_column(df.columns, ["jenis kelamin", "jenis_kelamin", "jk", "gender"])
            col_kelas = _pick_column(df.columns, ["kelas", "kelas siswa"])
            col_alamat = _pick_column(df.columns, ["alamat", "address"])
            col_hadir = _pick_column(df.columns, ["hadir"])
            col_izin = _pick_column(df.columns, ["izin"])
            col_sakit = _pick_column(df.columns, ["sakit"])
            col_alfa = _pick_column(df.columns, ["alfa"])
            col_tugas = _pick_column(df.columns, ["tugas"])
            col_terlambat = _pick_column(df.columns, ["terlambat tugas", "terlambat_tugas"])
            col_uts = _pick_column(df.columns, ["uts"])
            col_uas = _pick_column(df.columns, ["uas"])
            col_nilai = _pick_column(df.columns, ["nilai akhir", "nilai_akhir"])
            col_status = _pick_column(df.columns, ["status kelulusan", "status_kelulusan"])

            required_display = []
            if not col_nis:
                required_display.append("NIS")
            if not col_nama:
                required_display.append("Nama Siswa")
            if not col_jk:
                required_display.append("Jenis Kelamin")
            if not col_kelas:
                required_display.append("Kelas")

            if required_display:
                messages.error(
                    request,
                    f"Kolom tidak ditemukan: {required_display}",
                )
                return redirect("import_excel")

            berhasil = 0

            with transaction.atomic():
                for _, row in df.iterrows():
                    nis = _clean_cell(row.get(col_nis))
                    if nis is None:
                        continue
                    nis = str(int(nis)) if isinstance(nis, float) else str(nis)

                    if Siswa.objects.filter(nis=nis).exists():
                        continue

                    jk = _normalize_jenis_kelamin(row.get(col_jk))
                    if not jk:
                        continue

                    nama = _clean_cell(row.get(col_nama))
                    if not nama:
                        continue

                    alamat = _clean_cell(row.get(col_alamat)) if col_alamat else None

                    no = _clean_cell(row.get(col_no)) if col_no else None
                    if no is not None:
                        try:
                            no = int(no)
                        except (TypeError, ValueError):
                            no = None

                    kelas_name = _clean_cell(row.get(col_kelas))
                    kelas_obj = None
                    if kelas_name:
                        nama_kls, tingkat, jurusan = _parse_kelas(kelas_name)
                        kelas_obj = Kelas.objects.filter(
                            nama=nama_kls,
                            tingkat=tingkat,
                            jurusan=jurusan,
                        ).first()
                        if not kelas_obj:
                            kelas_obj = Kelas.objects.create(
                                nama=nama_kls,
                                tingkat=tingkat,
                                jurusan=jurusan,
                            )

                    if not kelas_obj:
                        continue

                    nilai_akhir = _clean_cell(row.get(col_nilai)) if col_nilai else None
                    if nilai_akhir is not None:
                        try:
                            nilai_akhir = float(nilai_akhir)
                        except (TypeError, ValueError):
                            nilai_akhir = None

                    status = _clean_cell(row.get(col_status)) if col_status else None

                    Siswa.objects.create(
                        no=no,
                        nis=nis,
                        nama=nama,
                        jenis_kelamin=jk,
                        alamat=alamat,
                        kelas=kelas_obj,
                        hadir=_int_val(row, col_hadir),
                        izin=_int_val(row, col_izin),
                        sakit=_int_val(row, col_sakit),
                        alfa=_int_val(row, col_alfa),
                        tugas=_int_val(row, col_tugas),
                        terlambat_tugas=_int_val(row, col_terlambat),
                        uts=_int_val(row, col_uts),
                        uas=_int_val(row, col_uas),
                        nilai_akhir=nilai_akhir,
                        status_kelulusan=status,
                    )

                    berhasil += 1

            excel_file.seek(0)
            imported = ImportedFile(
                file=excel_file,
                nama_asli=excel_file.name,
                jumlah_data=berhasil,
            )
            imported.save()

            messages.success(request, f"{berhasil} data berhasil diimport.")
            return redirect("import_excel")
    else:
        form = UploadExcelForm()

    return render(request, "import_excel.html", {"form": form})

@login_required(login_url='login')
def dashboard(request):
    return render(request, "dashboard.html")


@login_required(login_url='login')
def data_siswa(request):
    """Halaman terpisah untuk menampilkan daftar data siswa."""
    return render(request, "data_siswa.html")


@login_required(login_url='login')
def status_kelulusan(request):
    """Halaman terpisah untuk menampilkan data siswa berdasarkan status kelulusan."""
    return render(request, "Status_kelulusan.html")


@login_required(login_url='login')
def penyimpanan_file(request):
    """Halaman untuk menampilkan semua file Excel yang telah diimport."""
    files = ImportedFile.objects.all().order_by("-tanggal_upload")
    return render(request, "penyimpanan_file.html", {"files": files})


@login_required(login_url='login')
def hapus_file(request, pk):
    """Menghapus file Excel yang telah diimport dari penyimpanan."""
    file_obj = get_object_or_404(ImportedFile, pk=pk)
    if request.method == "POST":
        if file_obj.file:
            file_obj.file.delete(save=False)
        file_obj.delete()
        messages.success(request, f"File '{file_obj.nama_asli}' berhasil dihapus.")
    return redirect("penyimpanan_file")


@api_view(["GET"])
def dashboard_api(request):
    excel_path = Path(settings.BASE_DIR) / "DataSiswa_Hasil_kelas12.xlsx"
    data = pd.read_excel(excel_path)
    data.columns = data.columns.str.strip().str.replace(" ", "_")

    daftar_kelas = sorted(
        data["Kelas"].dropna().astype(str).str.strip().unique().tolist()
    )

    nilai_kelas = (
        data.groupby("Kelas")["Nilai_Akhir"].mean().round(2).reset_index()
    )
    nilai_kelas = (
        nilai_kelas.rename(columns={"Nilai_Akhir": "Rata_Rata"})
        .to_dict("records")
    )

    kehadiran_kelas = (
        data.groupby("Kelas")["Hadir"].mean().round(2).reset_index()
    )
    kehadiran_kelas = (
        kehadiran_kelas.rename(columns={"Hadir": "Rata_Rata_Hadir"})
        .to_dict("records")
    )

    terlambat_kelas = (
        data.groupby("Kelas")["Terlambat_Tugas"].sum().reset_index()
    )
    terlambat_kelas = (
        terlambat_kelas.rename(columns={"Terlambat_Tugas": "Jumlah_Terlambat"})
        .to_dict("records")
    )

    bins = [0, 50, 60, 70, 75, 80, 90, 100]
    labels_bins = ["0-49", "50-59", "60-69", "70-74", "75-79", "80-89", "90-100"]
    hist = pd.cut(
        data["Nilai_Akhir"], bins=bins, include_lowest=True
    ).value_counts()

    distribusi_nilai = []
    for lo_idx in range(len(bins) - 1):
        lo = bins[lo_idx]
        hi = bins[lo_idx + 1]
        label = f"{lo}-{100 if hi == 100 else hi - 1}"
        cnt = 0
        for interval, c in hist.items():
            if interval.left == lo and interval.right == hi:
                cnt = int(c)
                break
        distribusi_nilai.append({"label": label, "jumlah": cnt})

    kelas = request.GET.get("kelas")
    if kelas and kelas != "Semua":
        data = data[
            data["Kelas"].astype(str).str.strip() == kelas.strip()
        ]

    total = len(data)
    lulus = len(data[data["Status_Kelulusan"] == "Lulus"])
    tidak_lulus = len(data[data["Status_Kelulusan"] == "Tidak Lulus"])
    rata_rata = round(data["Nilai_Akhir"].mean(), 2) if total > 0 else 0
    persentase_lulus = round((lulus / total) * 100, 2) if total > 0 else 0
    nilai_tertinggi = round(data["Nilai_Akhir"].max(), 2) if total > 0 else 0
    nilai_terendah = round(data["Nilai_Akhir"].min(), 2) if total > 0 else 0
    rata_kehadiran = round(data["Hadir"].mean(), 2) if total > 0 else 0

    top10 = (
        data.sort_values(by="Nilai_Akhir", ascending=False)
        .head(10)[
            ["Nama_Siswa", "Kelas", "Nilai_Akhir", "Status_Kelulusan"]
        ]
        .to_dict("records")
    )

    status = request.GET.get("status")
    detail_data = data
    if status and status != "Semua":
        detail_data = detail_data[detail_data["Status_Kelulusan"] == status]

    detail_cols = [
        "No", "NIS", "Nama_Siswa", "Jenis_Kelamin", "Kelas",
        "Hadir", "Izin", "Sakit", "Alfa", "Tugas",
        "Terlambat_Tugas", "UTS", "UAS", "Nilai_Akhir",
        "Status_Kelulusan",
    ]
    siswa_detail = detail_data[detail_cols].to_dict("records")

    return Response(
        {
            "kelas_dipilih": kelas or "Semua",
            "status_dipilih": status or "Semua",
            "daftar_kelas": daftar_kelas,
            "total": total,
            "lulus": lulus,
            "tidak_lulus": tidak_lulus,
            "rata_rata": rata_rata,
            "persentase_lulus": persentase_lulus,
            "nilai_tertinggi": nilai_tertinggi,
            "nilai_terendah": nilai_terendah,
            "rata_kehadiran": rata_kehadiran,
            "nilai_per_kelas": nilai_kelas,
            "kehadiran_per_kelas": kehadiran_kelas,
            "keterlambatan_per_kelas": terlambat_kelas,
            "distribusi_nilai": distribusi_nilai,
            "top10": top10,
            "siswa_detail": siswa_detail,
        }
    )


class SiswaViewSet(viewsets.ModelViewSet):
    queryset = Siswa.objects.all()
    serializer_class = SiswaSerializer

    filter_backends = [
        filters.SearchFilter,
        filters.OrderingFilter,
    ]

    search_fields = ["nama", "nis"]
    ordering_fields = ["nama", "nis", "id"]
    ordering = ["nama"]


@api_view(["GET"])
def analisis_kelas_api(request):

    nama_kelas = request.GET.get("kelas")

    if not nama_kelas:
        return Response({
            "error": "Parameter kelas wajib diisi."
        }, status=400)

    # Ambil seluruh siswa dari database
    siswa = Siswa.objects.select_related(
        "kelas"
    ).all()

    # Buat DataFrame
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
        return Response({
            "error": "Data siswa tidak ditemukan."
        }, status=404)

    # Membuat nama kelas
    df["nama_kelas"] = (
        df["kelas__tingkat"].astype(str)
        + " "
        + df["kelas__jurusan"].astype(str)
        + " "
        + df["kelas__nama"].astype(str)
    )

    # Filter kelas
    df_kelas = df[
        df["nama_kelas"] == nama_kelas
    ].copy()

    if df_kelas.empty:
        return Response({
            "error": f"Kelas '{nama_kelas}' tidak ditemukan."
        }, status=404)

    # Statistik
    total = len(df_kelas)

    lulus = (
        df_kelas["status_kelulusan"]
        .eq("Lulus")
        .sum()
    )

    tidak_lulus = (
        df_kelas["status_kelulusan"]
        .eq("Tidak Lulus")
        .sum()
    )

    rata_rata = (
        df_kelas["nilai_akhir"].mean()
    )

    # Top 10
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

        "lulus": int(lulus),

        "tidak_lulus": int(
            tidak_lulus
        ),

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


class KelasViewSet(viewsets.ModelViewSet):

    queryset = Kelas.objects.all().order_by('id')
    serializer_class = KelasSerializer


@api_view(["GET"])
def dashboard_data_api(request):

    nama_kelas = request.GET.get("kelas")

    # ==============================
    # AMBIL DATA DARI DATABASE
    # ==============================

    siswa = Siswa.objects.select_related("kelas").all()

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
        return Response({
            "error": "Data siswa belum tersedia."
        }, status=404)

    # ==============================
    # MEMBUAT NAMA KELAS
    # ==============================

    df["nama_kelas"] = (
        df["kelas__tingkat"].fillna("").astype(str)
        + " "
        + df["kelas__jurusan"].fillna("").astype(str)
        + " "
        + df["kelas__nama"].fillna("").astype(str)
    ).str.replace(
        r"\s+",
        " ",
        regex=True
    ).str.strip()

    # ==============================
    # DAFTAR KELAS
    # ==============================

    daftar_kelas = sorted(
        df["nama_kelas"]
        .dropna()
        .unique()
        .tolist()
    )

    # ==============================
    # FILTER KELAS
    # ==============================

    if nama_kelas:

        df = df[
            df["nama_kelas"] == nama_kelas
        ].copy()

        if df.empty:
            return Response({
                "error": (
                    f"Kelas '{nama_kelas}' "
                    "tidak ditemukan."
                ),
                "daftar_kelas": daftar_kelas
            }, status=404)

    else:

        nama_kelas = "Semua Kelas"

    # ==============================
    # STATISTIK UTAMA
    # ==============================

    total = len(df)

    lulus = (
        df["status_kelulusan"]
        .eq("Lulus")
        .sum()
    )

    tidak_lulus = (
        df["status_kelulusan"]
        .eq("Tidak Lulus")
        .sum()
    )

    rata_rata = (
        df["nilai_akhir"].mean()
        if total > 0
        else 0
    )

    # ==============================
    # PERSENTASE
    # ==============================

    persentase_lulus = (
        lulus / total * 100
        if total > 0
        else 0
    )

    persentase_tidak_lulus = (
        tidak_lulus / total * 100
        if total > 0
        else 0
    )

    # ==============================
    # TOP 10
    # ==============================

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
                "status_kelulusan"
            ]
        ]
        .to_dict("records")
    )

    # ==============================
    # DATA NILAI
    # ==============================

    grafik_nilai = list(
    siswa
    .values(
        "nama",
        "nilai_akhir"
    )
    .order_by("-nilai_akhir")
)

    grafik_nilai = (
        df[
            [
                "nama",
                "nilai_akhir"
            ]
        ]
        .sort_values(
            by="nilai_akhir",
            ascending=False
        )
        .head(10)
        .to_dict("records")
    )

    # ==============================
    # KEHADIRAN
    # ==============================

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

    # ==============================
    # KETERLAMBATAN
    # ==============================

    keterlambatan = int(
        df["terlambat_tugas"].sum()
    )

    # ==============================
    # RESPONSE API
    # ==============================

    return Response({

        "kelas_dipilih": nama_kelas,

        "total": total,

        "lulus": int(lulus),

        "tidak_lulus": int(
            tidak_lulus
        ),

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

        "grafik_nilai": grafik_nilai,

        "kehadiran": kehadiran,

        "keterlambatan": keterlambatan,
    })