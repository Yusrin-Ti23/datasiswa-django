from django.urls import path, include
from rest_framework.routers import DefaultRouter

from . import views
from .views import (
    SiswaViewSet,
    KelasViewSet,
    import_excel,
    dashboard,
    dashboard_data_api,
    analisis_kelas_api,
    data_siswa,
    status_kelulusan,
    penyimpanan_file,
    hapus_file,
    siswa_list_api,
)


# ==========================================================
# ROUTER REST API
# ==========================================================

router = DefaultRouter()

router.register(
    r"rsiswa",
    SiswaViewSet,
    basename="siswa"
)

router.register(
    r"kelas",
    KelasViewSet,
    basename="kelas"
)


# ==========================================================
# URL PATTERNS
# ==========================================================

urlpatterns = [

    # ------------------------------------------------------
    # IMPORT EXCEL
    # ------------------------------------------------------

    path(
        "import_excel/",
        import_excel,
        name="import_excel"
    ),

    # ------------------------------------------------------
    # REST API ROUTER
    # ------------------------------------------------------

    path(
        "",
        include(router.urls)
    ),

    # ------------------------------------------------------
    # HALAMAN DASHBOARD
    # ------------------------------------------------------

    path(
        "dashboard/",
        dashboard,
        name="dashboard"
    ),

    # ------------------------------------------------------
    # API DASHBOARD
    # ------------------------------------------------------

    path(
        "dashboard-data/",
        dashboard_data_api,
        name="dashboard_data_api"
    ),

    # ------------------------------------------------------
    # API SISWA LIST (untuk halaman detail)
    # ------------------------------------------------------

    path(
        "siswa-list/",
        siswa_list_api,
        name="siswa_list_api"
    ),

    # ------------------------------------------------------
    # API ANALISIS KELAS
    # ------------------------------------------------------

    path(
        "analisis-kelas/",
        analisis_kelas_api,
        name="analisis_kelas_api"
    ),

    # ------------------------------------------------------
    # DATA SISWA
    # ------------------------------------------------------

    path(
        "data_siswa/",
        data_siswa,
        name="data_siswa"
    ),

    # ------------------------------------------------------
    # STATUS KELULUSAN
    # ------------------------------------------------------

    path(
        "status_kelulusan/",
        status_kelulusan,
        name="status_kelulusan"
    ),

    # ------------------------------------------------------
    # PENYIMPANAN FILE
    # ------------------------------------------------------

    path(
        "penyimpanan_file/",
        penyimpanan_file,
        name="penyimpanan_file"
    ),

    # ------------------------------------------------------
    # HAPUS FILE
    # ------------------------------------------------------

    path(
        "hapus_file/<int:pk>/",
        hapus_file,
        name="hapus_file"
    ),
]