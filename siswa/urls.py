from django.urls import path, include
from rest_framework.routers import DefaultRouter

from .views import SiswaViewSet, import_excel
from django.urls import path
from . import views
from .views import dashboard_api, data_siswa, penyimpanan_file, status_kelulusan, hapus_file
from .views import SiswaViewSet, KelasViewSet

router = DefaultRouter()

router.register(
    r'rsiswa',
    SiswaViewSet,
    basename='siswa'
)

router.register(
    r'kelas',
    KelasViewSet,
    basename='kelas'
)


urlpatterns = [
    path('import_excel/', import_excel, name='import_excel'),
    path('', include(router.urls)),
    path("dashboard/api/",dashboard_api,name="dashboard_api"),
    path('data_siswa/', data_siswa, name='data_siswa'),
    path('status_kelulusan/', status_kelulusan, name='status_kelulusan'),
    path('penyimpanan_file/', penyimpanan_file, name='penyimpanan_file'),
    path('hapus_file/<int:pk>/', hapus_file, name='hapus_file'),
    path("dashboard-data/",views.dashboard_data,name="dashboard_data"),
    path("analisis-kelas/",views.analisis_kelas_api,name="analisis_kelas_api"),

    
]
