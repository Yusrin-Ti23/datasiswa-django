import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'datasiswa.settings')
django.setup()

from rest_framework.test import APIRequestFactory
from siswa.views import dashboard_api

factory = APIRequestFactory()

# Test all data
request = factory.get('/api/dashboard/api/')
response = dashboard_api(request)
data = response.data
print('=== SEMUA ===')
print('total:', data['total'])
print('lulus:', data['lulus'])
print('tidak_lulus:', data['tidak_lulus'])
print('rata_rata:', data['rata_rata'])
print('persentase_lulus:', data['persentase_lulus'])
print('nilai_tertinggi:', data['nilai_tertinggi'])
print('nilai_terendah:', data['nilai_terendah'])
print('rata_kehadiran:', data['rata_kehadiran'])
print('daftar_kelas:', data['daftar_kelas'])
print('nilai_per_kelas:', data['nilai_per_kelas'])
print('distribusi_nilai:', data['distribusi_nilai'])
print('top10 (n):', len(data['top10']))

# Test filter kelas
request2 = factory.get('/api/dashboard/api/?kelas=XII IPA A')
response2 = dashboard_api(request2)
data2 = response2.data
print('\n=== FILTER XII IPA A ===')
print('total:', data2['total'])
print('lulus:', data2['lulus'])
print('tidak_lulus:', data2['tidak_lulus'])
print('rata_rata:', data2['rata_rata'])
print('nilai_per_kelas (harus tetap semua kelas):', data2['nilai_per_kelas'])
print('top10 (n):', len(data2['top10']))
