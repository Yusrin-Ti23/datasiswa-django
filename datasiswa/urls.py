from django.contrib import admin
from django.urls import include, path, re_path
from drf_spectacular.views import (
    SpectacularAPIView,
    SpectacularSwaggerView,
)
from django.http import HttpResponse
from django.conf import settings
from django.conf.urls.static import static
from django.shortcuts import redirect
from siswa import views



def home(request):
    if request.user.is_authenticated:
        return redirect('dashboard')

    return redirect('login')

urlpatterns = [
    path("", home, name="home"),
    path('dashboard/',views.dashboard,name='dashboard'),
    path('api/', include('siswa.urls')),
    path('admin/', admin.site.urls),

    path('akun/', include('akun.urls')),
    path('api/schema/', SpectacularAPIView.as_view(), name='schema'),
    path('api/schema/docs/',SpectacularSwaggerView.as_view(url_name='schema'),name='schema-docs',),
]

# Sajikan file media (upload file Excel) saat DEBUG mode
from django.conf import settings as dj_settings
if dj_settings.DEBUG:
    urlpatterns += static(dj_settings.MEDIA_URL, document_root=dj_settings.MEDIA_ROOT)
