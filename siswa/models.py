from django.db import models


class Kelas(models.Model):

    nama = models.CharField(max_length=50)
    tingkat = models.CharField(max_length=20, blank=True, null=True)
    jurusan = models.CharField(max_length=100, blank=True, null=True)

    def __str__(self):
        return f"{self.nama} ({self.tingkat or '-'} - {self.jurusan or '-'})"


class Siswa(models.Model):
    JENIS_KELAMIN_CHOICES = [
        ('L', 'Laki-laki'),
        ('P', 'Perempuan'),
    ]

    no = models.PositiveIntegerField(null=True, blank=True)
    nama = models.CharField(max_length=100)
    nis = models.CharField(max_length=20)
    jenis_kelamin = models.CharField(
        max_length=1,
        choices=JENIS_KELAMIN_CHOICES,
    )
    alamat = models.TextField(blank=True, null=True)

    kelas = models.ForeignKey(
        Kelas,
        on_delete=models.CASCADE,
        related_name='siswa',
    )

    # --- Kolom akademik dari Excel ---
    hadir = models.PositiveIntegerField(default=0)
    izin = models.PositiveIntegerField(default=0)
    sakit = models.PositiveIntegerField(default=0)
    alfa = models.PositiveIntegerField(default=0)
    tugas = models.PositiveIntegerField(default=0)
    terlambat_tugas = models.PositiveIntegerField(default=0)
    uts = models.PositiveIntegerField(default=0)
    uas = models.PositiveIntegerField(default=0)
    nilai_akhir = models.FloatField(default=0, null=True, blank=True)
    status_kelulusan = models.CharField(
        max_length=20,
        blank=True,
        null=True,
    )

    def __str__(self):
        return self.nama


class ImportedFile(models.Model):
    """Menyimpan file Excel yang telah diimport ke sistem."""
    file = models.FileField(upload_to='uploads/')
    nama_asli = models.CharField(max_length=255)
    jumlah_data = models.PositiveIntegerField(default=0, blank=True, null=True)
    tanggal_upload = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.nama_asli
