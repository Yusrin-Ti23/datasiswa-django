from rest_framework import serializers
from .models import Siswa, Kelas


class SiswaSerializer(serializers.ModelSerializer):

    class Meta:
        model = Siswa
        fields = '__all__'

    def validate_nis(self, value):
        value = str(value)

        if len(value) < 3:
            raise serializers.ValidationError(
                "NIS minimal terdiri dari 3 karakter."
            )

        return value

    def validate_nama(self, value):

        if len(value) < 3:
            raise serializers.ValidationError(
                "Nama minimal 3 karakter."
            )

        return value

    def validate_jenis_kelamin(self, value):

        if value not in ['L', 'P']:
            raise serializers.ValidationError(
                "Jenis kelamin harus L atau P."
            )

        return value


class KelasSerializer(serializers.ModelSerializer):

    class Meta:
        model = Kelas
        fields = '__all__'