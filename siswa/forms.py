from django import forms

class UploadExcelForm(forms.Form):
    file = forms.FileField(
        label="Pilih File Excel"
    )