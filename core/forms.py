from django import forms
from .models import FileConversion

class FileConversionForm(forms.ModelForm):
    class Meta:
        model = FileConversion
        fields = ['original_file', 'target_format']
        widgets = {
            'original_file': forms.FileInput(attrs={
                'class': 'form-control',
                'accept': '.xlsx,.xls,.csv,.json,.docx,.pdf'
            }),
            'target_format': forms.Select(attrs={
                'class': 'form-control'
            }),
        }
