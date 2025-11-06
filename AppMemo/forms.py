from django import forms
from .models import EtatSortie


class EtatSortieForm(forms.ModelForm):
    class Meta:
        model = EtatSortie
        fields = '__all__'
        widgets = {
            'date': forms.DateInput(attrs={'type': 'date'}),
            'marchandises': forms.Textarea(attrs={'rows':3}),
            'notes': forms.Textarea(attrs={'rows':3}),
        }
