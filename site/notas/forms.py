from django import forms

class JustificativasForm(forms.ModelForm):
    class Meta:
        model = Justificativa
        fields = '__all__'
        
    