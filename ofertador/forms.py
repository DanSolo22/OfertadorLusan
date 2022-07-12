from django import forms
from .validators import validate_file_extension


class CargarOferta(forms.Form):
    oferta = forms.FileField(required=True, validators=[validate_file_extension])
