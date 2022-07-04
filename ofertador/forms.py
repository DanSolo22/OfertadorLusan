from django import forms
from django.contrib.auth.forms import UserCreationForm
from django.forms import ModelForm

from ofertador.models import *


class CargarOferta(forms.Form):
    oferta = forms.FileField(required=True)
