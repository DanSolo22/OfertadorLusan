from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic.base import View

from ofertador.forms import CargarOferta

import csv


class Index(View):
    def get(self, request):
        form = CargarOferta()
        return render(request, 'index.html', {'form': form})

    def post(self, request):
        if request.POST:
            form = CargarOferta(request.POST, request.FILES)
            if form.is_valid():
                oferta = form.cleaned_data.get('oferta')

                with open(oferta) as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    line_count = 0
                    for row in csv_reader:
                        if line_count == 0:
                            print(f'Column names are {", ".join(row)}')
                            line_count += 1
                        else:
                            print(f'\t{row[0]} works in the {row[1]} department, and was born in {row[2]}.')
                            line_count += 1
                    print(f'Processed {line_count} lines.')

                return HttpResponse('Recibido correctamente')
            else:
                return HttpResponse('Error')
