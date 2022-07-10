import docx
from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic.base import View
from docx.shared import Inches

from ofertador.forms import CargarOferta

import csv
from docxtpl import DocxTemplate


class Index(View):
    def get(self, request):
        form = CargarOferta()
        return render(request, 'index.html', {'form': form})

    def post(self, request):
        if request.POST:
            form = CargarOferta(request.POST, request.FILES)
            if form.is_valid():
                oferta = form.cleaned_data.get('oferta')

                with open('csvofertas/oferta.csv', 'wb+') as destination:
                    for chunk in oferta.chunks():
                        destination.write(chunk)

                oferta = ''
                fecha = ''
                validez = ''
                cliente = ''
                proveedor = ''
                rsoc = ''
                empresa = ''
                dir = ''
                cp = ''
                pob = ''
                pro = ''
                tel = ''
                fax = ''
                moneda = ''
                des_moneda = ''

                doc = DocxTemplate("C:/ofertas/plantilla.docx")

                with open('csvofertas/oferta.csv') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    line_count = 0

                    for row in csv_reader:
                        if line_count == 0:
                            print(f'Column names are: {", ".join(row)}')
                            line_count += 1
                        elif line_count == 1:
                            oferta = row[0]
                            fecha = row[1]
                            validez = row[2]
                            cliente = row[3]
                            proveedor = row[4]
                            rsoc = row[5]
                            empresa = row[6]
                            dir = row[7]
                            cp = row[8]
                            pob = row[9]
                            tel = row[10]
                            fax = row[11]
                            moneda = row[15]
                            des_moneda = row[16]
                            print(f'\t{oferta}, {fecha}, {validez}, {cliente}, {proveedor}.')
                            line_count += 1
                        else:
                            print(row)
                            line_count += 1

                    print(f'Processed {line_count} lines.')

                context = \
                    {
                        'OFERTA': oferta,
                        'FECHA': fecha,
                        'VALIDEZ': validez,
                        'CLIENTE': cliente,
                        'PROVEEDOR': proveedor,
                        'RSOC': rsoc,
                        'EMPRESA': empresa,
                        'DIR': dir,
                        'CP': cp,
                        'POB': pob,
                        'PRO': pro,
                        'TEL': tel,
                        'FAX': fax,
                        'MONEDA': moneda,
                        'DES_MON': des_moneda,
                    }

                doc.render(context)
                doc.save("C:/ofertas/oferta.docx")

                doc = docx.Document("C:/ofertas/oferta.docx")

                table = doc.add_table(rows=1, cols=5)
                celda = 0

                for cell in table.columns[0].cells:
                    if celda == 0:
                        cell.width = Inches(3.5)
                    if celda == 1:
                        cell.width = Inches(1)
                    if celda == 2:
                        cell.width = Inches(1.5)
                    if celda == 3:
                        cell.width = Inches(0.25)
                    if celda == 4:
                        cell.width = Inches(1)
                    celda += 1

                hdr_cells = table.rows[0].cells

                hdr_cells[0].text = 'DESCRIPCION\nSpecification'
                hdr_cells[1].text = 'CANTIDAD\nQty'
                hdr_cells[2].text = 'PRECIO EUR\nPrice'
                hdr_cells[3].text = 'DTO\nDis'
                hdr_cells[4].text = 'IMPORTE\nAmount'

                with open('csvofertas/oferta.csv') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    count = 0

                    for row in csv_reader:
                        if count > 2:
                            row_cells = table.add_row().cells
                            row_cells[0].text = 'PLAZO/Delivery: \n' + row[23]
                            if fecha.strip() == str(row[16]).strip():
                                row_cells[1].text = '[STOCK]\n' + row[9]
                            else:
                                row_cells[1].text = row[16].strip() + '\n' + row[10]
                            row_cells[2].text = str(row[18]) + ' ' + str(row[17])
                            row_cells[3].text = row[19]
                            row_cells[4].text = row[20]

                        count += 1

                doc.save("C:/ofertas/oferta.docx")

                '''inputFile = "C:/ofertas/oferta.docx"
                outputFile = "C:/ofertas/oferta.pdf"

                convert(inputFile, outputFile)'''

                return HttpResponse('Recibido correctamente')
            else:
                return HttpResponse('Error')
