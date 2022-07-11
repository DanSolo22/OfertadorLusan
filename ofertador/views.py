import docx
from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic.base import View
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

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

                doc = DocxTemplate("csvofertas/plantilla.docx")

                with open('csvofertas/oferta.csv') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    line_count = 0

                    for row in csv_reader:
                        if line_count == 1:
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
                            pro = row[10]
                            tel = row[11]
                            mail = row[12]
                            moneda = row[15]
                            des_moneda = row[16]

                        line_count += 1

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
                        'MAIL': mail,
                        'MONEDA': moneda,
                        'DES_MON': des_moneda,
                    }

                doc.render(context)
                doc.save("C:/ofertas/oferta.docx")

                doc = docx.Document("C:/ofertas/oferta.docx")

                table = doc.add_table(rows=1, cols=6)

                for i in range(6):
                    for cell in table.columns[i].cells:
                        if i == 0:
                            cell.width = Inches(1.25)
                        elif i == 1:
                            cell.width = Inches(2.5)
                        elif i == 2:
                            cell.width = Inches(1)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 3:
                            cell.width = Inches(1)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 4:
                            cell.width = Inches(1)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 5:
                            cell.width = Inches(1)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells = table.rows[0].cells

                '''paragraph = hdr_cells[0].paragraphs[0]
                run = paragraph.runs
                font = run[0].font
                font.size = Pt(30)'''

                hdr_cells[0].paragraphs[0].add_run('REF.\n').font.size = Pt(8)
                hdr_cells[0].paragraphs[0].add_run('REF.\n').font.size = Pt(8)
                hdr_cells[0].paragraphs[0].runs[0].font.bold = True
                hdr_cells[0].paragraphs[0].runs[1].font.italic = True

                hdr_cells[1].paragraphs[0].add_run('DESCRIPCION\n').font.size = Pt(8)
                hdr_cells[1].paragraphs[0].add_run('SPECIFICATION\n').font.size = Pt(8)
                hdr_cells[1].paragraphs[0].runs[0].font.bold = True
                hdr_cells[1].paragraphs[0].runs[1].font.italic = True
                hdr_cells[1].paragraphs[0].runs[1].font.bold = False

                hdr_cells[2].paragraphs[0].add_run('CANTIDAD\n').font.size = Pt(8)
                hdr_cells[2].paragraphs[0].add_run('QTY\n').font.size = Pt(8)
                hdr_cells[2].paragraphs[0].runs[0].font.bold = True
                hdr_cells[2].paragraphs[0].runs[1].font.italic = True
                hdr_cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[3].paragraphs[0].add_run('PRECIO\n').font.size = Pt(8)
                hdr_cells[3].paragraphs[0].add_run('PRICE\n').font.size = Pt(8)
                hdr_cells[3].paragraphs[0].add_run('EUR x 100').font.size = Pt(8)
                hdr_cells[3].paragraphs[0].runs[0].font.bold = True
                hdr_cells[3].paragraphs[0].runs[1].font.italic = True
                hdr_cells[3].paragraphs[0].runs[2].font.bold = True
                hdr_cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[4].paragraphs[0].add_run('DTO\n').font.size = Pt(8)
                hdr_cells[4].paragraphs[0].add_run('DIS\n').font.size = Pt(8)
                hdr_cells[4].paragraphs[0].runs[0].font.bold = True
                hdr_cells[4].paragraphs[0].runs[1].font.italic = True
                hdr_cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[5].paragraphs[0].add_run('IMPORTE\n').font.size = Pt(8)
                hdr_cells[5].paragraphs[0].add_run('AMOUNT\n').font.size = Pt(8)
                hdr_cells[5].paragraphs[0].add_run('EUR').font.size = Pt(8)
                hdr_cells[5].paragraphs[0].runs[0].font.bold = True
                hdr_cells[5].paragraphs[0].runs[1].font.italic = True
                hdr_cells[5].paragraphs[0].runs[2].font.bold = True
                hdr_cells[5].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                with open('csvofertas/oferta.csv') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    count = 0

                    for row in csv_reader:
                        if count > 2:
                            row_cells = table.add_row().cells
                            row_cells[0].paragraphs[0].add_run(row[22]).font.size = Pt(8)
                            row_cells[0].paragraphs[0].add_run('\n' + row[4]).font.size = Pt(8)
                            row_cells[0].paragraphs[0].runs[1].font.italic = True

                            if fecha.strip() == str(row[16]).strip():
                                row_cells[1].paragraphs[0].add_run(row[23]).font.size = Pt(8)
                                row_cells[1].paragraphs[0].add_run('\nPLAZO/Delivery:   [STOCK]').font.size = Pt(8)
                                row_cells[1].paragraphs[0].runs[1].font.bold = True
                            else:
                                row_cells[1].paragraphs[0].add_run(row[23]).font.size = Pt(8)
                                row_cells[1].paragraphs[0].add_run(
                                    '\nPLAZO/Delivery:   ' + row[16].strip()).font.size = Pt(8)
                                row_cells[1].paragraphs[0].runs[1].font.bold = True

                            row_cells[2].text = row[9]
                            row_cells[2].paragraphs[0].runs[0].font.size = Pt(8)
                            row_cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                            row_cells[3].text = str(row[18])
                            row_cells[3].paragraphs[0].runs[0].font.size = Pt(8)
                            row_cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                            row_cells[4].text = row[19]
                            row_cells[4].paragraphs[0].runs[0].font.size = Pt(8)
                            row_cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                            row_cells[5].text = row[20]
                            row_cells[5].paragraphs[0].runs[0].font.size = Pt(8)
                            row_cells[5].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                        count += 1

                doc.save("C:/ofertas/oferta.docx")

                '''inputFile = "C:/ofertas/oferta.docx"
                outputFile = "C:/ofertas/oferta.pdf"

                convert(inputFile, outputFile)'''

                return HttpResponse('Recibido correctamente')
            else:
                return HttpResponse('Error')
