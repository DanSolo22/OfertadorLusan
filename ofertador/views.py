import docx
from django.http import HttpResponse
from django.shortcuts import render
from django.views.generic.base import View
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.shared import Inches, Pt, Cm
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
                mail = ''
                moneda = ''
                des_moneda = ''
                peso = ''
                transporte = ''
                contacto = ''
                num_contacto = ''
                importe_bruto = ''
                portes = ''
                dtopp = ''
                imp_dtopp = ''
                base_imp = ''
                iva = ''
                imp_iva = ''
                rec_quiv = ''
                imp_rec_quiv = ''
                total = ''
                forma_pago = ''
                iban = 'ES25 2100-1083-1102-0005-4013'

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
                            mail = row[13]
                            moneda = row[15]
                            des_moneda = row[16]
                            peso = row[17] + ' kg.'
                            contacto = row[33]
                            num_contacto = row[34]
                            importe_bruto = row[20]
                            portes = row[19]
                            dtopp = row[22]
                            imp_dtopp = row[23]
                            base_imp = row[26]
                            iva = row[27]
                            imp_iva = row[28]
                            rec_quiv = row[29]
                            imp_rec_quiv = row[30]
                            total = row[31]
                            forma_pago = row[49]
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
                            cell.width = Inches(1)
                        elif i == 1:
                            cell.width = Inches(4)
                        elif i == 2:
                            cell.width = Inches(0.8)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 3:
                            cell.width = Inches(0.8)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 4:
                            cell.width = Inches(0.8)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif i == 5:
                            cell.width = Inches(0.8)
                            cell.paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr = table.rows[0]
                hdr_cells = hdr.cells

                hdr_cells[0].paragraphs[0].add_run('REF.\n').font.size = Pt(9)
                hdr_cells[0].paragraphs[0].add_run('REF.\n').font.size = Pt(9)
                hdr_cells[0].paragraphs[0].runs[0].font.bold = True
                hdr_cells[0].paragraphs[0].runs[1].font.italic = True

                hdr_cells[1].paragraphs[0].add_run('DESCRIPCION\n').font.size = Pt(9)
                hdr_cells[1].paragraphs[0].add_run('SPECIFICATION\n').font.size = Pt(9)
                hdr_cells[1].paragraphs[0].runs[0].font.bold = True
                hdr_cells[1].paragraphs[0].runs[1].font.italic = True
                hdr_cells[1].paragraphs[0].runs[1].font.bold = False

                hdr_cells[2].paragraphs[0].add_run('CANTIDAD\n').font.size = Pt(9)
                hdr_cells[2].paragraphs[0].add_run('QUANTITY\n').font.size = Pt(9)
                hdr_cells[2].paragraphs[0].runs[0].font.bold = True
                hdr_cells[2].paragraphs[0].runs[1].font.italic = True
                hdr_cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[3].paragraphs[0].add_run('PRECIO\n').font.size = Pt(9)
                hdr_cells[3].paragraphs[0].add_run('PRICE\n').font.size = Pt(9)
                hdr_cells[3].paragraphs[0].add_run('EUROx100').font.size = Pt(9)
                hdr_cells[3].paragraphs[0].runs[0].font.bold = True
                hdr_cells[3].paragraphs[0].runs[1].font.italic = True
                hdr_cells[3].paragraphs[0].runs[2].font.bold = True
                hdr_cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[4].paragraphs[0].add_run('DTO.\n').font.size = Pt(9)
                hdr_cells[4].paragraphs[0].add_run('DIS.\n').font.size = Pt(9)
                hdr_cells[4].paragraphs[0].add_run('%').font.size = Pt(9)
                hdr_cells[4].paragraphs[0].runs[0].font.bold = True
                hdr_cells[4].paragraphs[0].runs[1].font.italic = True
                hdr_cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr_cells[5].paragraphs[0].add_run('IMPORTE\n').font.size = Pt(9)
                hdr_cells[5].paragraphs[0].add_run('AMOUNT\n').font.size = Pt(9)
                hdr_cells[5].paragraphs[0].add_run('EURO').font.size = Pt(9)
                hdr_cells[5].paragraphs[0].runs[0].font.bold = True
                hdr_cells[5].paragraphs[0].runs[1].font.italic = True
                hdr_cells[5].paragraphs[0].runs[2].font.bold = True
                hdr_cells[5].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                hdr.height = Cm(1.25)
                hdr.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                barra_cabeza = table.add_row()
                barra_cabeza_tabla = barra_cabeza.cells

                barra_cabeza_tabla[5].merge(barra_cabeza_tabla[4])
                barra_cabeza_tabla[4].merge(barra_cabeza_tabla[3])
                barra_cabeza_tabla[3].merge(barra_cabeza_tabla[2])
                barra_cabeza_tabla[2].merge(barra_cabeza_tabla[1])
                barra_cabeza_tabla[1].merge(barra_cabeza_tabla[0])

                barra_cabeza.height = Cm(0.5)
                barra_cabeza.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                barra_cabeza_tabla[0].text = '--------------------------------------------------------------------------------------------------------------------------------------------------------'

                with open('csvofertas/oferta.csv') as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    count = 0

                    for row in csv_reader:
                        if count > 2:
                            row_cells = table.add_row().cells
                            row_cells[0].paragraphs[0].add_run(row[22]).font.size = Pt(9)
                            row_cells[0].paragraphs[0].add_run('\n' + row[4]).font.size = Pt(9)
                            row_cells[0].paragraphs[0].runs[1].font.italic = True

                            if fecha.strip() == str(row[16]).strip():
                                row_cells[1].paragraphs[0].add_run(row[23]).font.size = Pt(9)
                                row_cells[1].paragraphs[0].add_run('\nPLAZO/Delivery:   [STOCK]').font.size = Pt(9)
                                row_cells[1].paragraphs[0].runs[1].font.bold = True
                                row_cells[1].paragraphs[0].runs[1].font.italic = True
                            else:
                                row_cells[1].paragraphs[0].add_run(row[23]).font.size = Pt(9)
                                row_cells[1].paragraphs[0].add_run(
                                    '\nPLAZO/Delivery:   ' + row[16].strip()).font.size = Pt(9)
                                row_cells[1].paragraphs[0].runs[1].font.bold = True

                            row_cells[2].text = row[9]
                            row_cells[2].paragraphs[0].runs[0].font.size = Pt(9)
                            row_cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                            row_cells[3].text = row[18]
                            row_cells[3].paragraphs[0].runs[0].font.size = Pt(9)
                            row_cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                            if str(row[19]).strip() == '':
                                row_cells[4].text = 'NETO'
                            else:
                                row_cells[4].text = row[19]

                            row_cells[4].paragraphs[0].runs[0].font.size = Pt(9)
                            row_cells[4].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                            row_cells[5].text = row[20]
                            row_cells[5].paragraphs[0].runs[0].font.size = Pt(9)
                            row_cells[5].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                        count += 1

                barra_pie = table.add_row()
                barra_pie_tabla = barra_pie.cells

                barra_pie_tabla[5].merge(barra_pie_tabla[4])
                barra_pie_tabla[4].merge(barra_pie_tabla[3])
                barra_pie_tabla[3].merge(barra_pie_tabla[2])
                barra_pie_tabla[2].merge(barra_pie_tabla[1])
                barra_pie_tabla[1].merge(barra_pie_tabla[0])

                barra_pie.height = Cm(0.5)
                barra_pie.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                barra_pie_tabla[0].text = '--------------------------------------------------------------------------------------------------------------------------------------------------------'

                pie_tabla = table.add_row().cells

                pie_tabla[5].merge(pie_tabla[4])
                pie_tabla[4].merge(pie_tabla[3])
                pie_tabla[3].merge(pie_tabla[2])
                pie_tabla[0].merge(pie_tabla[1])

                pie_tabla[0].text = 'PRECIOS VÁLIDOS PARA LAS CANTIDADES OFERTADAS'
                pie_tabla[0].paragraphs[0].runs[0].font.size = Pt(8)
                pie_tabla[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                pie_tabla[0].paragraphs[0].runs[0].font.bold = True

                pie_tabla[2].text = 'OFERTA VIGENTE HASTA LA FECHA: ' + validez
                pie_tabla[2].paragraphs[0].runs[0].font.size = Pt(8)
                pie_tabla[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                pie_tabla[2].paragraphs[0].runs[0].font.bold = True

                doc.add_paragraph()

                table_resumen = doc.add_table(rows=12, cols=6)

                for i in range(5):
                    for cell in table_resumen.columns[i].cells:
                        if i == 0:
                            cell.width = Inches(1.5)
                        if i == 1:
                            cell.width = Inches(1.5)
                        if i == 2:
                            cell.width = Inches(2.5)
                        if i == 3:
                            cell.width = Inches(2.5)
                        if i == 4:
                            cell.width = Inches(0.5)
                        if i == 4:
                            cell.width = Inches(1.25)

                for row in table_resumen.rows:
                    row.height = Cm(0.35)
                    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                table_resumen.cell(0, 0).paragraphs[0].add_run('PESO').font.size = Pt(8)
                table_resumen.cell(0, 0).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(0, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(1, 0).paragraphs[0].add_run('WEIGHT').font.size = Pt(8)
                table_resumen.cell(1, 0).paragraphs[0].runs[0].font.italic = True
                table_resumen.cell(1, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(3, 0).paragraphs[0].add_run('TRANSPORTE').font.size = Pt(8)
                table_resumen.cell(3, 0).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(3, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(4, 0).paragraphs[0].add_run('TRANSPORT').font.size = Pt(8)
                table_resumen.cell(4, 0).paragraphs[0].runs[0].font.italic = True
                table_resumen.cell(4, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(6, 0).paragraphs[0].add_run('CONTACTO').font.size = Pt(8)
                table_resumen.cell(6, 0).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(6, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(7, 0).paragraphs[0].add_run('CONTACT').font.size = Pt(8)
                table_resumen.cell(7, 0).paragraphs[0].runs[0].font.italic = True
                table_resumen.cell(7, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(0, 1).paragraphs[0].text = peso
                table_resumen.cell(0, 1).paragraphs[0].runs[0].font.size = Pt(8)

                table_resumen.cell(3, 1).paragraphs[0].text = transporte
                table_resumen.cell(3, 1).paragraphs[0].runs[0].font.size = Pt(8)

                table_resumen.cell(6, 1).paragraphs[0].text = contacto
                table_resumen.cell(6, 1).paragraphs[0].runs[0].font.size = Pt(8)

                table_resumen.cell(7, 1).paragraphs[0].text = num_contacto
                table_resumen.cell(7, 1).paragraphs[0].runs[0].font.size = Pt(8)

                table_resumen.cell(0, 3).paragraphs[0].add_run('IMPORTE BRUTO / ').font.size = Pt(8)
                table_resumen.cell(0, 3).paragraphs[0].add_run('GROSS AMOUNT').font.size = Pt(8)
                table_resumen.cell(0, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(0, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(0, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(0, 5).paragraphs[0].add_run(importe_bruto + ' €').font.size = Pt(8)
                table_resumen.cell(0, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(1, 3).paragraphs[0].add_run('PORTES / ').font.size = Pt(8)
                table_resumen.cell(1, 3).paragraphs[0].add_run('TRANSPORT').font.size = Pt(8)
                table_resumen.cell(1, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(1, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(1, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(1, 5).paragraphs[0].add_run(portes + ' €').font.size = Pt(8)
                table_resumen.cell(1, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(2, 3).paragraphs[0].add_run('DTOP. PP.').font.size = Pt(8)
                table_resumen.cell(2, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(2, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(2, 4).paragraphs[0].add_run(dtopp + ' %').font.size = Pt(8)
                table_resumen.cell(2, 4).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(2, 5).paragraphs[0].add_run(imp_dtopp + ' €').font.size = Pt(8)
                table_resumen.cell(2, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(3, 3).paragraphs[0].add_run('BASE IMPONIBLE / ').font.size = Pt(8)
                table_resumen.cell(3, 3).paragraphs[0].add_run('TAXABLE BASE').font.size = Pt(8)
                table_resumen.cell(3, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(3, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(3, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(3, 5).paragraphs[0].add_run(base_imp + ' €').font.size = Pt(8)
                table_resumen.cell(3, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(4, 3).paragraphs[0].add_run('IVA / ').font.size = Pt(8)
                table_resumen.cell(4, 3).paragraphs[0].add_run('IVA').font.size = Pt(8)
                table_resumen.cell(4, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(4, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(4, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(4, 4).paragraphs[0].add_run(iva + ' %').font.size = Pt(8)
                table_resumen.cell(4, 4).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(4, 5).paragraphs[0].add_run(imp_iva + ' €').font.size = Pt(8)
                table_resumen.cell(4, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(5, 3).paragraphs[0].add_run('REC. EQUIVALENCIA').font.size = Pt(8)
                table_resumen.cell(5, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(5, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(5, 4).paragraphs[0].add_run(rec_quiv + ' %').font.size = Pt(8)
                table_resumen.cell(5, 4).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(5, 5).paragraphs[0].add_run(imp_rec_quiv + ' €').font.size = Pt(8)
                table_resumen.cell(5, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(6, 3).merge(table_resumen.cell(6, 4))
                table_resumen.cell(6, 4).merge(table_resumen.cell(6, 5))
                table_resumen.cell(6, 3).paragraphs[0].add_run('-----------------------------------------------------------------------------------------------------').font.size = Pt(8)
                table_resumen.cell(6, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(6, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                table_resumen.cell(7, 3).paragraphs[0].add_run('IMPORTE TOTAL / ').font.size = Pt(9)
                table_resumen.cell(7, 3).paragraphs[0].add_run('TOTAL AMOUNT').font.size = Pt(9)
                table_resumen.cell(7, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(7, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(7, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(7, 5).paragraphs[0].add_run(total + ' €').font.size = Pt(9)
                table_resumen.cell(7, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                table_resumen.cell(10, 3).paragraphs[0].add_run('FORMA DE PAGO / ').font.size = Pt(8)
                table_resumen.cell(10, 3).paragraphs[0].add_run('MEANS OF PAYMENT').font.size = Pt(8)
                table_resumen.cell(10, 3).paragraphs[0].runs[0].font.bold = True
                table_resumen.cell(10, 3).paragraphs[0].runs[1].font.italic = True
                table_resumen.cell(10, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                table_resumen.cell(10, 5).paragraphs[0].add_run(forma_pago).font.size = Pt(8)
                table_resumen.cell(10, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                if str(forma_pago).strip() == 'TRANSFERENCIA' or str(forma_pago).strip() == 'CONTADO':
                    table_resumen.cell(11, 3).paragraphs[0].add_run('SWIFT/IBAN').font.size = Pt(9)
                    table_resumen.cell(11, 3).paragraphs[0].runs[0].font.bold = True
                    table_resumen.cell(11, 3).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                    table_resumen.cell(11, 5).paragraphs[0].add_run(iban).font.size = Pt(9)
                    table_resumen.cell(11, 5).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

                doc.save("C:/ofertas/oferta.docx")

                '''inputFile = "C:/ofertas/oferta.docx"
                outputFile = "C:/ofertas/oferta.pdf"

                convert(inputFile, outputFile)'''

                return HttpResponse('Recibido correctamente')
            else:
                return HttpResponse('Error')
