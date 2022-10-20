from docx.oxml.ns import qn
from docx.oxml.shared import OxmlElement


def set_repeat_table_header(row):
    """ set repeat table row on every new page
    """
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    tbl_header = OxmlElement('w:tblHeader')
    tbl_header.set(qn('w:val'), "true")
    tr_pr.append(tbl_header)
    return row


def insert_hr(paragraph):
    p = paragraph._p
    p_pr = p.get_or_add_pPr()
    p_bdr = OxmlElement('w:pBdr')
    p_pr.insert_element_before(
        p_bdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    p_bdr.append(bottom)


def comprovar_plazo(fecha):
    array_fecha = fecha.split("/")

    if array_fecha[1] != '00' and array_fecha[2] != '0000':
        return fecha
    elif array_fecha[1] == '00' and array_fecha[2] == '0000':
        return str(array_fecha[0]) + ' dias'
    elif array_fecha[0] == '00' and array_fecha[2] == '0000':
        if array_fecha[1] == '01' or array_fecha[1] == '1':
            return str(array_fecha[1]) + ' mes'
        else:
            return str(array_fecha[1]) + ' meses'
    elif array_fecha[0] != '00' and array_fecha[1] != '00' and array_fecha[2] == '0000':
        if array_fecha[1] == '01' or array_fecha[1] == '1':
            return str(array_fecha[0]) + ' dias y ' + str(array_fecha[1]) + ' mes'
        else:
            return str(array_fecha[0]) + ' dias y ' + str(array_fecha[1]) + ' meses'
    elif array_fecha[0] == '00' and array_fecha[1] == '00' and array_fecha[2] != '0000':
        return str(array_fecha[2])


def comprovar_stock(fecha_pedido, fecha_plazo):
    if str(fecha_pedido) == str(fecha_plazo):
        return True
    else:
        array_fecha_pedido = str(fecha_pedido).split('/')
        array_fecha_plazo = str(fecha_plazo).split('/')

        if int(str(array_fecha_pedido[0]).strip()) == int(str(array_fecha_plazo[0]).strip()) and int(
                str(array_fecha_pedido[1]).strip()) == int(str(array_fecha_plazo[1]).strip()) and int(
            str(array_fecha_pedido[2]).strip()) == int(str(array_fecha_plazo[2]).strip()):
            return True
        else:
            return False


def comprovar_usuario(usuario):
    match usuario:
        case 'lai':
            return 'Laia'
        case 'jor':
            return 'Jordi'
        case 'car':
            return 'Carmen'
        case 'jos':
            return 'José'
        case 'ant':
            return 'Antonio'
        case 'lou':
            return 'Lourdes'
        case 'ang':
            return 'Angeles'
        case 'luz':
            return 'Mari Luz'
        case 'jua':
            return 'Juan Manuel'
        case _:
            return 'Lusan'