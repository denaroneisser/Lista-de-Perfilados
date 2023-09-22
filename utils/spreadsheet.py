__author__ = 'FenanxCorp'
__version__ = '0.0.1'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-22
"""
from xlwt import add_palette_colour, Workbook, easyxf
from decimal import Decimal, ROUND_HALF_DOWN
import datetime

d = datetime.datetime.now()


def gerarPlanilha(fname, lbl_lst, group_list, list_insp):

    nome_produto = lbl_lst[0]
    local = lbl_lst[1]
    desenho = lbl_lst[2]
    cliente = lbl_lst[3]
    num_os = lbl_lst[4]

    dia = '%02d' % d.day
    mes = '%02d' % d.month
    ano = d.year
    data_atual = '{}/{}/{}'.format(dia, mes, ano)

    wb = Workbook(encoding='utf-8')

    ##
    #
    # CORTE
    #
    #

    add_palette_colour("row_greencolor", 0x21)
    wb.set_colour_RGB(0x21, 234, 247, 245)

    ws_corte = wb.add_sheet('Corte')
    ws_corte.col(0).width = 1800
    ws_corte.col(1).width = 5400
    ws_corte.col(2).width = 3000
    ws_corte.col(3).width = 2800
    ws_corte.col(4).width = 2800
    ws_corte.col(5).width = 2000
    ws_corte.col(6).width = 4350

    ws_corte.row(2).height_mismatch = True
    ws_corte.row(2).height = 200
    ws_corte.row(7).height_mismatch = True
    ws_corte.row(7).height = 400

    ws_corte.write(0, 0, 'VERSATIL ENG. E PROJETOS LTDA', easyxf(
        'font: name Courier New, bold True;borders: top 1, left 1;'))

    ws_corte.write(0, 1, '', easyxf('borders: top 1;'))
    ws_corte.write(0, 2, '', easyxf('borders: top 1, right 1;'))
    ws_corte.write(0, 3, '', easyxf('borders: top 1;'))
    ws_corte.write(0, 4, '', easyxf('borders: top 1;'))
    ws_corte.write(0, 5, '', easyxf('borders: top 1;'))

    ws_corte.write(0, 6, 'LISTA DE MATERIAIS PARA CORTE', easyxf(
        'font: name Courier New, bold True;'
        'align: horiz right;borders: top 1, right 1;'))
    ws_corte.write(1, 0, 'ENGEFAME LTDA', easyxf(
        'font: name Courier New, bold True;'
        'borders: left 1, bottom 1;'))

    # style border
    for i in range(1, 7):
        if(i == 2 or i == 6):
            ws_corte.write(1, i, '', easyxf('borders: right 1, bottom 1;'))
        else:
            ws_corte.write(1, i, '', easyxf('borders: bottom 1;'))

    stylfn = easyxf('font: name Courier New;')

    styl = easyxf(
        'font: name Courier New;'
        'borders: left 1;')

    ws_corte.write(2, 0, '', styl)
    ws_corte.write(3, 0, f'Produto....: {nome_produto}', styl)
    ws_corte.write(4, 0, f'Local / UF.: {local}', styl)
    ws_corte.write(5, 0, f'Desenho....: {desenho}', styl)
    ws_corte.write(5, 3, f'OS....: {num_os}', stylfn)

    ws_corte.write(2, 6, '', easyxf('borders: right 1;'))
    ws_corte.write(3, 6, '', easyxf('borders: right 1;'))
    ws_corte.write(4, 6, '', easyxf('borders: right 1;'))
    ws_corte.write(5, 6, '', easyxf('borders: right 1;'))

    ws_corte.write(6, 0, f'Cliente....: {cliente}', easyxf(
        'font: name Courier New;'
        'borders: left 1, bottom 1;'))

    ws_corte.write(6, 1, '', easyxf('borders: bottom 1;'))
    ws_corte.write(6, 2, '', easyxf('borders: bottom 1;'))

    ws_corte.write(6, 3, f'Data..: {data_atual}', easyxf(
        'font: name Courier New;'
        'borders: bottom 1;'))

    ws_corte.write(6, 4, '', easyxf('borders: bottom 1;'))
    ws_corte.write(6, 5, '', easyxf('borders: bottom 1;'))
    ws_corte.write(6, 6, '', easyxf('borders: right 1, bottom 1;'))

    cols_corte = ['Pos', 'Material', 'Dimensão', 'Peso(kg/m)',
                  'Peso unit.', 'Qtd', 'Peso parcial(kg)']

    for i in range(len(cols_corte)):
        style_easy = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: left 1, right 1, bottom 1;')
        ws_corte.write(7, i, cols_corte[i], style_easy)

    #
    #
    # GERA LISTA DE MATERIAIS DE CORTE
    #

    arr_len = len(group_list)

    rng_ct = 8
    for a in range(arr_len):
        it_rows = group_list[a][1]
        lst_mat = group_list[a][2]
        peso_total = group_list[a][3]
        peso_total = roundDecimal(peso_total)

        ap_rows = it_rows + 2

        m = -1
        for l in range(rng_ct, rng_ct + ap_rows):
            m = m + 1
            row = l
            e_style = easyxf(
                'font: bold True;'
                'pattern: pattern solid, fore_colour row_greencolor')

            if(m < len(lst_mat)):
                for col in range(7):
                    it_material = lst_mat[m][col]
                    if(col == 3):
                        it_material = it_material.replace(".", ",")
                    elif(col == 4 or col == 6):
                        it_material = str(it_material).replace(".", ",")

                    ws_corte.write(row, col, it_material, easyxf())

            if(row == rng_ct + ap_rows - 2):
                ws_corte.write(row, 4, 'PESO TOTAL=', e_style)

                str_outpeso = str(peso_total).replace(".", ",")
                ws_corte.write(row, 6, str_outpeso, e_style)
                ws_corte.row(row).set_style(e_style)

        rng_ct += ap_rows

    ##
    #
    #
    #
    # INSPEÇÃO
    #
    #

    ws_insp = wb.add_sheet('Inspeção')

    ws_insp.col(0).width = 1800
    ws_insp.col(1).width = 6400
    ws_insp.col(2).width = 4000

    ws_insp.col(3).width = 2000
    ws_insp.col(4).width = 2000
    ws_insp.col(5).width = 2000
    ws_insp.col(6).width = 2000

    ws_insp.row(2).height_mismatch = True
    ws_insp.row(2).height = 200

    ws_insp.row(8).height_mismatch = True
    ws_insp.row(8).height = 400

    ws_insp.write(0, 0, 'VERSATIL ENG. E PROJETOS LTDA', easyxf(
        'font: name Courier New, bold True;borders: top 1, left 1;'))

    ws_insp.write(0, 1, '', easyxf('borders: top 1;'))
    ws_insp.write(0, 2, '', easyxf('borders: top 1, right 1;'))
    ws_insp.write(0, 3, '', easyxf('borders: top 1;'))
    ws_insp.write(0, 4, '', easyxf('borders: top 1;'))

    ws_insp.write(0, 5, 'LISTA DE INSPEÇÃO', easyxf(
        'font: name Courier New, bold True;'
        'align: horiz center;borders: top 1;'))

    ws_insp.write(0, 6, '', easyxf('borders: top 1, right 1;'))

    ws_insp.write(1, 0, 'ENGEFAME LTDA', easyxf(
        'font: name Courier New, bold True;'
        'borders: left 1, bottom 1;'))

    # style border
    for i in range(1, 7):
        if(i == 2 or i == 6):
            ws_insp.write(1, i, '', easyxf('borders: right 1, bottom 1;'))
        else:
            ws_insp.write(1, i, '', easyxf('borders: bottom 1;'))

    ws_insp.write(2, 0, '', styl)
    ws_insp.write(3, 0, f'Produto....: {nome_produto}', styl)
    ws_insp.write(4, 0, f'Local / UF.: {local}', styl)
    ws_insp.write(5, 0, f'Desenho....: {desenho}', styl)
    ws_insp.write(5, 3, f'OS....: {num_os}', stylfn)

    ws_insp.write(2, 6, '', easyxf('borders: right 1;'))
    ws_insp.write(3, 6, '', easyxf('borders: right 1;'))
    ws_insp.write(4, 6, '', easyxf('borders: right 1;'))
    ws_insp.write(5, 6, '', easyxf('borders: right 1;'))

    ws_insp.write(6, 0, f'Cliente....: {cliente}', easyxf(
        'font: name Courier New;'
        'borders: left 1, bottom 1;'))

    ws_insp.write(6, 1, '', easyxf('borders: bottom 1;'))
    ws_insp.write(6, 2, '', easyxf('borders: bottom 1;'))

    ws_insp.write(6, 3, f'Data..: {data_atual}', easyxf(
        'font: name Courier New;'
        'borders: bottom 1;'))

    ws_insp.write(6, 4, '', easyxf('borders: bottom 1;'))
    ws_insp.write(6, 5, '', easyxf('borders: bottom 1;'))
    ws_insp.write(6, 6, '', easyxf('borders: right 1, bottom 1;'))

    cols_insp = ['Pos', 'Material', 'Dimensão', 'Qtd', '1', '2']

    for i in range(7):
        style_easy = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: left 1, right 1;')

        style_easyno = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: bottom 1;')

        style_easynoh = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center, horiz right;'
            'borders: bottom 1;')

        style_easyr = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: right 1, bottom 1;')

        style_easyb = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: left 1, right 1, bottom 1;')
        subtitle = ''
        if(i == 4):
            ws_insp.write(7, i, subtitle, style_easyno)
            ws_insp.write(8, i, cols_insp[i], style_easynoh)
        elif(i == 5):
            subtitle = 'Inspeção'
            ws_insp.write(7, i, subtitle, style_easyno)
            ws_insp.write(8, i, '', style_easyno)
        elif(i == 6):
            ws_insp.write(7, i, subtitle, style_easyb)
            ws_insp.write(8, i, cols_insp[i - 1], style_easyr)
        else:
            ws_insp.write(7, i, subtitle, style_easy)

        if(i <= 3):
            ws_insp.write(8, i, cols_insp[i], style_easyb)

    #
    #
    # GERA LISTA DE INSPEÇÃO
    #

    lst_insp_len = len(list_insp)

    rng_insp = 9
    m = -1
    for a in range(rng_insp, rng_insp + lst_insp_len):
        m = m + 1
        it_qtd = list_insp[m][5]

        for i2 in range(5):
            if(i2 == 3):
                ws_insp.write(a, i2, it_qtd, easyxf())
            elif(i2 == 4):
                ws_insp.write(a, i2, '', easyxf())
            else:
                it_mat = list_insp[m][i2]
                ws_insp.write(a, i2, it_mat, easyxf())

    ##
    #
    #
    #
    # RESUMO
    #
    #

    ws_res = wb.add_sheet('Resumo')

    ws_res.col(0).width = 6000
    ws_res.col(1).width = 3800

    ws_res.col(2).width = 2300
    ws_res.col(3).width = 2300
    ws_res.col(4).width = 2300
    ws_res.col(5).width = 2300
    ws_res.col(6).width = 2300

    ws_res.row(2).height_mismatch = True
    ws_res.row(2).height = 200

    ws_res.row(7).height_mismatch = True
    ws_res.row(7).height = 400

    ws_res.write(0, 0, 'VERSATIL ENG. E PROJETOS LTDA', easyxf(
        'font: name Courier New, bold True;borders: top 1, left 1;'))

    ws_res.write(0, 1, '', easyxf('borders: top 1, right 1;'))
    ws_res.write(0, 2, '', easyxf('borders: top 1;'))
    ws_res.write(0, 3, '', easyxf('borders: top 1;'))
    ws_res.write(0, 4, 'REQUISIÇÃO AO ALMOXARIFADO', easyxf(
        'font: name Courier New, bold True;'
        'align: horiz center;borders: top 1;'))

    ws_res.write(0, 5, '', easyxf('borders: top 1;'))
    ws_res.write(0, 6, '', easyxf('borders: top 1, right 1;'))

    ws_res.write(1, 0, 'ENGEFAME LTDA', easyxf(
        'font: name Courier New, bold True;'
        'borders: left 1, bottom 1;'))

    # style border
    for i in range(1, 7):
        if(i == 1 or i == 6):
            ws_res.write(1, i, '', easyxf('borders: right 1, bottom 1;'))
        elif(i == 4):
            ws_res.write(1, i, 'RESUMO DE LAMINADOS', easyxf(
                'font: name Courier New, bold True;'
                'align: horiz center;borders: bottom 1;'))
        else:
            ws_res.write(1, i, '', easyxf('borders: bottom 1;'))

    ws_res.write(2, 0, '', styl)
    ws_res.write(3, 0, f'Produto....: {nome_produto}', styl)
    ws_res.write(4, 0, f'Local / UF.: {local}', styl)
    ws_res.write(5, 0, f'Desenho....: {desenho}', styl)
    ws_res.write(5, 3, f'OS....: {num_os}', stylfn)

    ws_res.write(2, 6, '', easyxf('borders: right 1;'))
    ws_res.write(3, 6, '', easyxf('borders: right 1;'))
    ws_res.write(4, 6, '', easyxf('borders: right 1;'))
    ws_res.write(5, 6, '', easyxf('borders: right 1;'))

    ws_res.write(6, 0, f'Cliente....: {cliente}', easyxf(
        'font: name Courier New;'
        'borders: left 1, bottom 1;'))

    ws_res.write(6, 1, '', easyxf('borders: bottom 1;'))
    ws_res.write(6, 2, '', easyxf('borders: bottom 1;'))

    ws_res.write(6, 3, f'Data..: {data_atual}', easyxf(
        'font: name Courier New;'
        'borders: bottom 1;'))

    ws_res.write(6, 4, '', easyxf('borders: bottom 1;'))
    ws_res.write(6, 5, '', easyxf('borders: bottom 1;'))
    ws_res.write(6, 6, '', easyxf('borders: right 1, bottom 1;'))

    cols_res = ['Perfil', 'Peso(kg)', 'Obs.:']

    for i in range(7):
        style_easy = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: bottom 1;')

        style_easy_rl = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: left 1, right 1, bottom 1;')

        style_easy_l = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: left 1, bottom 1;')

        style_easy_r = easyxf(
            'font: name Courier New, height 180;'
            'align: vert center;'
            'borders: right 1, bottom 1;')

        if(i <= 2):
            if(i == 2):
                ws_res.write(7, i, cols_res[i], style_easy_l)
            else:
                ws_res.write(7, i, cols_res[i], style_easy_rl)

        elif(i == 6):
            ws_res.write(7, i, '', style_easy_r)
        else:
            ws_res.write(7, i, '', style_easy)

    #
    #
    # GERA LISTA DE RESUMO
    #
    #

    rng_res = 8
    arr_peso = []
    smpt = 0
    al = -1
    for row in range(rng_res, rng_res + arr_len):
        al = al + 1
        perfil = group_list[al][0]
        peso_total = group_list[al][3]
        arr_peso.append(peso_total)

        for col in range(3):
            if(col == 0):
                ws_res.write(row, col, perfil, easyxf())
            elif(col == 1):
                peso_total = roundDecimal(peso_total)
                ptr = str(peso_total).replace(".", ",")
                ws_res.write(row, col, ptr, easyxf())
            else:
                ws_res.write(row, col, '', easyxf())

    if(len(arr_peso) >= 1):
        smpt = sum(map(float, arr_peso))
        smpt = roundDecimal(smpt)

    e_style = easyxf(
        'font: bold True;'
        'pattern: pattern solid, fore_colour row_greencolor')

    ws_res.write(rng_res + arr_len, 0, 'TOTAL', e_style)
    ws_res.write(rng_res + arr_len, 1, str(smpt).replace(".", ","), e_style)
    ws_res.row(rng_res + arr_len).set_style(e_style)

    if(fname != ''):
        wb.save(fname)
    else:
        wb.save('planilha_versatil.xls')


def roundDecimal(num):
    ptd = Decimal(num)
    out = Decimal(ptd.quantize(Decimal('.01'),
                               rounding=ROUND_HALF_DOWN))
    return out
