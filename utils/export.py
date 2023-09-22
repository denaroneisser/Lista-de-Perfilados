__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import wx
from xlsxwriter import Workbook
from utils.read_file import listByPos, roundDecimal
from utils import database
import datetime
import json

d = datetime.datetime.now()


def gerarPlanilha(fname='', cod_proj='',
                  lbl_lst=[], group_list=[], lst_insp=[]):

    db = database.Database()

    len_lb = len(lbl_lst)
    nome_produto = lbl_lst[0] if (len_lb >= 1) else ''
    local = lbl_lst[1] if (len_lb >= 2) else ''
    desenho = lbl_lst[2] if (len_lb >= 3) else ''
    cliente = lbl_lst[3] if (len_lb >= 4) else ''
    num_os = lbl_lst[4] if (len_lb >= 5) else ''

    dia = '%02d' % d.day
    mes = '%02d' % d.month
    ano = d.year
    data_atual = '{}/{}/{}'.format(dia, mes, ano)

    try:
        with open(fname, "a+") as filee:
            filee.close()
    except IOError:
        dlg = wx.MessageDialog(
            None,
            'Não foi possível salvar!\n Feche o excel!',
            'Erro', wx.OK | wx.ICON_ERROR)
        result = dlg.ShowModal()
        if(result == wx.ID_NO):
            pass
        return False

    if(fname != ''):
        wb = Workbook(fname)
    else:
        wb = Workbook('planilha_versatil.xlsx')

    ws_capa = wb.add_worksheet('Capa')
    ws_capa.set_margins(0.62, 0.3, 0.3, 0.7)

    ws_capa.set_column(0, 0, 20)
    ws_capa.set_column(1, 1, 21)
    ws_capa.set_column(2, 2, 2)
    ws_capa.set_column(3, 3, 16)
    ws_capa.set_column(4, 4, 9)
    ws_capa.set_column(5, 5, 19)

    FMT_TITLE = wb.add_format({
        'font_name': 'Arial',
        'font_size': 25,
        'align': 'center',
        'valign': 'vcenter'
    })

    FMT_LIST = wb.add_format({
        'bold': 1,
        'font_name': 'Arial',
        'font_size': 32,
        'align': 'center',
        'valign': 'vcenter'
    })

    ws_capa.merge_range(17, 0, 17, 5, f'{nome_produto.upper()}', FMT_TITLE)
    ws_capa.merge_range(18, 0, 18, 5, 'LISTA DE MATERIAIS', FMT_LIST)

    #
    #
    #
    fmt_img = wb.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#000000'
    })

    ws_capa.set_row(40, 30)
    ws_capa.set_row(41, 35)
    ws_capa.set_row(42, 35)
    ws_capa.set_row(43, 25)
    ws_capa.set_row(44, 25)

    fmt_merge = wb.add_format({
        'top': 1,
        'bottom': 1,
        'border_color': '#000000'
    })

    ws_capa.merge_range(40, 2, 42, 2, '', fmt_merge)

    ws_capa.merge_range(40, 0, 42, 1, '', fmt_img)
    ws_capa.insert_image(40, 0, 'assets/engefame.png',
                         {'x_offset': 10,
                          'y_offset': 25,
                          'object_position': 1})

    fmt_title_client = wb.add_format({
        'bold': 1,
        'font_size': 12,
        'font_name': 'Arial',
        'valign': 'bottom',
        'top': 1,
        'right': 1,
        'border_color': '#000000'
    })

    ws_capa.merge_range(40, 3, 40, 5, 'CLIENTE:', fmt_title_client)

    fmt_client = wb.add_format({
        'font_size': 11,
        'font_name': 'Arial',
        'text_wrap': 1,
        'valign': 'top',
        'bottom': 1,
        'right': 1,
        'border_color': '#000000'
    })

    ws_capa.merge_range(41, 3, 42, 5, f'{cliente.upper()}', fmt_client)

    fmt_title = wb.add_format({
        'bold': 1,
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'border_color': '#000000'
    })

    ws_capa.write(43, 0, 'DATA', fmt_title)
    ws_capa.write(43, 1, 'O.S', fmt_title)

    ws_capa.merge_range(43, 2, 43, 3, u'Nº DO DESENHO', fmt_title)

    ws_capa.write(43, 4, 'REV.', fmt_title)
    ws_capa.write(43, 5, 'TOT. DE FOLHAS', fmt_title)

    fmt_data = wb.add_format({
        'font_size': 11,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': '@',
        'border': 1,
        'border_color': '#000000'
    })

    ws_capa.write(44, 0, f'{data_atual}', fmt_data)
    ws_capa.write(44, 1, f'{num_os}', fmt_data)

    ws_capa.merge_range(44, 2, 44, 3, f'{desenho}', fmt_data)

    ws_capa.write(44, 4, '', fmt_data)
    ws_capa.write(44, 5, '', fmt_data)

    #
    #
    #
    #
    #

    ws_ccapa = wb.add_worksheet('Contra capa')
    ws_ccapa.set_margins(0.8, 0.3, 0.3, 0.7)

    ws_ccapa.set_column(0, 0, 14)
    ws_ccapa.set_column(1, 1, 27)
    ws_ccapa.set_column(2, 2, 6)
    ws_ccapa.set_column(3, 3, 12)
    ws_ccapa.set_column(4, 4, 8)
    ws_ccapa.set_column(5, 5, 13)

    merge_format = wb.add_format({
        'bold': 1,
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter'
    })

    ws_ccapa.merge_range('A1:F1',
                         'ENGEFAME ESTRUTURAS METÁLICAS LTDA',
                         merge_format)
    ws_ccapa.merge_range('A2:F2', 'TODOS OS TÍTULOS', merge_format)

    ws_ccapa.write(3, 0, f'Produto.....: {nome_produto}')
    ws_ccapa.write(4, 0, f'Local/UF....: {local}')
    ws_ccapa.write(5, 0, f'Desenho....: {desenho}')
    ws_ccapa.write(6, 0, f'Cliente.......: {cliente}')

    ws_ccapa.set_header('&C                                                  \
                                                &P', {
        'margin': 0.93
    })

    ws_ccapa.write(3, 4, 'Pág.:')
    ws_ccapa.write(5, 4, f'O.S.: {num_os}')
    ws_ccapa.write(6, 4, f'Data: {data_atual}')

    cols_ccapa = ['Desenho', 'Título', 'Qtd', 'Peso(kg)',
                  'Galv.(%)', 'Peso tot.(kg)']

    titles_format = wb.add_format({
        'bold': 1,
        'font_size': 11,
        'border': 7,
        'border_color': '#1A1A1A'
    })

    fmt_cc = wb.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'num_format': '0.00',
        'border': 7,
        'border_color': '#1A1A1A'
    })

    colc = 0
    for item in (cols_ccapa):
        ws_ccapa.write(8, colc, item, titles_format)
        colc += 1

    rowsc = db.select("SELECT * FROM v_componentes\
                      WHERE codigo_projeto=?",
                      (cod_proj, ))

    cc_rng_start = 9
    for ccomp in range(len(rowsc)):

        nome_desenho = rowsc[ccomp][3]
        titulo = rowsc[ccomp][4]
        qtd = rowsc[ccomp][5]
        data = rowsc[ccomp][6]
        res = json.loads(data)
        cc_mat_len = len(res)

        smpt_cc = 0
        arr_peso_cc = []
        for m_cc in range(cc_mat_len):
            peso_parcial = res[m_cc][6]
            arr_peso_cc.append(peso_parcial)

        if(len(arr_peso_cc) >= 1):
            smpt_cc = sum(map(float, arr_peso_cc))
            smpt_cc = roundDecimal(smpt_cc)

        ws_ccapa.write(cc_rng_start, 0, nome_desenho, fmt_cc)
        ws_ccapa.write(cc_rng_start, 1, titulo, fmt_cc)
        ws_ccapa.write(cc_rng_start, 2, str(qtd), fmt_cc)
        ws_ccapa.write(cc_rng_start, 3, smpt_cc, fmt_cc)
        ws_ccapa.write(cc_rng_start, 4, '3.5', fmt_cc)

        total_galv_cc = smpt_cc + (smpt_cc * (3.5 / 100))
        total_galv_cc = roundDecimal(total_galv_cc)

        ws_ccapa.write(cc_rng_start, 5, total_galv_cc, fmt_cc)

        cc_rng_start = cc_rng_start + 1

    #
    #
    #
    #
    #

    ws_corte = wb.add_worksheet('Corte')
    left = 0.7
    right = 0.7
    top = 0.7
    bottom = 0.5

    ws_corte.set_margins(left, right, top, bottom)
    ws_corte.repeat_rows(0, 8)

    ws_corte.set_column(0, 0, 7)
    ws_corte.set_column(1, 1, 16)
    ws_corte.set_column(2, 2, 11)
    ws_corte.set_column(3, 3, 11)
    ws_corte.set_column(4, 4, 10)
    ws_corte.set_column(5, 5, 10)
    ws_corte.set_column(6, 6, 6)
    ws_corte.set_column(7, 7, 10)

    ws_corte.merge_range('A1:H1',
                         'ENGEFAME ESTRUTURAS METÁLICAS LTDA',
                         merge_format)
    ws_corte.merge_range('A2:H2', 'LISTA DE MATERIAIS - CORTE', merge_format)

    ws_corte.write(3, 0, f'Produto.....: {nome_produto}')
    ws_corte.write(4, 0, f'Local/UF....: {local}')
    ws_corte.write(5, 0, f'Desenho....: {desenho}')
    ws_corte.write(6, 0, f'Cliente.......: {cliente}')

    ws_corte.set_header('&C                                                   \
                                         &P', {
        'margin': 1.32
    })

    ws_corte.write(3, 5, 'Pág.:')
    ws_corte.write(5, 5, f'O.S.: {num_os}')
    ws_corte.write(6, 5, f'Data: {data_atual}')

    cols_corte = ['Pos', 'Material', 'Código',
                  'Dimensão', 'Peso(kg/m)', 'Peso un.',
                  'Qtd', 'Peso tot.']

    col = 0
    for item in (cols_corte):
        ws_corte.write(8, col, item, titles_format)
        col += 1

    fmt_corte = wb.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'num_format': '0.00',
        'top': 7,
        'bottom': 7,
        'left': 7,
        'right': 7,
        'border_color': '#E2EBEA'
    })

    fmt_hg = wb.add_format({
        'bold': 1,
        'font_size': 11,
        'align': 'left',
        'num_format': '0.00',
        'bg_color': '#E2EBEA'
    })

    data_fmt1 = wb.add_format({
        'num_format': '0.00',
        'bg_color': '#E2EBEA'
    })

    arr_len = len(group_list)
    rng_start = 9

    for a in range(arr_len):
        it_rows = group_list[a][1]
        lst_mat = group_list[a][2]
        peso_total = group_list[a][3]
        peso_total = roundDecimal(peso_total)

        ap_rows = it_rows + 2

        m = -1
        for row in range(rng_start, rng_start + ap_rows):
            m = m + 1

            if(m < len(lst_mat)):
                for col in range(8):
                    if(col == 2):
                        ws_corte.write(row, col, '', fmt_corte)
                    elif(col > 2):
                        newcol = col - 1
                        ws_corte.write(row, col, lst_mat[m][newcol], fmt_corte)
                    else:
                        ws_corte.write(row, col, lst_mat[m][col], fmt_corte)

            if(row == rng_start + ap_rows - 2):
                ws_corte.write(row, 0, '', fmt_hg)
                ws_corte.write(row, 4, 'PESO TOTAL=', fmt_hg)
                ws_corte.write(row, 7, peso_total, fmt_hg)
                ws_corte.set_row(row, cell_format=data_fmt1)

        rng_start += ap_rows

    #
    #
    #
    #
    #

    ws_insp = wb.add_worksheet('Inspeção')

    ws_insp.set_margins(1.2, right, top, bottom)
    ws_insp.repeat_rows(0, 9)

    ws_insp.set_column(0, 0, 7)
    ws_insp.set_column(1, 1, 17)
    ws_insp.set_column(2, 2, 14)
    ws_insp.set_column(3, 3, 8)
    ws_insp.set_column(4, 4, 13)
    ws_insp.set_column(5, 5, 13)

    ws_insp.merge_range('A1:F1',
                        'ENGEFAME ESTRUTURAS METÁLICAS LTDA',
                        merge_format)
    ws_insp.merge_range('A2:F2', 'LISTA DE MATERIAIS - INSPEÇÃO', merge_format)

    ws_insp.write(3, 0, f'Produto.....: {nome_produto}')
    ws_insp.write(4, 0, f'Local/UF....: {local}')
    ws_insp.write(5, 0, f'Desenho....: {desenho}')
    ws_insp.write(6, 0, f'Cliente.......: {cliente}')

    ws_insp.set_header('&C                                                   \
            &P', {
        'margin': 1.32
    })

    ws_insp.write(3, 4, 'Pág.:')
    ws_insp.write(5, 4, f'O.S.: {num_os}')
    ws_insp.write(6, 4, f'Data: {data_atual}')

    cols_insp = ['Pos', 'Material', 'Dimensão', 'Qtd', '1', '2']

    merge_format2 = wb.add_format({
        'bold': 1,
        'font_size': 11,
        'border': 7,
        'border_color': '#1A1A1A',
        'align': 'center',
        'valign': 'vcenter'
    })

    for i in range(7):
        subtitle = ''
        if(i == 4):
            subtitle = 'Inspeção'
            ws_insp.merge_range(8, i, 8, i + 1, subtitle,
                                merge_format2)
            ws_insp.write(9, i, cols_insp[i], merge_format2)
        elif(i == 5):
            ws_insp.write(8, i, subtitle)
            ws_insp.write(9, i, cols_insp[i], merge_format2)
        else:
            ws_insp.write(8, i, subtitle)

        if(i <= 3):
            ws_insp.write(9, i, cols_insp[i], titles_format)

    fmt_insp = wb.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'num_format': '0.00',
        'border': 7,
        'border_color': '#1A1A1A'
    })

    lst_insp_len = len(lst_insp)
    rng_insp = 10
    m = -1
    for a in range(rng_insp, rng_insp + lst_insp_len):
        m = m + 1
        it_qtd = lst_insp[m][5]

        for i2 in range(6):
            if(i2 == 3):
                ws_insp.write(a, i2, it_qtd, fmt_insp)
            elif(i2 == 4 or i2 == 5):
                ws_insp.write(a, i2, '', merge_format2)
            else:
                it_mat = lst_insp[m][i2]
                ws_insp.write(a, i2, it_mat, fmt_insp)

    #
    #
    #
    #
    #

    ws_res = wb.add_worksheet('Resumo')

    ws_res.set_margins(0.8, right, top, bottom)
    ws_res.repeat_rows(0, 8)

    ws_res.set_column(0, 0, 19)
    ws_res.set_column(1, 1, 16)
    ws_res.set_column(2, 2, 14)
    ws_res.set_column(3, 3, 10)
    ws_res.set_column(4, 4, 10)
    ws_res.set_column(5, 5, 10)

    ws_res.merge_range('A1:F1',
                       'ENGEFAME ESTRUTURAS METÁLICAS LTDA',
                       merge_format)
    ws_res.merge_range('A2:F2', 'LISTA DE MATERIAIS - RESUMO', merge_format)

    ws_res.write(3, 0, f'Produto.....: {nome_produto}')
    ws_res.write(4, 0, f'Local/UF....: {local}')
    ws_res.write(5, 0, f'Desenho....: {desenho}')
    ws_res.write(6, 0, f'Cliente.......: {cliente}')

    ws_res.set_header('&C                                                   \
                                                           &P', {
        'margin': 1.32
    })

    ws_res.write(3, 4, 'Pág.:')
    ws_res.write(5, 4, f'O.S.: {num_os}')
    ws_res.write(6, 4, f'Data: {data_atual}')

    cols_res = ['Perfil', 'Peso(kg)', 'Obs.:']

    merge_format3 = wb.add_format({
        'bold': 1,
        'font_size': 11,
        'border': 7,
        'border_color': '#1A1A1A',
        'align': 'left',
        'valign': 'vcenter'
    })

    for i in range(7):
        if(i <= 2):
            if(i == 2):
                ws_res.merge_range(8, i, 8, i + 3, cols_res[i],
                                   merge_format3)
            else:
                ws_res.write(8, i, cols_res[i], merge_format3)

        elif(i == 6):
            ws_res.write(8, i, '')
        else:
            ws_res.write(8, i, '')

    fmt_res = wb.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'num_format': '0.00',
        'border': 7,
        'border_color': '#1A1A1A'
    })

    rng_res = 9
    arr_peso = []
    smpt = 0
    al = -1
    for row in range(rng_res, rng_res + arr_len):
        al = al + 1
        perfil = group_list[al][0]
        peso_total = group_list[al][3]
        peso_total = roundDecimal(peso_total)
        arr_peso.append(peso_total)

        for i3 in range(3):
            if(i3 == 0):
                ws_res.write(row, i3, perfil, fmt_res)
            elif(i3 == 1):
                ws_res.write(row, i3, peso_total, fmt_res)
            else:
                ws_res.merge_range(row, i3, row, i3 + 3, '',
                                   fmt_res)

    if(len(arr_peso) >= 1):
        smpt = sum(map(float, arr_peso))
        smpt = roundDecimal(smpt)

    ws_res.write(rng_res + arr_len, 0, 'TOTAL=', fmt_hg)
    ws_res.write(rng_res + arr_len, 1, smpt, fmt_hg)
    ws_res.set_row(rng_res + arr_len, cell_format=data_fmt1)

    #
    #
    #
    #
    #

    ws_lista = wb.add_worksheet('Lista de materiais')
    ws_lista.set_margins(0.8, 0.3, top, 0.7)

    ws_lista.set_column(0, 0, 7)   # pos
    ws_lista.set_column(1, 1, 17)  # denomina
    ws_lista.set_column(2, 2, 12)  # dimensao
    ws_lista.set_column(3, 3, 13)  # peso
    ws_lista.set_column(4, 4, 11)  # peso unit
    ws_lista.set_column(5, 5, 7)   # qtd
    ws_lista.set_column(6, 6, 12)  # peso parcial

    fmt_lista = wb.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'align': 'left',
        'num_format': '0.00',
        'border': 7,
        'border_color': '#1A1A1A'
    })

    data_fmtg1 = wb.add_format({
        'num_format': '0.00',
        'bg_color': '#F9FBFC'
    })

    cols_lista = ['Pos', 'Denominação', 'Dimensão',
                  'Peso(kg/m)', 'Peso un.', 'Qtd', 'Peso parc.']

    rows = db.select("SELECT * FROM v_componentes\
                     WHERE codigo_projeto=?",
                     (cod_proj, ))
    arr_mat = []
    rpb = 0  # range page break
    total_linhas = 0
    for components in range(len(rows)):

        nome_desenho = rows[components][3]
        titulo = rows[components][4]
        qtd = rows[components][5]
        data = rows[components][6]
        res = json.loads(data)
        arr_mat = listByPos(res)
        lst_mat_len = len(arr_mat)

        ws_lista.merge_range(rpb, 0, rpb, 6,
                             'ENGEFAME ESTRUTURAS METÁLICAS LTDA',
                             merge_format)
        ws_lista.merge_range(rpb + 1, 0, rpb + 1, 6,
                             'LISTA DE MATERIAIS', merge_format)

        ws_lista.write(rpb + 3, 0, f'Desenho...: {nome_desenho}')
        ws_lista.write(rpb + 4, 0, f'Título........: {titulo}')
        ws_lista.write(rpb + 5, 0, f'Quant.......: {qtd}')

        for lst in range(7):
            ws_lista.write(rpb + 7, lst, cols_lista[lst], merge_format3)

        # ######### SOMA TOTAL DE LINHAS
        total_linhas = total_linhas + 8

        rng_list = rpb + 8
        ml = -1
        arr_peso_lst = []
        smpt_lst = 0
        for al in range(rng_list, rng_list + lst_mat_len):
            total_linhas = total_linhas + 1
            ml = ml + 1
            peso_parcial = arr_mat[ml][6]
            arr_peso_lst.append(peso_parcial)

            for m2 in range(7):
                it_mat = arr_mat[ml][m2]
                if(m2 == 3 or m2 == 4 or m2 == 6):
                    it_mat = str(it_mat).replace(".", ",")
                ws_lista.write(al, m2, str(it_mat), fmt_lista)

        if(len(arr_peso_lst) >= 1):
            smpt_lst = sum(map(float, arr_peso_lst))
            smpt_lst = roundDecimal(smpt_lst)

        ws_lista.write(rng_list + lst_mat_len, 4, 'TOTAL PARCIAL=', fmt_hg)
        ws_lista.write(rng_list + lst_mat_len, 6, smpt_lst, fmt_hg)
        ws_lista.set_row(rng_list + lst_mat_len, cell_format=data_fmt1)

        ws_lista.write(rng_list + lst_mat_len + 1, 4, 'GALVANIZAÇÃO')
        ws_lista.write(rng_list + lst_mat_len + 1, 6, '3,5%')
        ws_lista.set_row(rng_list + lst_mat_len + 1, cell_format=data_fmtg1)

        total_galv = smpt_lst + (smpt_lst * (3.5 / 100))
        total_galv = roundDecimal(total_galv)

        ws_lista.write(rng_list + lst_mat_len + 2, 4, 'PESO TOTAL=', fmt_hg)
        ws_lista.write(rng_list + lst_mat_len + 2, 6, total_galv, fmt_hg)
        ws_lista.set_row(rng_list + lst_mat_len + 2, cell_format=data_fmt1)

        total_linhas = total_linhas + 3

        fator_catalizador = 0
        if components < (len(rows) - 1):
            if total_linhas == 50:
                total_linhas = 50
            elif total_linhas > 50:
                total_linhas = total_linhas - 50

            fator_catalizador = 50 - total_linhas

        total_linhas = 0
        rpb = rpb + 8 + lst_mat_len + 3 + fator_catalizador

    try:
        wb.close()
    except Exception:
        pass


#
#
#
#
#
#


def gerarPDF(fname='', lbl_lst=[], group_list=[], lst_insp=[]):
    pass

#     layout = '<div>ENGEFAME ESTRUTURAS METÁLICAS LTDA</div>'
#     '<div>LISTA DE MATERIAIS - CORTE</div>'
#     '<p>Aqui vai um teste de parágrafo para um'
#     ' documento PDF gerado pelo Python</p>'
#     print("gerar PDF:", fname)
