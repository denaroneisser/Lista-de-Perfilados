__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

from decimal import Decimal, ROUND_HALF_UP
from itertools import groupby
import re
import os


def roundDecimal(num, sfloat=False, rounding=ROUND_HALF_UP):
    if(num == ''):
        num = 0

    if sfloat:
        num = float(num)

    ptd = Decimal(num)
    out = Decimal(ptd.quantize(Decimal('.01'),
                               rounding=rounding))
    out = float(out)
    return out


def atoi(elm):
    return int(elm) if elm.isdigit() else elm


def natural_keys(text):
    return [atoi(c) for c in re.split('(\\d+)', text[1])]


def readFile(files_paths, lst_desenhos=0):
    array_objects = []
    k = -1
    for path in files_paths:
        data = open(path, "r")
        qtd_mult = 1
        k = k + 1

        file_basename = os.path.basename(path)
        fn_wo_ext = file_basename.split('.')[0]
        if(len(lst_desenhos) >= 1 and lst_desenhos[k][0] == fn_wo_ext):
            qtd_mult = int(lst_desenhos[k][2])

        for line in data:
            line = line.strip()
            if(line != '\n' and str(line) != ''):

                flds = line.split(";")

                pos = flds[0]
                material = flds[1].strip()
                dimensao = flds[2]
                peso_por_metro = flds[3].replace(",", ".")
                peso_por_metro = roundDecimal(peso_por_metro, True)
                peso_calculado = flds[4].replace(",", ".")
                peso_calculado = roundDecimal(peso_calculado, True)

                qtd = int(flds[5]) * qtd_mult
                qtd = str(qtd)

                peso_parcial = float(peso_calculado) * int(qtd)
                peso_parcial = roundDecimal(peso_parcial)

                array_objects.append(
                                    (pos, material, dimensao,
                                     peso_por_metro, peso_calculado,
                                     qtd, peso_parcial))

                array_objects.sort(key=natural_keys)

            else:
                pass

        data.close()

    return array_objects


def arrSortPos(text):
    return [atoi(c) for c in re.split('(\\d+)', text[0])]


def arrSortMat(elm):
    return elm[1]


def arrSortDimensao(elm):
    # result = float(elm[4]) if (elm[4] != "") else 0
    divide = re.split('x', elm[2])
    result = 0
    if(len(divide) == 1):
        result = int(divide[0])
    else:
        result = int(divide[0]) * int(divide[1])

    return result


def listByPos(arr):
    arr.sort(key=arrSortPos)
    arr_len = len(arr)

    new_arr = []
    it_rep = []
    app_qtd = []
    sum_it_qtd = 0
    for it in range(arr_len):
        pos = arr[it][0]
        pos1 = arr[it + 1][0] if (arr_len > it + 1) else ''
        perf = arr[it][1]
        dim = arr[it][2]
        peso_dim = arr[it][3]
        peso_calc = arr[it][4]
        qtd = arr[it][5]
        peso_tot = arr[it][6]

        if (pos == pos1):

            it_rep.append(pos)
            app_qtd.append(qtd)

        else:

            if pos in (it_rep):
                app_qtd.append(qtd)

                if(len(app_qtd) >= 1):
                    sum_it_qtd = sum(map(int, app_qtd))

            else:
                sum_it_qtd = qtd

            app_qtd = []
            arr_app = [pos, perf, dim, peso_dim,
                       peso_calc, str(sum_it_qtd), peso_tot]

            new_arr.append(arr_app)

    return new_arr


def groupByMaterial(arr):

    lst_group = []
    for (i, g) in groupby(arr, key=arrSortMat):
        g_list = list(g)
        g_list.sort(key=arrSortPos)
        lst_len = len(g_list)

        res_merg_list = []
        app_rpt = []
        app_qtd = []
        app_peso = []
        sum_it_qtd = 0
        sum_it_peso = 0
        for un in range(lst_len):
            pos = g_list[un][0]
            pos1 = g_list[un + 1][0] if (lst_len > un + 1) else ''
            perf = g_list[un][1]
            dim = g_list[un][2]
            peso_dim = g_list[un][3]
            peso_calc = g_list[un][4]
            qtd = g_list[un][5]
            peso_tot = g_list[un][6]

            if (pos == pos1):

                app_rpt.append(pos)
                app_qtd.append(qtd)
                app_peso.append(peso_tot)

            else:

                if pos in (app_rpt):
                    app_qtd.append(qtd)
                    app_peso.append(peso_tot)

                    if(len(app_qtd) >= 1):
                        sum_it_qtd = sum(map(int, app_qtd))

                    if(len(app_peso) >= 1):
                        sum_it_peso = sum(map(float, app_peso))

                else:
                    sum_it_qtd = qtd
                    sum_it_peso = peso_tot

                app_qtd = []
                app_peso = []
                arr_app = [pos,
                           perf, dim, peso_dim,
                           peso_calc, str(sum_it_qtd), sum_it_peso]
                res_merg_list.append(arr_app)

        m_lst_len = len(res_merg_list)

        arr_peso = []
        soma_peso = 0
        for v in range(lst_len):
            arr_peso.append(g_list[v][6])

        if(len(arr_peso) >= 1):
            soma_peso = sum(map(float, arr_peso))

        lst_group.append([i, m_lst_len, res_merg_list, soma_peso])

    return lst_group
