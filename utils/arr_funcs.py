__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-05-17
"""

from utils.read_file import roundDecimal


def multiplyArray(arr=[], qtd_mult=1):

    newarr = []

    for r in range(len(arr)):
        pos = arr[r][0]
        material = arr[r][1]
        dimensao = arr[r][2]
        peso_p_m = arr[r][3]
        peso_calculado = arr[r][4]
        qtd = int(arr[r][5]) * int(qtd_mult)
        qtd = str(qtd)

        peso_parcial = arr[r][6]
        peso_parcial = float(peso_calculado) * int(qtd)
        peso_parcial = roundDecimal(peso_parcial)
        newarr.append([pos, material, dimensao,
                       peso_p_m, peso_calculado,
                       qtd, peso_parcial])

    return newarr
