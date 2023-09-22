__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import random


def mt_rand(low=0, high=9):

    num_id = ""
    chars = "0123456789"

    for i in range(1, 9):
        num_id += str(chars[random.randint(0, len(chars) - 1)])
    else:
        pass
    return num_id
