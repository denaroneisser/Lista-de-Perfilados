__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-17
"""


import os
import sqlite3


class Database(object):
    def __init__(self, db_file="database/versatil.db"):
        database_allready_exists = os.path.exists(db_file)
        self.db = sqlite3.connect(db_file)

        if not database_allready_exists:
            pass

    def select(self, sql, values):
        cursor = self.db.cursor()
        if values != '':
            cursor.execute(sql, values)
        else:
            cursor.execute(sql)

        records = cursor.fetchall()
        cursor.close()
        return records

    def insert(self, sql, values):
        newID = 0
        cursor = self.db.cursor()
        cursor.execute(sql, values)
        newID = cursor.lastrowid
        self.db.commit()
        cursor.close()
        return newID

    def save(self, sql):
        cursor = self.db.cursor()
        cursor.execute(sql)
        self.db.commit()
        cursor.close()

    def delete(self, sql, id, multi=False):
        cursor = self.db.cursor()
        if multi:
            cursor.execute(sql, id)
        else:
            cursor.execute(sql, (id,))

        self.db.commit()
        cursor.close()
