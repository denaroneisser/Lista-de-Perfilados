__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-21
"""

import wx
from pubsub import pub
import os
import json
from utils.read_file import readFile
from utils import database, rand


class EditProjeto(wx.Dialog):

    def __init__(self, codigo, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((830, 600))
        self.SetSizeHints(830, 600, 830, 600)
        self.SetTitle("Editar projeto")
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))
        self.SetBackgroundColour(wx.Colour(236, 236, 236))

        self.codigo = codigo
        self.files_path = []
        self.desenhos = []
        self.titles = []
        self.title_values = []
        self.qtds = []
        self.qtd_values = []

        self.db = database.Database()
        rows_proj = self.db.select("SELECT * FROM v_projetos\
                                   WHERE codigo=?",
                                   (self.codigo,))

        produto = ''
        local = ''
        desenho = ''
        cliente = ''
        self.os = ''
        if(len(rows_proj) >= 1):
            produto = rows_proj[0][2]
            local = rows_proj[0][3]
            desenho = rows_proj[0][4]
            cliente = rows_proj[0][5]
            self.os = rows_proj[0][6]

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        self.sizer.Add((20, 20), 0, 0, 0)

        gs1 = wx.FlexGridSizer(5, 3, 5, 20)
        self.sizer.Add(gs1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 35)

        st_prod = wx.StaticText(self.panel, wx.ID_ANY, "Produto")
        st_prod.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                                wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_prod, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        st_local = wx.StaticText(self.panel, wx.ID_ANY, "Local(cidade, uf)")
        st_local.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                                 wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_local, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        gs1.Add((0, 0), 0, 0, 0)

        self.txt_prod = wx.TextCtrl(self.panel, wx.ID_ANY, produto)
        self.txt_prod.SetMinSize((200, 23))
        gs1.Add(self.txt_prod, 0, wx.EXPAND, 0)

        self.txt_local = wx.TextCtrl(self.panel, wx.ID_ANY, local)
        self.txt_local.SetMinSize((200, 23))
        gs1.Add(self.txt_local, 0, wx.EXPAND, 0)

        gs1.Add((0, 0), 0, 0, 0)

        gs1.Add((20, 10), 0, 0, 0)

        gs1.Add((0, 0), 0, 0, 0)

        gs1.Add((0, 0), 0, 0, 0)

        st_des = wx.StaticText(self.panel, wx.ID_ANY, "Desenho")
        st_des.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_des, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        st_cli = wx.StaticText(self.panel, wx.ID_ANY, "Cliente")
        st_cli.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_cli, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        st_os = wx.StaticText(self.panel, wx.ID_ANY, "Num. O.S")
        st_os.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                              wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_os, 0, wx.ALIGN_CENTER_VERTICAL, 0)

        self.txt_desenho = wx.TextCtrl(self.panel, wx.ID_ANY, desenho)
        self.txt_desenho.SetMinSize((200, 23))
        gs1.Add(self.txt_desenho, 0, wx.EXPAND, 0)

        self.txt_cliente = wx.TextCtrl(self.panel, wx.ID_ANY, cliente)
        self.txt_cliente.SetMinSize((200, 23))
        gs1.Add(self.txt_cliente, 0, wx.EXPAND, 0)

        self.txt_num_os = wx.TextCtrl(self.panel, wx.ID_ANY, self.os)
        self.txt_num_os.Bind(wx.EVT_CHAR, self.handle_keypress)
        gs1.Add(self.txt_num_os, 0, 0, 0)

        self.sizer.Add((20, 20), 0, wx.EXPAND, 0)

        #
        # SCROLLED PANEL
        #
        #

        self.panel1 = wx.ScrolledWindow(self.panel, wx.ID_ANY,
                                        style=wx.TAB_TRAVERSAL)
        self.panel1.SetScrollRate(10, 10)
        self.sizer.Add(self.panel1, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 36)
        #
        #
        #

        self.gs2 = wx.FlexGridSizer(5, 10, 10)

        _st1 = wx.StaticText(self.panel1, wx.ID_ANY, "Desenho")
        _st1.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st1, 0, wx.LEFT, 0)

        _st2 = wx.StaticText(self.panel1, wx.ID_ANY, u"Título")
        _st2.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st2, 1, wx.EXPAND | wx.LEFT, 10)

        _st3 = wx.StaticText(self.panel1, wx.ID_ANY, "Qtd.")
        _st3.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st3, 0, wx.LEFT, 10)

        self.gs2.Add((20, 20), 0, 0, 0)

        self.gs2.Add((10, 20), 0, 0, 0)

        rows = self.db.select("SELECT * FROM v_componentes\
                              WHERE codigo_projeto=?",
                              (self.codigo, ))

        self.codigos = []
        for r in range(len(rows)):
            cod = rows[r][1]
            des = rows[r][3]
            title = rows[r][4]
            qtd = rows[r][5]

            self.codigos.append(cod)
            self.files_path.append("")
            self.desenhos.append(des)

            st_n_des = wx.StaticText(self.panel1, wx.ID_ANY, des)
            self.gs2.Add(st_n_des, 0, wx.ALIGN_CENTER_VERTICAL, 0)

            self.txt_titulo = wx.TextCtrl(self.panel1, wx.ID_ANY, title)
            self.titles.append(self.txt_titulo)
            self.title_values.append(title)
            self.txt_titulo.SetMinSize((-1, 23))
            self.gs2.Add(self.txt_titulo, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | wx.LEFT, 10)

            self.txt_qtd = wx.TextCtrl(self.panel1, wx.ID_ANY, str(qtd))
            self.qtds.append(self.txt_qtd)
            self.qtd_values.append(str(qtd))
            self.txt_qtd.SetMinSize((60, 23))
            self.gs2.Add(self.txt_qtd, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)

            self.btn_del = wx.Button(self.panel1, r, "X",
                                     style=wx.BORDER_NONE)
            self.btn_del.SetMinSize((26, 23))
            self.btn_del.SetBackgroundColour(wx.Colour(204, 50, 50))
            self.btn_del.SetForegroundColour(wx.Colour(255, 255, 255))
            self.btn_del.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT,
                                         wx.FONTSTYLE_NORMAL,
                                         wx.FONTWEIGHT_BOLD, 0, ""))
            self.btn_del.SetToolTip("Excluir")
            self.btn_del.SetCursor(wx.Cursor(wx.CURSOR_HAND))
            self.Bind(wx.EVT_BUTTON,
                      lambda evt, index=r, cod=cod:
                      self.removeItem(evt, index, cod),
                      self.btn_del)
            self.gs2.Add(self.btn_del, 0, 0, 0)

            self.gs2.Add((10, 20), 0, 0, 0)
        #
        #
        #
        #

        self.sizer.Add((20, 10), 0, wx.EXPAND, 0)

        self.gs3 = wx.FlexGridSizer(1, 4, 0, 10)
        self.sizer.Add(self.gs3, 0, wx.ALIGN_RIGHT | wx.RIGHT, 30)

        self.btn_import = wx.Button(self.panel, wx.ID_ANY,
                                    "Importar arquivo(s)",
                                    style=wx.BORDER_NONE)
        self.btn_import.SetBackgroundColour(wx.Colour(135, 161, 158))
        self.btn_import.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_import.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.importFiles, self.btn_import)
        self.gs3.Add(self.btn_import, 0, wx.ALIGN_RIGHT, 0)

        self.gs3.Add((20, 20), 0, 0, 0)

        self.btn_save = wx.Button(self.panel, wx.ID_ANY, "Salvar",
                                  style=wx.BORDER_NONE)
        self.btn_save.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_save.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_save.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnSaveGo, self.btn_save)
        self.gs3.Add(self.btn_save, 0, wx.ALIGN_RIGHT, 0)

        self.btn_cancel = wx.Button(self.panel, wx.ID_ANY, "Cancelar",
                                    style=wx.BORDER_NONE)
        self.btn_cancel.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_cancel.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_cancel.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.btn_cancel)
        self.gs3.Add(self.btn_cancel, 0, wx.ALIGN_RIGHT, 0)

        self.sizer.Add((20, 20), 0, 0, 0)

        self.gs2.AddGrowableCol(1)
        self.panel1.SetSizer(self.gs2)

        gs1.AddGrowableCol(0)
        gs1.AddGrowableCol(1)

        self.panel.SetSizer(self.sizer)

        self.Centre()
        self.Layout()

        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def handle_keypress(self, event):
        keycode = event.GetKeyCode()
        if keycode < 255:
            if keycode == 8 or chr(keycode).isdigit():
                event.Skip()

    def importFiles(self, event):

        frmts = "Todos os formatos(*.*)|*.*|"
        frmts += "Arquivos LPE(*.lpe)|*.lpe|"
        frmts += "Arquivos Excel(*.xl*;*.xlsx;*.xlsm;*.xls)|*.xl*|"
        frmts += "Arquivos de texto(*.txt;*.csv)|*.txt;*.csv"

        with wx.FileDialog(self,
                           "Importar arquivo(s)", '', '',
                           frmts,
                           wx.FD_MULTIPLE | wx.FD_FILE_MUST_EXIST
                           ) as opnFileDlg:

            if opnFileDlg.ShowModal() == wx.ID_CANCEL:
                return

            lst_files_path = opnFileDlg.GetPaths()
            for path in (lst_files_path):
                file_basename = os.path.basename(path)
                fname_wo_ext = file_basename.split('.')[0]
                if fname_wo_ext in (self.desenhos):
                    pass
                else:
                    self.files_path.append(path)
                    self.codigos.append("")

            self.title_values = []
            self.qtd_values = []

            for d in range(len(self.files_path)):
                if d <= len(self.titles) - 1:
                    txttitle = self.titles[d]
                    txtqtd = self.qtds[d]

                    self.title_values.append(txttitle.GetValue())
                    self.qtd_values.append(txtqtd.GetValue())
                else:
                    self.title_values.append("")
                    self.qtd_values.append("1")

            self.refreshItems(event)

    def removeItem(self, event, index, cod):

        dlg = wx.MessageDialog(
            None,
            "Deseja excluir o item {}?".format(index),
            'Aviso de remoção!', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            self.db.delete("DELETE FROM v_componentes\
                           WHERE codigo=? and codigo_projeto=?",
                           (cod, self.codigo), True)

            self.codigos.remove(cod)
            self.files_path.pop(index)
            self.desenhos.pop(index)
            self.titles.pop(index)
            self.qtds.pop(index)

            self.title_values = []
            self.qtd_values = []

            for d in range(len(self.files_path)):
                if d <= len(self.titles) - 1:
                    txttitle = self.titles[d]
                    txtqtd = self.qtds[d]

                    self.title_values.append(txttitle.GetValue())
                    self.qtd_values.append(txtqtd.GetValue())
                else:
                    self.title_values.append("")
                    self.qtd_values.append("1")

            self.refreshItems(event)

        else:
            pass

    def refreshItems(self, event):

        if self.panel1 is None:
            print("panel não existe mais")
            return

        self.newpanel = wx.ScrolledWindow(self.panel, wx.ID_ANY,
                                          style=wx.TAB_TRAVERSAL)
        self.newpanel.SetScrollRate(10, 10)
        self.newpanel.Hide()

        self.gs2 = wx.FlexGridSizer(5, 10, 10)

        _st1 = wx.StaticText(self.newpanel, wx.ID_ANY, "Desenho")
        _st1.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st1, 0, wx.LEFT, 0)

        _st2 = wx.StaticText(self.newpanel, wx.ID_ANY, u"Título")
        _st2.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st2, 1, wx.EXPAND | wx.LEFT, 10)

        _st3 = wx.StaticText(self.newpanel, wx.ID_ANY, "Qtd.")
        _st3.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st3, 0, wx.LEFT, 10)

        self.gs2.Add((20, 20), 0, 0, 0)

        self.gs2.Add((10, 20), 0, 0, 0)

        self.titles = []
        self.qtds = []

        i = -1
        for item in self.files_path:
            i = i + 1

            if(item != ''):
                file_basename = os.path.basename(item)
                fname_wo_ext = file_basename.split('.')[0]
                self.desenhos.append(fname_wo_ext)

            cod = self.codigos[i]
            des = self.desenhos[i]
            title = self.title_values[i]
            qtd = self.qtd_values[i]

            st_n_des = wx.StaticText(self.newpanel, wx.ID_ANY, des)
            self.gs2.Add(st_n_des, 0, wx.ALIGN_CENTER_VERTICAL, 0)

            self.txt_titulo = wx.TextCtrl(self.newpanel, wx.ID_ANY, title)
            self.titles.append(self.txt_titulo)
            self.txt_titulo.SetMinSize((-1, 23))
            self.gs2.Add(self.txt_titulo, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.EXPAND | wx.LEFT, 10)

            self.txt_qtd = wx.TextCtrl(self.newpanel, wx.ID_ANY, str(qtd))
            self.qtds.append(self.txt_qtd)
            self.txt_qtd.SetMinSize((60, 23))
            self.gs2.Add(self.txt_qtd, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 10)

            self.btn_del = wx.Button(self.newpanel, i, "X",
                                     style=wx.BORDER_NONE)
            self.btn_del.SetMinSize((26, 23))
            self.btn_del.SetBackgroundColour(wx.Colour(204, 50, 50))
            self.btn_del.SetForegroundColour(wx.Colour(255, 255, 255))
            self.btn_del.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT,
                                         wx.FONTSTYLE_NORMAL,
                                         wx.FONTWEIGHT_BOLD, 0, ""))
            self.btn_del.SetToolTip("Excluir")
            self.btn_del.SetCursor(wx.Cursor(wx.CURSOR_HAND))
            self.Bind(wx.EVT_BUTTON,
                      lambda evt, index=i, cod=cod:
                      self.removeItem(evt, index, cod),
                      self.btn_del)
            self.gs2.Add(self.btn_del, 0, 0, 0)

            self.gs2.Add((10, 20), 0, 0, 0)
        #
        #
        #
        #

        self.gs2.AddGrowableCol(1)

        self.newpanel.SetSizer(self.gs2)
        self.sizer.Replace(self.panel1, self.newpanel)
        self.panel1.Destroy()
        self.panel1 = self.newpanel
        self.newpanel.Show()
        self.panel.Layout()

    def OnSaveGo(self, event):

        produto = self.txt_prod.GetValue()
        local = self.txt_local.GetValue()
        desenho = self.txt_desenho.GetValue()
        cliente = self.txt_cliente.GetValue()
        num_os = self.txt_num_os.GetValue()

        rows_proj = self.db.select("SELECT * FROM v_projetos\
                                   WHERE num_os=?",
                                   (num_os,))

        if(num_os != self.os and len(rows_proj) >= 1):
            wx.MessageBox('Já existe um projeto com este número de O.S!',
                          'Info', wx.OK)

        elif (produto == ''):
            wx.MessageBox('Digite o nome do produto!', 'Info', wx.OK)

        elif(num_os == ''):
            wx.MessageBox('Digite o número da O.S!', 'Info', wx.OK)
        else:

            k = -1
            for item in self.files_path:
                k = k + 1
                if(item == ''):
                    tuple_comp = (self.titles[k].GetValue(),
                                  self.qtds[k].GetValue(),
                                  self.codigos[k],
                                  self.codigo)
                    self.db.insert("UPDATE v_componentes\
                           SET titulo=?, qtd=?\
                           WHERE codigo=? and codigo_projeto=?", tuple_comp)
                else:
                    uniqid = rand.mt_rand()
                    only_path = []
                    only_path.append(item)
                    arr_obj = readFile(only_path, [])

                    tuple_values = (uniqid, self.codigo,
                                    self.desenhos[k],
                                    self.titles[k].GetValue(),
                                    self.qtds[k].GetValue(),
                                    json.dumps(arr_obj))
                    self.db.insert("INSERT INTO v_componentes\
                                    VALUES (NULL, ?, ?, ?, ?, ?, ?)",
                                   tuple_values)

            tuple_update = (produto, local, desenho,
                            cliente, num_os, self.codigo)
            self.db.insert("UPDATE v_projetos\
                           SET produto=?, local=?,\
                           desenho=?, cliente=?, num_os=?\
                           WHERE codigo=?",
                           tuple_update)

            wx.MessageBox('Projeto atualizado com sucesso!', 'Info', wx.OK)

            pub.sendMessage('refreshList', message='refresh', event=event)
            self.OnClose(event)

    def OnClose(self, event):
        event.Skip()
        self.Destroy()


class EditProjetoInit(wx.App):
    def OnInit(self):
        self.frame = EditProjeto('123456', None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


if __name__ == "__main__":
    app = EditProjetoInit(0)
    app.MainLoop()
