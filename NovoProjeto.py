__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import wx
from pubsub import pub
import os
import json
from utils.read_file import readFile
from utils import database, rand
import ListProjetos


class NovoProjeto(wx.Dialog):

    def __init__(self, files_p, org, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((1000, 550))
        self.SetSizeHints(1000, 550, 1000, 550)
        self.SetTitle("Novo projeto")
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))
        self.SetBackgroundColour(wx.Colour(236, 236, 236))

        self.files_path = files_p
        self.org = org
        self.desenhos = []
        self.titles = []
        self.title_values = []
        self.qtds = []
        self.qtd_values = []

        self.panel = wx.Panel(self, wx.ID_ANY)

        self.sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.sizer.Add((30, 20), 0, 0, 0)

        sizer1 = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(sizer1, 0, wx.EXPAND | wx.RIGHT, 40)

        gs1 = wx.GridSizer(10, 1, 0, 3)
        sizer1.Add(gs1, 1, wx.EXPAND | wx.TOP, 20)

        st_prod = wx.StaticText(self.panel, wx.ID_ANY, "Produto")
        st_prod.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                                wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_prod, 0, 0, 0)

        self.txt_prod = wx.TextCtrl(self.panel, wx.ID_ANY, "")
        gs1.Add(self.txt_prod, 0, wx.EXPAND, 0)

        st_loc = wx.StaticText(self.panel, wx.ID_ANY, "Local (cidade, uf)")
        st_loc.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_loc, 0, wx.TOP, 5)

        self.txt_loc = wx.TextCtrl(self.panel, wx.ID_ANY, "")
        gs1.Add(self.txt_loc, 0, wx.EXPAND, 0)

        st_des = wx.StaticText(self.panel, wx.ID_ANY, "Desenho")
        st_des.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_des, 0, wx.TOP, 5)

        self.txt_des = wx.TextCtrl(self.panel, wx.ID_ANY, "")
        gs1.Add(self.txt_des, 0, wx.EXPAND, 0)

        st_cli = wx.StaticText(self.panel, wx.ID_ANY, "Cliente")
        st_cli.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                               wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_cli, 0, wx.TOP, 5)

        self.txt_cli = wx.TextCtrl(self.panel, wx.ID_ANY, "")
        gs1.Add(self.txt_cli, 0, wx.EXPAND, 0)

        st_os = wx.StaticText(self.panel, wx.ID_ANY, "Num. O.S")
        st_os.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                              wx.FONTWEIGHT_BOLD, 0, ""))
        gs1.Add(st_os, 0, wx.TOP, 5)

        self.txt_os = wx.TextCtrl(self.panel, wx.ID_ANY, "")
        self.txt_os.SetMinSize((120, 23))
        self.txt_os.Bind(wx.EVT_CHAR, self.handle_keypress)
        gs1.Add(self.txt_os, 0, 0, 0)

        self.btn_import_files = wx.Button(self.panel, wx.ID_ANY,
                                          "Importar arquivo(s)",
                                          style=wx.BORDER_NONE)
        self.btn_import_files.SetBackgroundColour(wx.Colour(135, 161, 158))
        self.btn_import_files.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_import_files.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.importFiles, self.btn_import_files)
        sizer1.Add(self.btn_import_files, 0, wx.TOP, 30)

        sizer1.Add((20, 20), 1, wx.EXPAND | wx.RIGHT, 50)

        gs3 = wx.GridSizer(1, 2, 0, 10)
        sizer1.Add(gs3, 0, wx.BOTTOM | wx.EXPAND, 20)

        self.btn_save = wx.Button(self.panel, wx.ID_ANY, "Salvar",
                                  style=wx.BORDER_NONE)
        self.btn_save.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_save.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_save.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnSaveGo, self.btn_save)
        gs3.Add(self.btn_save, 0, 0, 0)

        self.btn_cancel = wx.Button(self.panel, wx.ID_ANY, "Cancelar",
                                    style=wx.BORDER_NONE)
        self.btn_cancel.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_cancel.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_cancel.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.btn_cancel)
        gs3.Add(self.btn_cancel, 0, 0, 0)

        self.panel1 = wx.ScrolledWindow(self.panel, wx.ID_ANY,
                                        style=wx.TAB_TRAVERSAL)
        self.panel1.SetScrollRate(10, 10)
        self.sizer.Add(self.panel1, 1, wx.EXPAND | wx.RIGHT, 0)

        self.gs2 = wx.FlexGridSizer(5, 5, 10)

        _st1 = wx.StaticText(self.panel1, wx.ID_ANY, "Desenho")
        _st1.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st1, 0, wx.TOP, 20)

        _st2 = wx.StaticText(self.panel1, wx.ID_ANY, u"Título")
        _st2.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st2, 1, wx.EXPAND | wx.TOP, 20)

        _st3 = wx.StaticText(self.panel1, wx.ID_ANY, "Qtd.")
        _st3.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st3, 0, wx.TOP, 20)

        self.gs2.Add((20, 20), 0, 0, 0)

        self.gs2.Add((10, 20), 0, 0, 0)

        i = -1
        for item in self.files_path:
            i = i + 1
            file_basename = os.path.basename(item)
            fname_wo_ext = file_basename.split('.')[0]
            self.desenhos.append(fname_wo_ext)

            titulo = ''
            qtde = '1'

            st_desenho = wx.StaticText(self.panel1, wx.ID_ANY,
                                       fname_wo_ext)
            self.gs2.Add(st_desenho, 0, wx.ALIGN_CENTER_VERTICAL, 0)

            self.txt_titulo = wx.TextCtrl(self.panel1, wx.ID_ANY, titulo)
            self.titles.append(self.txt_titulo)
            self.title_values.append(titulo)
            self.gs2.Add(self.txt_titulo, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 0)

            self.txt_qtd = wx.TextCtrl(self.panel1, wx.ID_ANY, str(qtde))
            self.qtds.append(self.txt_qtd)
            self.qtd_values.append(str(qtde))
            self.txt_qtd.SetMinSize((60, 23))
            self.txt_qtd.Bind(wx.EVT_CHAR, self.handle_keypress)
            self.gs2.Add(self.txt_qtd, 0, wx.ALIGN_CENTER_VERTICAL, 0)

            self.btn_del = wx.Button(self.panel1, i, "X",
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
                      lambda evt, index=i: self.removeItem(evt, index),
                      self.btn_del)
            self.gs2.Add(self.btn_del, 0, 0, 0)

            self.gs2.Add((0, 0), 0, 0, 0)

        self.gs2.AddGrowableCol(1)
        self.panel1.SetSizer(self.gs2)

        self.panel.SetSizer(self.sizer)
        self.Centre()
        self.Layout()

        self.Bind(wx.EVT_CLOSE, self.OnClose)

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

    def removeItem(self, event, index):
        dlg = wx.MessageDialog(
            None,
            "Deseja excluir o item {}?".format(index),
            'Aviso de remoção!', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            self.files_path.pop(index)
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

        self.gs2 = wx.FlexGridSizer(5, 5, 10)

        _st1 = wx.StaticText(self.newpanel, wx.ID_ANY, "Desenho")
        _st1.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st1, 0, wx.TOP, 20)

        _st2 = wx.StaticText(self.newpanel, wx.ID_ANY, u"Título")
        _st2.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st2, 1, wx.EXPAND | wx.TOP, 20)

        _st3 = wx.StaticText(self.newpanel, wx.ID_ANY, "Qtd.")
        _st3.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                             wx.FONTWEIGHT_BOLD, 0, ""))
        self.gs2.Add(_st3, 0, wx.TOP, 20)

        self.gs2.Add((20, 20), 0, 0, 0)

        self.gs2.Add((10, 20), 0, 0, 0)

        #
        #
        self.desenhos = []
        self.titles = []
        self.qtds = []

        i = -1
        for item in self.files_path:
            i = i + 1
            file_basename = os.path.basename(item)
            fname_wo_ext = file_basename.split('.')[0]
            self.desenhos.append(fname_wo_ext)

            titulo = ''
            qtde = '1'

            if(len(self.title_values) >= len(self.files_path)):
                titulo = self.title_values[i]
                qtde = self.qtd_values[i]

            st_desenho = wx.StaticText(self.newpanel, wx.ID_ANY,
                                       fname_wo_ext)
            self.gs2.Add(st_desenho, 0, wx.ALIGN_CENTER_VERTICAL, 0)

            self.txt_titulo = wx.TextCtrl(self.newpanel, wx.ID_ANY, titulo)
            self.titles.append(self.txt_titulo)
            self.gs2.Add(self.txt_titulo, 0,
                         wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 0)

            self.txt_qtd = wx.TextCtrl(self.newpanel, wx.ID_ANY, str(qtde))
            self.qtds.append(self.txt_qtd)
            self.txt_qtd.SetMinSize((60, 23))
            self.txt_qtd.Bind(wx.EVT_CHAR, self.handle_keypress)
            self.gs2.Add(self.txt_qtd, 0, wx.ALIGN_CENTER_VERTICAL, 0)

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
                      lambda evt, index=i: self.removeItem(evt, index),
                      self.btn_del)
            self.gs2.Add(self.btn_del, 0, 0, 0)

            self.gs2.Add((0, 0), 0, 0, 0)
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

    def handle_keypress(self, event):
        keycode = event.GetKeyCode()
        if keycode < 255:
            if keycode == 8 or chr(keycode).isdigit():
                event.Skip()

    def OnSaveGo(self, event):
        lst_desenhos = []
        self.db = database.Database()

        uniqid = rand.mt_rand()
        produto = self.txt_prod.GetValue()
        local = self.txt_loc.GetValue()
        desenho = self.txt_des.GetValue()
        cliente = self.txt_cli.GetValue()
        num_os = self.txt_os.GetValue()

        for d in range(len(self.desenhos)):

            sttxt = self.desenhos[d]
            txttitle = self.titles[d]
            txtqtd = self.qtds[d]

            lst_desenhos.append(
                [sttxt, txttitle.GetValue(), txtqtd.GetValue()])

        tuple_values = (uniqid, produto, local, desenho, cliente, num_os)

        rows_proj = self.db.select("SELECT * FROM v_projetos\
                                   WHERE num_os=?",
                                   (num_os,))

        if(len(rows_proj) >= 1):
            wx.MessageBox('Já existe um projeto com este número de O.S!',
                          'Info', wx.OK)

        elif (produto == ''):
            wx.MessageBox('Digite o nome do produto!', 'Info', wx.OK)

        elif(num_os == ''):
            wx.MessageBox('Digite o número da O.S!', 'Info', wx.OK)

        elif(len(lst_desenhos) == 0):
            wx.MessageBox('Importe pelo menos um arquivo LPE!', 'Info', wx.OK)

        else:
            db = database.Database()
            db.insert("INSERT INTO v_projetos\
                      VALUES (NULL, ?, ?, ?, ?, ?, ?)",
                      tuple_values)

            for ia in range(len(lst_desenhos)):
                unid = rand.mt_rand()
                nome_desenho = lst_desenhos[ia][0]
                titulo_desenho = lst_desenhos[ia][1]
                qtd_desenho = lst_desenhos[ia][2]

                only_path = []
                only_path.append(self.files_path[ia])
                arr_obj = readFile(only_path, [])

                rows = self.db.select("SELECT * FROM v_componentes\
                                      WHERE codigo_projeto=? AND desenho=?",
                                      (uniqid, nome_desenho,))
                if(len(rows) == 0):
                    tuple_values = (unid, uniqid, nome_desenho,
                                    titulo_desenho, qtd_desenho,
                                    json.dumps(arr_obj))
                    self.db.insert("INSERT INTO v_componentes\
                                    VALUES (NULL, ?, ?, ?, ?, ?, ?)",
                                   tuple_values)
                else:
                    tuple_update = (titulo_desenho, qtd_desenho, nome_desenho)
                    self.db.insert("UPDATE v_componentes\
                                   SET titulo=?, qtd=?\
                                   WHERE desenho=?",
                                   tuple_update)

            wx.MessageBox('Projeto salvo com sucesso!', 'Info', wx.OK)

            if(self.org == 'main'):
                frm = ListProjetos.Projetos(None, wx.ID_ANY, "")
                frm.ShowModal()

            pub.sendMessage('refreshList', message='refresh', event=event)
            self.OnClose(event)

    def OnClose(self, event):
        event.Skip()
        self.Destroy()


class NovoInit(wx.App):
    def OnInit(self):
        self.frame = NovoProjeto(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True
