__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-21
"""

import wx
import wx.lib.mixins.listctrl as lm
import NovoProjeto
import ViewProjeto
import EditProjeto
from utils.database import Database
from pubsub import pub


class Projetos(wx.Dialog, lm.ColumnSorterMixin):

    def __init__(self, *args, **kwds):
        kwds["style"] = kwds.get("style", 0) | wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.SetSize((1200, 600))
        self.SetTitle("Listar/selecionar projeto")
        self.SetSizeHints(1200, 600, 1220, 600)
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))
        self.SetBackgroundColour(wx.Colour(236, 236, 236))

        sizer1 = wx.BoxSizer(wx.VERTICAL)

        gs1 = wx.FlexGridSizer(1, 7, 0, 10)
        sizer1.Add(gs1, 0, wx.EXPAND | wx.TOP, 10)

        gs1.Add((10, 20), 0, 0, 0)

        path_ass = 'assets/'

        self.btn_novo = wx.BitmapButton(self, wx.ID_ANY,
                                        wx.Bitmap(path_ass + "new.png",
                                                  wx.BITMAP_TYPE_ANY))
        self.btn_novo.SetSize(self.btn_novo.GetBestSize())
        self.btn_novo.SetToolTip("Novo projeto")
        self.btn_novo.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.novoProjeto, self.btn_novo)
        gs1.Add(self.btn_novo, 0, 0, 0)

        self.btn_edit = wx.BitmapButton(self, wx.ID_ANY,
                                        wx.Bitmap(path_ass + "edit.png",
                                                  wx.BITMAP_TYPE_ANY))
        self.btn_edit.SetSize(self.btn_edit.GetBestSize())
        self.btn_edit.SetToolTip("Editar")
        self.btn_edit.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.btn_edit.Enable(False)
        self.Bind(wx.EVT_BUTTON, self.editProjeto, self.btn_edit)
        gs1.Add(self.btn_edit, 0, 0, 0)

        self.btn_delete = wx.BitmapButton(self, wx.ID_ANY,
                                          wx.Bitmap(path_ass + "delete.png",
                                                    wx.BITMAP_TYPE_ANY))
        self.btn_delete.SetSize(self.btn_delete.GetBestSize())
        self.btn_delete.SetToolTip("Excluir")
        self.btn_delete.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.btn_delete.Enable(False)
        self.Bind(wx.EVT_BUTTON, self.deleteProjeto, self.btn_delete)
        gs1.Add(self.btn_delete, 0, 0, 0)

        self.busca = wx.TextCtrl(self, wx.ID_ANY, "")
        self.busca.SetHint('Buscar projeto por nome ou desenho')
        self.busca.SetMinSize((220, 23))
        gs1.Add(self.busca, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 20)

        gs1.Add((20, 20), 1, wx.EXPAND, 0)

        self.btn_fechar = wx.Button(self, wx.ID_ANY, "Fechar",
                                    style=wx.BORDER_NONE)
        self.btn_fechar.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_fechar.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_fechar.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.btn_fechar)
        gs1.Add(self.btn_fechar, 0, wx.RIGHT, 20)

        sizer2 = wx.BoxSizer(wx.HORIZONTAL)
        sizer1.Add(sizer2, 1, wx.EXPAND | wx.TOP, 10)

        sizer2.Add((20, 20), 0, 0, 0)

        styles_lst = wx.LC_HRULES | wx.LC_REPORT | wx.LC_VRULES
        self.listproject = wx.ListCtrl(self, wx.ID_ANY, style=styles_lst)
        self.listproject.SetBackgroundColour(wx.Colour(236, 236, 236))
        self.listproject.SetForegroundColour(wx.Colour(78, 78, 78))
        self.listproject.AppendColumn(u"Cód.", format=wx.LIST_FORMAT_LEFT,
                                      width=-1)
        self.listproject.AppendColumn("Produto", format=wx.LIST_FORMAT_LEFT,
                                      width=-1)
        self.listproject.AppendColumn("Local (cidade, UF)",
                                      format=wx.LIST_FORMAT_LEFT,
                                      width=wx.LIST_AUTOSIZE_USEHEADER)
        self.listproject.AppendColumn("Desenho", format=wx.LIST_FORMAT_LEFT,
                                      width=-1)
        self.listproject.AppendColumn("Cliente", format=wx.LIST_FORMAT_LEFT,
                                      width=-1)
        self.listproject.AppendColumn("O.S", format=wx.LIST_FORMAT_LEFT,
                                      width=-1)

        db = Database()
        records = db.select(
            """SELECT * FROM v_projetos ORDER BY num_os DESC;""", '')
        count = 0
        self.itemDataMap = dict()

        for v in records:
            codigo = v[1] if v[1] is not None else ''
            produto = v[2] if v[2] is not None else ''
            local = v[3] if v[3] is not None else ''
            desenho = v[4] if v[4] is not None else ''
            cliente = v[5] if v[5] is not None else ''
            os = v[6] if v[6] is not None else ''

            index = self.listproject.InsertItem(0, str(codigo))
            self.listproject.SetItem(index, 1, produto)
            self.listproject.SetItem(index, 2, local)
            self.listproject.SetItem(index, 3, desenho)
            self.listproject.SetItem(index, 4, cliente)
            self.listproject.SetItem(index, 5, os)
            self.listproject.SetItemData(index, count)
            self.itemDataMap[count] = (str(count), produto, local,
                                       desenho, cliente, os)
            count += 1

        self.listproject.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        self.listproject.SetColumnWidth(4, wx.LIST_AUTOSIZE_USEHEADER)
        lm.ColumnSorterMixin.__init__(self, 6)

        sizer2.Add(self.listproject, 1, wx.EXPAND, 0)

        sizer2.Add((20, 20), 0, 0, 0)

        sizer1.Add((20, 20), 0, 0, 0)

        gs1.AddGrowableCol(5)

        self.SetSizer(sizer1)
        self.Layout()
        self.Centre()

        self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.openProjeto,
                  self.listproject)
        self.Bind(wx.EVT_LIST_ITEM_DESELECTED, self.disableItem,
                  self.listproject)
        self.Bind(wx.EVT_LIST_ITEM_SELECTED, self.enableItem,
                  self.listproject)

        self.Bind(wx.EVT_LIST_COL_CLICK, self.OnColClick, self.listproject)

        self.Bind(wx.EVT_TEXT, self.buscaProjeto, self.busca)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        pub.subscribe(self.listener, 'refreshList')

    def GetListCtrl(self):
        return self.listproject

    def openProjeto(self, event):
        ix = self.listproject.GetFocusedItem()
        cod = self.listproject.GetItemText(ix, 0)
        framegrid = ViewProjeto.ViewProjeto(cod, None, wx.ID_ANY, "")
        framegrid.ShowModal()

    def novoProjeto(self, event):
        files_path = []
        origem = 'list'
        framenew = NovoProjeto.NovoProjeto(files_path, origem, None,
                                           wx.ID_ANY, "")
        framenew.ShowModal()

    def enableItem(self, event):
        self.btn_edit.Enable(True)
        self.btn_delete.Enable(True)

    def disableItem(self, event):
        self.btn_edit.Enable(False)
        self.btn_delete.Enable(False)

    def OnColClick(self, event):
        pass

    def editProjeto(self, event):
        ix = self.listproject.GetFocusedItem()
        cod = self.listproject.GetItemText(ix, 0)

        frmedit = EditProjeto.EditProjeto(cod, None,
                                          wx.ID_ANY, "")
        frmedit.ShowModal()

    def deleteProjeto(self, event):
        ix = self.listproject.GetFocusedItem()
        cod = self.listproject.GetItemText(ix, 0)
        nome = self.listproject.GetItemText(ix, 1)

        dlg = wx.MessageDialog(
            None,
            "Deseja excluir o projeto {}?".format(nome),
            'Aviso de remoção!', wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            db = Database()
            res = db.select("SELECT * FROM v_componentes\
                             WHERE codigo_projeto=?", (cod,))

            for r in (res):
                id_comp = r[1]
                db.delete("DELETE FROM v_componentes\
                          WHERE codigo=?", id_comp)

            db.delete("DELETE FROM v_projetos WHERE codigo=?", cod)
            self.listproject.DeleteItem(ix)
        else:
            pass

    def buscaProjeto(self, event):
        txt_search = self.busca.GetValue()

        if(txt_search == ''):
            self.refreshList(event)
            return
        else:
            self.listproject.DeleteAllItems()
            db = Database()
            records = db.select(
                """SELECT * FROM v_projetos\
                WHERE produto LIKE ? OR desenho LIKE ?
                ORDER BY num_os DESC;""",
                ('%' + txt_search + '%', '%' + txt_search + '%'))
            count = 0

            for v in records:
                codigo = v[1] if v[1] is not None else ''
                produto = v[2] if v[2] is not None else ''
                local = v[3] if v[3] is not None else ''
                desenho = v[4] if v[4] is not None else ''
                cliente = v[5] if v[5] is not None else ''
                os = v[6] if v[6] is not None else ''

                index = self.listproject.InsertItem(0, str(codigo))
                self.listproject.SetItem(index, 1, produto)
                self.listproject.SetItem(index, 2, local)
                self.listproject.SetItem(index, 3, desenho)
                self.listproject.SetItem(index, 4, cliente)
                self.listproject.SetItem(index, 5, os)
                count += 1

            self.listproject.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
            self.listproject.SetColumnWidth(4, wx.LIST_AUTOSIZE_USEHEADER)

    def listener(self, message, event):
        self.refreshList(event)

    def refreshList(self, event):
        self.listproject.DeleteAllItems()
        db = Database()
        records = db.select(
            """SELECT * FROM v_projetos ORDER BY num_os DESC;""", '')
        count = 0

        for v in records:
            codigo = v[1] if v[1] is not None else ''
            produto = v[2] if v[2] is not None else ''
            local = v[3] if v[3] is not None else ''
            desenho = v[4] if v[4] is not None else ''
            cliente = v[5] if v[5] is not None else ''
            os = v[6] if v[6] is not None else ''

            index = self.listproject.InsertItem(0, str(codigo))
            self.listproject.SetItem(index, 1, produto)
            self.listproject.SetItem(index, 2, local)
            self.listproject.SetItem(index, 3, desenho)
            self.listproject.SetItem(index, 4, cliente)
            self.listproject.SetItem(index, 5, os)
            count += 1

        self.listproject.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)
        self.listproject.SetColumnWidth(4, wx.LIST_AUTOSIZE_USEHEADER)

    def OnClose(self, event):
        self.Destroy()


class ProjetosInit(wx.App):
    def OnInit(self):
        self.frame = Projetos(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True
