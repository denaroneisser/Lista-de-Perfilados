__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-21
"""


import wx
import json
from utils.read_file import arrSortDimensao,\
    groupByMaterial, listByPos, natural_keys, roundDecimal

from utils.create_grid import createGrid
from utils import database
from utils.arr_funcs import multiplyArray
from utils.export import gerarPlanilha, gerarPDF


tuple_corte = [
    'Pos', 'Material', 'Dimensão', 'Peso (kg/m)',
    'Peso unit. (kg)', 'Qtd', 'Peso parcial (kg)']

tuple_insp = ['Pos', 'Material', 'Dimensão', 'Qtd', 'Inspeção']

tuple_res = ['Perfil', 'Peso (kg)', 'Obs.:']

tuple_list = ['Pos', 'Denominação', 'Dimensão', 'Peso (kg/m)',
              'Peso unit. (kg)', 'Qtd', 'Peso parcial (kg)']


class ViewProjeto(wx.Dialog):

    def __init__(self, cod_proj, *args, **kwds):
        wx.Frame.__init__(self, *args, style=wx.DEFAULT_DIALOG_STYLE)
        self.SetSize((1200, 650))
        self.SetSizeHints(1200, 650, 1200, 650)
        self.SetTitle(u'Lista de perfilados')
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))
        self.SetBackgroundColour(wx.Colour(236, 236, 236))

        self.cod_proj = cod_proj
        self.db = database.Database()
        rows_proj = self.db.select("SELECT * FROM v_projetos\
                                   WHERE codigo=?",
                                   (self.cod_proj,))

        produto = ''
        local = ''
        desenho = ''
        cliente = ''
        os = ''
        if(len(rows_proj) >= 1):
            produto = rows_proj[0][2]
            local = rows_proj[0][3]
            desenho = rows_proj[0][4]
            cliente = rows_proj[0][5]
            os = rows_proj[0][6]

        rows = self.db.select("SELECT * FROM v_componentes\
                              WHERE codigo_projeto=?",
                              (self.cod_proj, ))

        qtd_mult = 1
        data = '[]'
        arr_data = []
        for com in range(len(rows)):
            qtd_mult = rows[com][5]
            data = rows[com][6]
            res = json.loads(data)
            multiply = multiplyArray(res, qtd_mult)
            for m in range(len(multiply)):
                arr_data.append(multiply[m])

        arr_data.sort(key=natural_keys)

        self.group_list = groupByMaterial(arr_data)

        self.lst_insp = listByPos(arr_data)

        self.panel = wx.Panel(self, wx.ID_ANY)
        self.sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.sizer.Add((20, 20), 0, 0, 0)

        self.sizer1 = wx.BoxSizer(wx.VERTICAL)
        self.sizer.Add(self.sizer1, 0, wx.EXPAND | wx.RIGHT, 50)

        self.gs1 = wx.FlexGridSizer(10, 1, 0, 0)
        self.sizer1.Add(self.gs1, 1, wx.EXPAND, 0)

        set_main_font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                                wx.FONTWEIGHT_BOLD, 0, "")

        st_prod = wx.StaticText(self.panel, wx.ID_ANY, "Produto")
        st_prod.SetFont(set_main_font)
        self.gs1.Add(st_prod, 0, wx.TOP, 20)

        self.st_produto = wx.StaticText(self.panel, wx.ID_ANY, produto)
        self.gs1.Add(self.st_produto, 0, 0, 0)

        st_des = wx.StaticText(self.panel, wx.ID_ANY, "Desenho")
        st_des.SetFont(set_main_font)
        self.gs1.Add(st_des, 0, wx.TOP, 10)

        self.st_desenho = wx.StaticText(self.panel, wx.ID_ANY, desenho)
        self.gs1.Add(self.st_desenho, 0, 0, 0)

        st_loc = wx.StaticText(self.panel, wx.ID_ANY, "Local")
        st_loc.SetFont(set_main_font)
        self.gs1.Add(st_loc, 0, wx.TOP, 10)

        self.st_local = wx.StaticText(self.panel, wx.ID_ANY, local)
        self.gs1.Add(self.st_local, 0, 0, 0)

        st_cli = wx.StaticText(self.panel, wx.ID_ANY, "Cliente")
        st_cli.SetFont(set_main_font)
        self.gs1.Add(st_cli, 0, wx.TOP, 10)

        self.st_cliente = wx.StaticText(self.panel, wx.ID_ANY, cliente)
        self.gs1.Add(self.st_cliente, 0, 0, 0)

        st_os = wx.StaticText(self.panel, wx.ID_ANY, "O.S")
        st_os.SetFont(set_main_font)
        self.gs1.Add(st_os, 0, wx.TOP, 10)

        self.st_num_os = wx.StaticText(self.panel, wx.ID_ANY, os)
        self.gs1.Add(self.st_num_os, 0, 0, 0)

        ###########
        # SET GRIDSIZER 2
        #
        self.gs2 = wx.FlexGridSizer(1, 1, 0, 20)
        self.sizer1.Add(self.gs2, 0, wx.EXPAND | wx.TOP, 10)

        bmp_excel = wx.Bitmap("assets/excel.png", wx.BITMAP_TYPE_PNG)
        # bmp_pdf = wx.Bitmap("assets/pdf.png", wx.BITMAP_TYPE_PNG)

        styles_bmp = wx.BORDER_NONE | wx.BU_AUTODRAW |\
            wx.BU_EXACTFIT | wx.BU_NOTEXT

        self.bt_excel = wx.BitmapButton(self.panel, wx.ID_ANY, bmp_excel,
                                        style=styles_bmp)
        self.bt_excel.SetToolTip("Exportar para excel")
        self.bt_excel.SetSize(self.bt_excel.GetBestSize())
        self.bt_excel.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.exportExcel, self.bt_excel)
        self.gs2.Add(self.bt_excel, 0, 0, 0)

        # self.bt_pdf = wx.BitmapButton(self.panel, wx.ID_ANY, bmp_pdf,
        #                               style=styles_bmp)
        # self.bt_pdf.SetToolTip("Exportar para PDF")
        # self.bt_pdf.SetSize(self.bt_pdf.GetBestSize())
        # self.bt_pdf.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        # self.Bind(wx.EVT_BUTTON, self.exportPDF, self.bt_pdf)
        # self.gs2.Add(self.bt_pdf, 0, 0, 0)

        self.sizer1.Add((200, 20), 1, wx.EXPAND, 0)

        self.btn_fechar = wx.Button(self.panel, wx.ID_ANY, "Fechar",
                                    style=wx.BORDER_NONE)
        self.btn_fechar.SetBackgroundColour(wx.Colour(60, 155, 119))
        self.btn_fechar.SetForegroundColour(wx.Colour(255, 255, 255))
        self.btn_fechar.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_BUTTON, self.OnClose, self.btn_fechar)
        self.sizer1.Add(self.btn_fechar, 0, wx.BOTTOM, 20)

        self.notebook = wx.Notebook(self.panel, wx.ID_ANY, style=wx.NB_BOTTOM)
        self.sizer.Add(self.notebook, 1, wx.ALL | wx.EXPAND, 10)

        self.grd_corte = createGrid(self, self.notebook, 7, 120, tuple_corte)
        self.notebook.AddPage(self.grd_corte, "Corte")

        self.grd_insp = createGrid(self, self.notebook, 5, 120, tuple_insp)
        self.notebook.AddPage(self.grd_insp, u"Inspeção")

        self.grd_res = createGrid(self, self.notebook, 3, 120, tuple_res)
        self.notebook.AddPage(self.grd_res, "Resumo")

        self.grd_list = createGrid(self, self.notebook, 7, 120, tuple_list)
        self.notebook.AddPage(self.grd_list, "Lista de materiais")

        #
        #
        #

        arr_len = len(self.group_list)

        num_rows_corte = self.grd_corte.GetNumberRows()
        num_rows_insp = self.grd_insp.GetNumberRows()
        num_rows_res = self.grd_res.GetNumberRows()
        num_rows_list = self.grd_list.GetNumberRows()

        #
        #
        # CORTE GRID
        #
        #

        if num_rows_corte >= 1:
            self.grd_corte.DeleteRows(0, num_rows_corte)

        strc = 0
        for a in range(arr_len):
            it_rows = self.group_list[a][1]
            lst_mat = self.group_list[a][2]
            lst_mat.sort(key=arrSortDimensao, reverse=True)

            peso_total = self.group_list[a][3]
            peso_total = roundDecimal(peso_total)

            ap_rows = it_rows + 2
            self.grd_corte.AppendRows(ap_rows)
            self.grd_corte.ForceRefresh()

            m = -1
            for j in range(strc, strc + ap_rows):
                m = m + 1
                row = j

                if(m < len(lst_mat)):
                    for col in range(7):
                        it_material = lst_mat[m][col]
                        if(col == 3):
                            it_material = str(it_material).replace(".", ",")
                        elif(col == 4 or col == 6):
                            it_material = roundDecimal(it_material)
                            it_material = str(it_material).replace(".", ",")

                        self.grd_corte.SetCellValue(row, col, it_material)

                if(j == strc + ap_rows - 2):
                    self.grd_corte.SetCellValue(row, 4, 'PESO TOTAL=')
                    self.grd_corte.SetCellFont(row, 4,
                                               wx.Font(8,
                                                       wx.FONTFAMILY_DEFAULT,
                                                       wx.NORMAL, wx.BOLD))

                    str_outpeso = str(peso_total).replace(".", ",")
                    self.grd_corte.SetCellValue(row, 6, str_outpeso)
                    self.grd_corte.SetCellFont(row, 6,
                                               wx.Font(8,
                                                       wx.FONTFAMILY_DEFAULT,
                                                       wx.NORMAL, wx.BOLD))
                    attr = wx.grid.GridCellAttr()
                    attr.SetBackgroundColour(('#eaf7f5'))
                    self.grd_corte.SetRowAttr(row, attr)

            strc += ap_rows

        #
        #
        #
        # INSPEÇÃO GRID
        #
        #

        if num_rows_insp >= 1:
            self.grd_insp.DeleteRows(0, num_rows_insp)

        lst_insp_len = len(self.lst_insp)

        self.grd_insp.AppendRows(lst_insp_len)
        self.grd_insp.ForceRefresh()

        for a in range(lst_insp_len):
            it_qtd = self.lst_insp[a][5]

            for i2 in range(5):
                if(i2 == 3):
                    self.grd_insp.SetCellValue(a, i2, it_qtd)
                elif(i2 == 4):
                    self.grd_insp.SetCellValue(a, i2, '')
                else:
                    it_mat = self.lst_insp[a][i2]
                    self.grd_insp.SetCellValue(a, i2, it_mat)

        #
        #
        #
        # RESUMO GRID
        #
        #

        if num_rows_res >= 1:
            self.grd_res.DeleteRows(0, num_rows_res)

        self.grd_res.AppendRows(arr_len + 1)
        self.grd_res.ForceRefresh()

        arr_peso = []
        smpt = 0
        for a in range(arr_len):
            perfil = self.group_list[a][0]
            peso_total = self.group_list[a][3]
            arr_peso.append(peso_total)

            for i3 in range(3):
                if(i3 == 0):
                    self.grd_res.SetCellValue(a, i3, perfil)
                elif(i3 == 1):
                    peso_total = roundDecimal(peso_total)
                    ptr = str(peso_total).replace(".", ",")
                    self.grd_res.SetCellValue(a, i3, ptr)
                else:
                    self.grd_res.SetCellValue(a, i3, '')

        if(len(arr_peso) >= 1):
            smpt = sum(map(float, arr_peso))
            smpt = roundDecimal(smpt)

        self.grd_res.SetCellValue(arr_len, 0, 'TOTAL=')
        self.grd_res.SetCellFont(arr_len, 0,
                                 wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                         wx.NORMAL, wx.BOLD))

        self.grd_res.SetCellValue(arr_len, 1, str(smpt).replace(".", ","))
        self.grd_res.SetCellFont(arr_len, 1,
                                 wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                         wx.NORMAL, wx.BOLD))

        attr = wx.grid.GridCellAttr()
        attr.SetBackgroundColour(('#eaf7f5'))
        self.grd_res.SetRowAttr(arr_len, attr)

        #
        #
        #
        # LISTA DE MATERIAIS
        #
        #

        if num_rows_list >= 1:
            self.grd_list.DeleteRows(0, num_rows_list)

        qtd_data_rows = len(arr_data) + (len(rows) * 4)
        self.grd_list.AppendRows(qtd_data_rows)
        self.grd_list.ForceRefresh()

        rpb = 0  # range page break
        # rw = - 1
        for components in range(len(rows)):
            # nome_desenho = rows[components][3]
            # titulo = rows[components][4]
            qtd_mat = rows[components][5]

            data = rows[components][6]
            res = json.loads(data)
            arr_mult = multiplyArray(res, qtd_mat)
            arr_mat = listByPos(arr_mult)
            lst_mat_len = len(arr_mat)

            rng_list = rpb
            m = -1
            arr_peso_lst = []
            smpt_lst = 0
            for al in range(rng_list, rng_list + lst_mat_len):
                m = m + 1
                # rw = rw + 1
                peso_parcial = arr_mat[m][6]
                arr_peso_lst.append(peso_parcial)

                for m2 in range(7):
                    it_mat = arr_mat[m][m2]
                    if(m2 == 3 or m2 == 4 or m2 == 6):
                        it_mat = str(it_mat).replace(".", ",")
                    self.grd_list.SetCellValue(al, m2, str(it_mat))

            if(len(arr_peso_lst) >= 1):
                smpt_lst = sum(map(float, arr_peso_lst))
                smpt_lst = roundDecimal(smpt_lst)

            self.grd_list.SetCellValue(rng_list + lst_mat_len, 4,
                                       'TOTAL PARCIAL=')
            self.grd_list.SetCellFont(rng_list + lst_mat_len, 4,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            self.grd_list.SetCellValue(rng_list + lst_mat_len, 6,
                                       str(smpt_lst).replace(".", ","))
            self.grd_list.SetCellFont(rng_list + lst_mat_len, 6,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            self.grd_list.SetCellValue(rng_list + lst_mat_len + 1, 4,
                                       'GALVANIZAÇÃO')
            self.grd_list.SetCellTextColour(rng_list + lst_mat_len + 1, 4,
                                            wx.Colour('#666666'))
            self.grd_list.SetCellAlignment(rng_list + lst_mat_len + 1, 4,
                                           wx.ALIGN_LEFT, wx.ALIGN_CENTRE)
            self.grd_list.SetCellFont(rng_list + lst_mat_len + 1, 4,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            self.grd_list.SetCellValue(rng_list + lst_mat_len + 1, 6, '3,5%')
            self.grd_list.SetCellTextColour(rng_list + lst_mat_len + 1, 6,
                                            wx.Colour('#666666'))

            self.grd_list.SetCellAlignment(rng_list + lst_mat_len + 1, 6,
                                           wx.ALIGN_LEFT, wx.ALIGN_CENTRE)
            self.grd_list.SetCellFont(rng_list + lst_mat_len + 1, 6,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            self.grd_list.SetCellValue(rng_list + lst_mat_len + 2, 4,
                                       'PESO TOTAL=')
            self.grd_list.SetCellFont(rng_list + lst_mat_len + 2, 4,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            total_galv = smpt_lst + (smpt_lst * (3.5 / 100))
            total_galv = roundDecimal(total_galv)

            self.grd_list.SetCellValue(rng_list + lst_mat_len + 2, 6,
                                       str(total_galv).replace(".", ","))
            self.grd_list.SetCellFont(rng_list + lst_mat_len + 2, 6,
                                      wx.Font(8, wx.FONTFAMILY_DEFAULT,
                                              wx.NORMAL, wx.BOLD))

            attr2 = wx.grid.GridCellAttr()
            attr2.SetBackgroundColour(('#f9fbfc'))

            attr.IncRef()
            self.grd_list.SetRowAttr(rng_list + lst_mat_len, attr)
            self.grd_list.SetRowAttr(lst_mat_len + 1, attr2)
            self.grd_list.SetRowAttr(rng_list + lst_mat_len + 2, attr)
            attr.IncRef()

            rpb = rpb + lst_mat_len + 4

        #
        #
        #

        self.panel.SetSizer(self.sizer)
        self.Layout()
        self.Centre()

        self.Bind(wx.EVT_CLOSE, self.OnClose)
        # pub.subscribe(self.listener, 'choiceList')

    def exportExcel(self, event):

        produto = self.st_produto.GetLabel()
        local = self.st_local.GetLabel()
        desenho = self.st_desenho.GetLabel()
        cliente = self.st_cliente.GetLabel()
        num_os = self.st_num_os.GetLabel()
        lbl_lst = [produto, local, desenho, cliente, num_os]

        dirname = ''
        filename = ''
        mask = "Planilha (*.xlsx)|*.xlsx;*.xls*"
        with wx.FileDialog(self, "Exportar como",
                           dirname, filename,
                           mask, wx.FD_SAVE) as dlg:

            if dlg.ShowModal() == wx.ID_CANCEL:
                return

            path = dlg.GetPath()

            gp = gerarPlanilha(path, self.cod_proj, lbl_lst,
                               self.group_list, self.lst_insp)

            if(gp is not False):
                wx.MessageBox(f'Planilha slava em: {path}', 'Info', wx.OK)

    def exportPDF(self, event):

        dirname = ''
        filename = ''
        mask = "PDF - Adobe Portable Document Format (*.pdf)|*.pdf"
        with wx.FileDialog(self, "Exportar como",
                           dirname, filename,
                           mask, wx.FD_SAVE) as dlg:

            if dlg.ShowModal() == wx.ID_CANCEL:
                return

            path = dlg.GetPath()
            gerarPDF(path)
        pass

    def OnClose(self, event):
        self.Destroy()


class ShowProject(wx.App):
    def OnInit(self):
        self.frame = ViewProjeto(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True
