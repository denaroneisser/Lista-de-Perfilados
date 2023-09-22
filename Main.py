__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import wx
import wx.adv
import sys
import NovoProjeto
import ListProjetos
import About


class Main(wx.Frame):

    def __init__(self, parent):
        style = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(
            self, parent, id=wx.ID_ANY,
            title=u"Lista de Perfilados - Versátil Engenharia",
            pos=wx.DefaultPosition, size=wx.Size(1200, 650),
            style=style)
        self.SetSizeHints(wx.Size(850, 550), wx.DefaultSize)
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))

        self.SetBackgroundColour('#f5f5f5')
        self.Maximize(True)

        self.menubar = wx.MenuBar(0 | wx.BORDER_NONE)

        # #############
        # ## MENU PRINCIPAL

        self.menuFile = wx.Menu()
        self.menubar.Append(self.menuFile, u"Menu")

        self.m_menunovo = wx.MenuItem(
            self.menuFile, wx.ID_ANY,
            u"Novo projeto", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menunovo.SetBitmap(wx.Bitmap('assets/new.png',
                                            wx.BITMAP_TYPE_PNG))
        self.menuFile.Append(self.m_menunovo)

        self.Bind(wx.EVT_MENU, self.novoProjeto, self.m_menunovo)

        self.menuFile.AppendSeparator()

        self.m_menuprojetos = wx.MenuItem(
            self.menuFile, wx.ID_ANY,
            u"Projetos", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menuprojetos.SetBitmap(wx.Bitmap('assets/projects.png',
                                                wx.BITMAP_TYPE_PNG))
        self.menuFile.Append(self.m_menuprojetos)
        self.Bind(wx.EVT_MENU, self.projetos, self.m_menuprojetos)

        self.menuFile.AppendSeparator()

        self.m_menuimport = wx.MenuItem(
            self.menuFile, wx.ID_ANY,
            u"Importar arquivo(s)", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menuimport.SetBitmap(wx.Bitmap('assets/import.png',
                                              wx.BITMAP_TYPE_PNG))
        self.menuFile.Append(self.m_menuimport)

        self.Bind(wx.EVT_MENU, self.importArquivo, self.m_menuimport)

        self.menuFile.AppendSeparator()

        self.m_menusair = wx.MenuItem(
            self.menuFile, wx.ID_ANY,
            u"Sair", wx.EmptyString, wx.ITEM_NORMAL)
        self.menuFile.Append(self.m_menusair)

        self.Bind(wx.EVT_MENU, self.OnExit, self.m_menusair)

        # #############
        # ## MENU AJUDA

        self.menuAjuda = wx.Menu()
        self.menubar.Append(self.menuAjuda, u"Ajuda e suporte")

        self.m_menusobre = wx.MenuItem(
            self.menuAjuda, wx.ID_ANY,
            u"Sobre o sistema", wx.EmptyString, wx.ITEM_NORMAL)
        self.m_menusobre.SetBitmap(wx.Bitmap('assets/info.png',
                                             wx.BITMAP_TYPE_PNG))
        self.menuAjuda.Append(self.m_menusobre)
        self.Bind(wx.EVT_MENU, self.sobre, self.m_menusobre)

        self.SetMenuBar(self.menubar)

        self.initUI(None)
        self.Centre(wx.BOTH)

    def initUI(self, event):

        self.panel = wx.Panel(self, wx.ID_ANY)

        bmp_init = wx.Bitmap("assets/ico_versatil_peb.png", wx.BITMAP_TYPE_PNG)
        self.bitmp = wx.StaticBitmap(self.panel, wx.ID_ANY, bmp_init)
        self.bitmp.SetBackgroundColour(wx.Colour(240, 240, 240))

        self.sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.sizer1 = wx.BoxSizer(wx.VERTICAL)

        self.sizer.Add(self.sizer1, 0, wx.EXPAND, 0)
        self.sizer1.Add((280, 50), 0, 0, 0)

        self.btn_projeto = wx.Button(self.panel, wx.ID_ANY, "Novo projeto",
                                     style=wx.BORDER_NONE)
        self.btn_projeto.SetBackgroundColour('#F5F5F5')
        self.btn_projeto.SetForegroundColour(wx.Colour(128, 128, 128))
        self.btn_projeto.SetFont(wx.Font(13, wx.FONTFAMILY_DEFAULT,
                                         wx.FONTSTYLE_NORMAL,
                                         wx.FONTWEIGHT_BOLD, 0, ""))
        self.btn_projeto.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.sizer1.Add(self.btn_projeto, 0, wx.LEFT, 10)
        self.Bind(wx.EVT_BUTTON, self.novoProjeto, self.btn_projeto)

        self.sizer1.Add((100, 20), 0, 0, 0)

        self.btn_listprojetos = wx.Button(self.panel, wx.ID_ANY, "Projetos",
                                          style=wx.BORDER_NONE)
        self.btn_listprojetos.SetBackgroundColour('#F5F5F5')
        self.btn_listprojetos.SetForegroundColour(wx.Colour(128, 128, 128))
        self.btn_listprojetos.SetFont(wx.Font(13, wx.FONTFAMILY_DEFAULT,
                                              wx.FONTSTYLE_NORMAL,
                                              wx.FONTWEIGHT_BOLD, 0, ""))
        self.btn_listprojetos.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.sizer1.Add(self.btn_listprojetos, 0, wx.LEFT, 10)
        self.Bind(wx.EVT_BUTTON, self.projetos, self.btn_listprojetos)

        self.sizer1.Add((100, 20), 0, 0, 0)

        self.import_file = wx.Button(self.panel, wx.ID_ANY,
                                     "Importar arquivo(s)",
                                     style=wx.BORDER_NONE)
        self.import_file.SetBackgroundColour('#F5F5F5')
        self.import_file.SetForegroundColour(wx.Colour(128, 128, 128))
        self.import_file.SetFont(wx.Font(13, wx.FONTFAMILY_DEFAULT,
                                         wx.FONTSTYLE_NORMAL,
                                         wx.FONTWEIGHT_BOLD, 0, ""))
        self.import_file.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.sizer1.Add(self.import_file, 0, wx.LEFT, 10)
        self.Bind(wx.EVT_BUTTON, self.importArquivo, self.import_file)

        self.sizer1.Add((100, 30), 0, 0, 0)

        self.sizer.Add(self.bitmp, 1, wx.ALL | wx.EXPAND, 10)
        self.panel.SetSizer(self.sizer)
        self.panel.Layout()

    def novoProjeto(self, event):
        files_path = []
        origem = 'main'
        framenew = NovoProjeto.NovoProjeto(files_path, origem, None,
                                           wx.ID_ANY, "")
        framenew.ShowModal()

    def projetos(self, event):
        frmprojetos = ListProjetos.Projetos(None, wx.ID_ANY, "")
        frmprojetos.ShowModal()

    def importArquivo(self, event):
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

            files_path = opnFileDlg.GetPaths()

        origem = 'main'
        frameproj = NovoProjeto.NovoProjeto(files_path, origem, None,
                                            wx.ID_ANY, "")
        frameproj.ShowModal()

    def sobre(self, event):
        framesobre = About.About(None, wx.ID_ANY, "")
        framesobre.ShowModal()

    def OnExit(self, event):
        self.Close()
        sys.exit(0)


class SplashScreen(wx.adv.SplashScreen):

    def __init__(self, parent=None):

        aBitmap = wx.Bitmap("assets/splash.png", wx.BITMAP_TYPE_PNG)
        splashStyle = wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT
        splashDuration = 1500
        sstyle = wx.NO_BORDER | wx.FRAME_NO_TASKBAR | wx.STAY_ON_TOP

        wi, hi = aBitmap.GetWidth(), aBitmap.GetHeight()

        wx.adv.SplashScreen.__init__(
            self, aBitmap, splashStyle, splashDuration,
            parent, id=wx.ID_ANY, size=wx.Size(wi, hi), style=sstyle)

        self.Bind(wx.EVT_CLOSE, self.OnExit)

    def OnExit(self, evt):
        self.Hide()
        startframe = Main(None)
        app.SetTopWindow(startframe)
        startframe.Centre()
        startframe.Show(True)

        evt.Skip()


if __name__ == '__main__':

    app = wx.App(False)
    runApp = SplashScreen()
    runApp.Show(True)
    app.MainLoop()
