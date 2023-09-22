__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import wx
import datetime

now = datetime.datetime.now()


class About(wx.Dialog):

    def __init__(self, *args, **kwds):
        wx.Dialog.__init__(self, *args, style=wx.DEFAULT_DIALOG_STYLE)
        self.SetSize((400, 220))
        self.SetSizeHints(400, 220, 400, 220)
        self.SetTitle(u'Sobre o sistema')
        self.SetIcon(wx.Icon('assets/icon.ico', wx.BITMAP_TYPE_ICO))
        self.SetBackgroundColour(wx.Colour(236, 236, 236))

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add((20, 10), 0, 0, 0)

        flexsizer = wx.FlexGridSizer(1, 2, 0, 0)
        sizer.Add(flexsizer, 1, wx.EXPAND, 0)

        setbitmap = wx.Bitmap("assets/icone_versatil.png", wx.BITMAP_TYPE_ANY)
        image = setbitmap.ConvertToImage()
        bmp = wx.Bitmap(image.Scale(70, 70, wx.IMAGE_QUALITY_HIGH))

        # image.Rescale(70, 70, wx.IMAGE_QUALITY_HIGH)
        # bmp = wx.Bitmap(image)

        bitmap = wx.StaticBitmap(self, wx.ID_ANY, bmp)

        bitmap.SetMinSize((80, 80))
        flexsizer.Add(bitmap, 1, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 59)

        sizer_1 = wx.FlexGridSizer(1, 1, 0, 0)
        flexsizer.Add(sizer_1, 1, wx.EXPAND | wx.LEFT, 50)

        nome_version = f"Lista de Perfilados\nVersão: {__version__}"

        static_text_1 = wx.StaticText(self, wx.ID_ANY, f"{nome_version}")
        sizer_1.Add(static_text_1, 1, wx.ALIGN_CENTER_VERTICAL, 0)

        sizer.Add((20, 10), 0, 0, 0)

        static_line_1 = wx.StaticLine(self, wx.ID_ANY)
        sizer.Add(static_line_1, 0, wx.EXPAND, 0)

        ano_atual = now.year
        nome_developer = f"Versátil Engenharia © {ano_atual}\n"
        nome_developer += "Desenvolvido por: FenanxCorp"

        static_text = wx.StaticText(self, wx.ID_ANY, f"{nome_developer}",
                                    style=wx.ALIGN_CENTER |
                                    wx.ST_ELLIPSIZE_MIDDLE)
        sizer.Add(static_text, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        sizer_1.AddGrowableRow(0)
        sizer_1.AddGrowableCol(0)

        flexsizer.AddGrowableRow(0)
        flexsizer.AddGrowableCol(1)

        self.SetSizer(sizer)
        self.Layout()
        self.Centre()


class ShowAbout(wx.App):
    def OnInit(self):
        self.frame = About(None, wx.ID_ANY, "")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True
