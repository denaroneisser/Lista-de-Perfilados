__author__ = 'FenanxCorp'
__version__ = '0.0.2'

"""
  Lista de Perfilados - Versátil Engenharia - Engefame LTDA
  2020-04-15
"""

import wx
import wx.grid as gridlib


def createGrid(self,
               setNotebook, sizeGrid, colSize, labels=[]):
    self.grid = gridlib.Grid(setNotebook, wx.ID_ANY, size=(1, 1))
    self.grid.CreateGrid(0, sizeGrid)

    self.grid.SetRowLabelSize(0)
    self.grid.SetColLabelSize(20)
    self.grid.SetColLabelAlignment(wx.ALIGN_LEFT, wx.ALIGN_CENTRE)
    self.grid.SetDefaultColSize(colSize, True)

    if(labels[2] == 'Obs.:'):
        self.grid.SetColSize(2, 320)

    self.grid.EnableEditing(0)
    self.grid.EnableDragColSize(0)
    self.grid.EnableDragRowSize(0)

    self.grid.SetGridLineColour(wx.Colour(192, 192, 192))
    self.grid.SetCellBackgroundColour(-1, -1, wx.Colour(245, 245, 245))
    self.grid.SetCellHighlightColour(wx.Colour(99, 175, 146))

    for i in range(len(labels)):
        self.grid.SetColLabelValue(i, labels[i])

    return self.grid
