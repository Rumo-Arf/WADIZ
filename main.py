import wx
import os

from ReadExcel import ReadExcel
from WriteWord import WriteWord


class WadizWindow(wx.Dialog):
    excel_document = None
    word_document = None

    def __init__(self, parent, id, title):
        wx.Dialog.__init__(self, parent, id, title, size=(300, 200))

        wx.Button(self, 1, '1 참조할 엑셀 파일', (75, 50), size=(150, 20))
        wx.Button(self, 2, '2 베이스 워드 파일', (75, 100), size=(150, 20))
        wx.Button(self, 3, '3 종료', (75, 150), size=(150, 20))

        self.text = wx.StaticText(self, 4, "진행경과", pos=(0, 10), size=(300, 20), style=wx.ALIGN_CENTER)

        # self.number_cell1 = wx.TextCtrl(self, 5, )
        # self.number_cell2 = wx.TextCtrl(self, 5, )
        # self.howmuch_cell1 = wx.TextCtrl(self, 5, )
        # self.howmuch_cell2 = wx.TextCtrl(self, 5, )

        self.progress = wx.Gauge(self, range=100, size=(300, 10), style=wx.GA_HORIZONTAL)

        self.Bind(wx.EVT_BUTTON, self.openExcel, id=1)
        self.Bind(wx.EVT_BUTTON, self.openWord, id=2)
        self.Bind(wx.EVT_BUTTON, self.close, id=3)

        self.Centre()
        self.ShowModal()
        self.Destroy()

    def openExcel(self, event):
        dirname = '.'
        dlg = wx.FileDialog(self, "엑셀 파일을 선택 해 주세요", dirname, "", "*.xlsx", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)

        if dlg.ShowModal() == wx.ID_OK:
            self.excel_document = ReadExcel(dlg.GetPath())

        dlg.Destroy()

    def openWord(self, event):
        dirname = '.'
        dlg = wx.FileDialog(self, "워드 base 파일을 선택 해 주세요", dirname, "", "*.docx", wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)

        if dlg.ShowModal() == wx.ID_OK:
            self.word_document = WriteWord(dlg.GetPath(), self.excel_document, self.progress, self.text)

        dlg.Destroy()

    def close(self, event):
        self.Close(True)


app = wx.App(0)
WadizWindow(None, -1, title="와디즈 사무 자동화")
app.MainLoop()
