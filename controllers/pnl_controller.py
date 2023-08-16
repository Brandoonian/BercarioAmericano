from PyScripts.views.pnl_view import PNLView
import openpyxl

class PNLController:
    def __init__(self, root_controller):
        self.view = PNLView(self, root_controller.view)
        self.file_name = "./BerAmer_3.xlsx"
        self.book = openpyxl.load_workbook(self.file_name)
        self.sheet_1 = self.book["Sheet1"]
        self.sheet_1.title = "Sheet1"