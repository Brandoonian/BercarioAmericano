from tkinter import Tk, Label, Button


class RootView(Tk):
    """Will inherit class from TKinter. CLASS IS TKINTER WINDOW"""
    def __init__(self, controller):
        super().__init__()
        self.geometry("500x500")
        self.title("TK WINDOW")

        Label(self, text="Home").grid(row=0, column=2)

        Button(self, text="Report Sale, Production, or Expense", command=controller.report_controller).grid(row=1, column=2)

        Button(self, text="View Expenses, Revenue, and Profit", command=controller.pnl_controller).grid(row=2, column=2)

