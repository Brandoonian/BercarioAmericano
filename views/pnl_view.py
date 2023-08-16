from tkinter import Label, Toplevel,Button
import openpyxl

class PNLView(Toplevel):
    def __init__(self, controller, rootview):
        super().__init__(rootview)
        self.title("Expenses, Revenue, Profit")
        self.back_butt = Button(self, text="Back", command=self.destroy).grid(row=4, column=1)
        file_name = "./BerAmer_3.xlsx"
        self.book = openpyxl.load_workbook(file_name)
        self.sheet_1 = self.book["Sheet1"]
        self.sheet_1.title = "Sheet1"
        self.controller = controller
        self.make_display()


    def make_display(self):
        self.update_variables()

        expenses_label = Label(self, text="EXPENSES:    ")
        expenses_label.grid(row=0, column=0)
        expenses_num = Label(self, text=self.total_expenses)
        expenses_num.grid(row=0, column=1)

        revenue_label = Label(self, text="REVENUE:    ")
        revenue_label.grid(row=1, column=0)
        revenue_num = Label(self, text=self.total_revenue)
        revenue_num.grid(row=1, column=1)

        profit_label = Label(self, text="PROFIT:    ")
        profit_label.grid(row=2, column=0)
        profit_num = Label(self, text=self.total_profit)
        profit_num.grid(row=2, column=1)

        profit_per_label = Label(self, text="PROFIT/UNIT SOLD:    ")
        profit_per_label.grid(row=3, column=0)
        #    profit_per_num = Label(self, text=profit_per)
        #   profit_per_num.grid(row=3, column=1)




    def update_variables(self):
        # Production
        self.total_qty_produced = int(self.sheet_1["B3"].value or 0)
        # Sales
        self.total_units_sold = int(self.sheet_1["C3"].value or 0)
        # Costs
        self.total_medium_cost = int(
            # R$ A generous estimate of the cost of growing medium is used for business purposes
            self.sheet_1["D3"].value or 0)  # Can be revised retrospectively.
        self.total_container_cost = int(self.sheet_1["E3"].value or 0)  # Sale Containers (Stays with product when sold)
        self.total_seed_cost = int(self.sheet_1["F3"].value or 0)  # R$
        self.total_variable_costs = int(self.sheet_1["G3"].value or 0)  # Costs that are NOT reoccurring ($R)
        self.total_delivery_costs = int(self.sheet_1["H3"].value or 0)

        self.total_expenses = int(self.sheet_1["D3"].value or 0) + int(self.sheet_1["E3"].value or 0) + \
                              int(self.sheet_1["F3"].value or 0) + int(self.sheet_1["G3"].value or 0) + \
                              int(self.sheet_1["H3"].value or 0)
        print(f"\n TOTAL COSTS: {self.total_expenses}")

        # cost_per_plant = total_expenses / int(sheet_1["B3"].value or 0)
        # print(f" COST PER PLANT: {cost_per_plant}")

        self.PRICE = 2  # R$

        self.total_revenue = self.total_units_sold * self.PRICE

        self.total_profit = self.total_revenue - self.total_expenses