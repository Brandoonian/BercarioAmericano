from tkinter import Label, Button, Toplevel, Entry
import openpyxl

class ReportView(Toplevel):
    def __init__(self, controller, rootview):
        super().__init__(rootview)
        self.geometry("500x500")
        self.title("NEW WINDOW")
        file_name = "./BerAmer_3.xlsx"
        self.book = openpyxl.load_workbook(file_name)
        self.sheet_1 = self.book["Sheet1"]
        self.sheet_1.title = "Sheet1"
        self.controller = controller
        self.open_report()
        self.date_entry = Entry(self) ##### NEW #####
        self.date_entry.grid(row=0, column=1)
        self.production_entry = Entry(self)
        self.production_entry.grid(row=1, column=1)
        self.sales_entry = Entry(self)
        self.sales_entry.grid(row=2, column=1)
        self.medium_buy_entry = Entry(self)
        self.medium_buy_entry.grid(row=3, column=1)
        self.container_buy_entry = Entry(self)
        self.container_buy_entry.grid(row=4, column=1)
        self.seed_buy_entry = Entry(self)
        self.seed_buy_entry.grid(row=5, column=1)
        self.variable_buy_entry = Entry(self)
        self.variable_buy_entry.grid(row=6, column=1)
        self.delivery_buy_entry = Entry(self)
        self.delivery_buy_entry.grid(row=7, column=1)

        # Production
        self.total_qty_produced = int(self.sheet_1["B3"].value or 0)
        # Sales
        self.total_units_sold = int(self.sheet_1["C3"].value or 0)
        # Costs
        self.total_medium_cost = int(  # R$ A generous estimate of the cost of growing medium is used for business purposes
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

        # self.profit_per = self.total_profit / self.total_qty_produced

        # print(f" PROFIT PER PLANT: {self.profit_per_plant}")


    def open_report(self):
        global date_entry, production_entry, sales_entry, medium_buy_entry, \
            container_buy_entry, seed_buy_entry, variable_buy_entry, \
            delivery_buy_entry

        self.update_variables()

        date_label = Label(self, text="Today's Date:")
        date_label.grid(row=0, column=0)

        production_label = Label(self, text="New Production:")
        production_label.grid(row=1, column=0)

        sales_label = Label(self, text="Sales:")
        sales_label.grid(row=2, column=0)

        medium_label = Label(self, text="Medium Purchase:")
        medium_label.grid(row=3, column=0)

        container_label = Label(self, text="Container Purchase:")
        container_label.grid(row=4, column=0)

        seed_label = Label(self, text="Seed Purchase:")
        seed_label.grid(row=5, column=0)

        variable_label = Label(self, text="Variable Purchase:")
        variable_label.grid(row=6, column=0)

        delivery_label = Label(self, text="Delivery Fee:")
        delivery_label.grid(row=7, column=0)
        print(f"{dir(self.controller)}")

        # Create 'Submit' button
        sub_butt = Button(self, text="Submit", command=self.controller.update_sheets)
        sub_butt.grid(row=8, column=1)

        back_butt = Button(self, text="Back", command=self.destroy)
        back_butt.grid(row=7, column=2)

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
