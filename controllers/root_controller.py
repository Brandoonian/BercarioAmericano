from PyScripts.views.root_view import RootView
from PyScripts.controllers.report_controller import ReportController
from PyScripts.controllers.pnl_controller import PNLController



class RootController:

    def __init__(self):
        self.view = RootView(self)

    def run(self):
        self.view.mainloop()

    def report_controller(self):
        ReportController(self)

    def pnl_controller(self):
        PNLController(self)


