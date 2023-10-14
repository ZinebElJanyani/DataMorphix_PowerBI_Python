import os.path
from datetime import date

import openpyxl
from openpyxl.reader.excel import load_workbook
from ttkbootstrap import*
import tkinter as tk
from tkinter import messagebox
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.dialogs import Messagebox

from gestionRestauration import RestaurationInterface
from gestionTransport import TransportInterface

if __name__ == "__main__":
    app=RestaurationInterface()
    app.mainloop()