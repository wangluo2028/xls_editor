import tkinter as tk
from tksheet import Sheet

class SheetTable(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.state("zoomed")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sheet_demo = Sheet(self,
                                width=1000,
                                height=700,
                                align="w",
                                header_align="center",
                                row_index_align="center",
                                column_width=180,
                                row_index_width=50,
                                total_rows=5000,
                                total_columns=100,
                                headers=[f"Header {c}" for c in range(100)])
        self.sheet_demo.enable_bindings(("single",
                                         "drag_select",
                                         "column_drag_and_drop",
                                         "row_drag_and_drop",
                                         "column_select",
                                         "row_select",
                                         "column_width_resize",
                                         "double_click_column_resize",
                                         "row_width_resize",
                                         "column_height_resize",
                                         "arrowkeys",
                                         "row_height_resize",
                                         "double_click_row_resize",
                                         "edit_bindings"))
        self.sheet_demo.grid(row=0, column=0, sticky="nswe")
        self.sheet_demo.highlight_cells(row=0, column=0, bg="orange", fg="blue")
        self.sheet_demo.highlight_cells(row=0, bg="orange", fg="blue", canvas="row_index")
        self.sheet_demo.highlight_cells(column=0, bg="orange", fg="blue", canvas="header")

        """_________________________ EXAMPLES _________________________ """
        """_____________________________________________________________"""

        # __________ SETTING OR RESETTING TABLE DATA __________

        # self.data = [[f"Row {r} Column {c}" for c in range(100)] for r in range(5000)]
        # self.sheet_demo.data_reference(self.data)

        # __________ SETTING HEADERS __________

        # self.headers = [f"Header {c}" for c in range(100)]
        # self.sheet_demo.headers(self.headers)