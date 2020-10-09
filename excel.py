import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from itertools import count
from graph_1 import Graph
import os


class Excel():
    _ids = count(0)

    def __init__(self, writer, df, sheet, application_path, unit=False):

        self.id = next(self._ids)

        self.df = df

        # Change the df "All" column and index created using df.pivot_table to "Total"
        self.df = self.df.rename(columns={"All": "Total"})
        self.df = self.df.rename(index={"All": "Total"})

        self.sheet = sheet
        self.writer = writer
        self.workbook = self.writer.book
        self.application_path = application_path

        # Setup bold format
        self.bold = self.workbook.add_format({"bold": True})

        # Create instance of graph class and save the created figure to working (or other chosen) directory
        self.graph = Graph(self.df).store_figure(application_path)

        # Create the next section to deal with formatting differently depending on whether we are
        # creating tables with revenue (currency) or quantities
        self.unit = unit

        if self.unit:
            self.currency = self.workbook.add_format({"num_format": "£#,##0.00"})
            self.bold_curr = self.workbook.add_format({"bold": True, "num_format": "£#,##0.00"})
            self.bold_curr_border = self.workbook.add_format({"bold": True,
                                                              "num_format": "£#,##0.00", "border": 1})
        else:
            self.currency = ""
            self.bold_curr = self.workbook.add_format({"bold": True})
            self.bold_curr_border = self.workbook.add_format({"bold": True, "border": 1})
            self.borders = self.workbook.add_format({"border": 1})

        # Set header format
        self.header_format = self.workbook.add_format({"bold": True,
                                                       "text_wrap": True,
                                                       "valign": "center", "fg_color": "#4F81BD",
                                                       "border": 1,
                                                       "font_color": "white"})

        # Setup start and end rows to determine where to apply formats etc.
        self.start_row = 1
        self.end_row = self.start_row + len(self.df) - 1
        self.count = 0

    def create_workbook(self):

        """Function to use the information and graphs from the Daily_Data and Graph classes
        to create an excel workbook displaying this data."""

        # Set up excel workbook
        self.df.to_excel(self.writer, sheet_name=self.sheet, startrow=self.start_row, header=False)

        # Set up worksheets
        self.worksheet = self.writer.sheets[self.sheet]
        [self.worksheet.write(0, col_num + 1, value, self.header_format) for col_num, value
         in enumerate(self.df.columns.values)]

        # Formatting
        self.worksheet.set_column(1, 3, 12, self.currency)
        self.worksheet.set_column(0, 1, 20)
        self.worksheet.set_column(3, 3, 12, self.bold_curr)
        self.worksheet.set_row(self.end_row, 15, self.bold_curr)

        # Insert graph
        self.worksheet.insert_image("F2", os.path.join(self.application_path, "Graph {}.png".format(self.id)))

        # Additional formatting depending on number of columns in the data
        lister = []

        if len(self.df.columns) > 2:
            for i in ["B", "C", "D"]:
                for j in range(2, len(self.df) + 2):
                    val = "{}{}".format(i, j)
                    lister.append(val)
            self.cell_range = "{}:{}".format(lister[0], lister[-1])
            self.worksheet.conditional_format(self.cell_range,
                                              {"type": "no_errors", "format":
                                                  self.workbook.add_format({'top': 1, 'bottom': 1,
                                                                            'left': 1, 'right': 1})})

        else:
            for i in ["B", "C"]:
                for j in range(2, len(self.df) + 2):
                    val = "{}{}".format(i, j)
                    lister.append(val)
            self.cell_range = "{}:{}".format(lister[0], lister[-1])
            self.worksheet.conditional_format(self.cell_range, {"type": "no_errors",
                                                                "format": self.workbook.add_format({
                                                                    'top': 1, 'bottom': 1,
                                                                    'left': 1, 'right': 1})})
