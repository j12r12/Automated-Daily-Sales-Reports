from itertools import count
from datetime import datetime, date
import numpy as np


class Daily_Data:

    """Class to take in the initial information and process it ready to be used in creating graphs
    and excel sheets."""

    _ids = count(0)

    def __init__(self, df, type_remove, sales_person, stage):

        self.id = next(self._ids)

        self.df = df

        # Filter by Stage
        self.df = self.df[self.df["Stage"].isin(stage)]
        # Remove unwanted record types
        self.df = self.df[~self.df["Opportunity Type"].isin(type_remove)]
        # Filter by Sales Person
        self.df = self.df[self.df["Salesperson"].isin(sales_person)]

        # Set date to today's date and correct the format
        self.date = datetime.today().date().strftime("%d/%m/%Y")

    def pivot_revenue(self, val, ind):

        """Function to create a pivot table to show revenue per person."""

        self.df_today = self.df[self.df["Close Date"] == self.date]

        try:
            pvt_table = self.df_today.pivot_table(values=val, index=ind, columns="Stage",
                                                  aggfunc=np.sum, margins=True)
        except:
            pvt_table = pd.DataFrame([])
            print("No data available")

        return pvt_table

    def pivot_quantity(self, val, ind):

        """Function to create a pivot table to show quantities of sales per person."""

        self.df_today = self.df[self.df["Close Date"] == self.date]

        # Need to drop duplicates here as otherwise get the wrong quantity of sales.
        self.df_today = self.df_today.drop_duplicates(subset="Opportunity Name", keep="first")

        try:
            pvt_table = self.df_today.pivot_table(values=val, index=ind, columns="Stage", aggfunc="count", margins=True)
        except:
            pvt_table = pd.DataFrame([])

        return pvt_table

