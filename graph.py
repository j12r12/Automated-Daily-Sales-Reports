import matplotlib.pyplot as plt
import os
from itertools import count
import numpy as np


class Graph:
    """This class uses the information created by the Daily_Data class to plot simple bar charts
    for as many different pivot tables as a required."""

    _ids = count(0)

    def __init__(self, df):

        self.id = next(self._ids)

        # Set the style to ggplot
        plt.style.use("ggplot")

        self.df = df

        # The data coming from the Daily_Data class instance has a total, however we don't want this
        # in the graph
        self.df = self.df.drop("Total")

        # Create figure and axis
        self.fig, self.ax = plt.subplots(edgecolor="black")
        self.fig.set_size_inches(5, 4)

        # Create graph
        try:
            x = np.arange(len(self.df.index))

            # Set the width of each bar
            self.width = np.min(np.diff(x)) / 3

            rects1 = self.ax.bar((x - self.width / 2), self.df["Won"],
                                 width=self.width, label="Won")
            rects2 = self.ax.bar((x + self.width / 2), self.df["Lost"],
                                 width=self.width, label="Lost")

            # Format created graph
            self.ax.legend(handles=[rects1, rects2], loc="upper right")
            self.ax.set_xticks(x)
            self.ax.set_xticklabels(self.df.index, rotation=70)
            self.ax.set_ylabel("Amount")
            self.fig.tight_layout()
            plt.rcParams.update({'font.size': 14})

        except:
            x = np.arange(len(self.df.index))

            # Set the width of each bar
            self.width = 0.8

            # Use column names for labels
            cols = self.df.columns
            cols = list(cols)
            lab = cols[0]

            rects1 = self.ax.bar(x, self.df[lab], width=self.width, label=lab)

            # Format created graph
            self.ax.legend(handles=[rects1], loc="upper right")
            self.ax.set_xticks(x)
            self.ax.set_xticklabels(self.df.index, rotation=45)
            self.ax.set_ylabel("Amount")
            self.fig.tight_layout()
            plt.rcParams.update({'font.size': 14})

    def store_figure(self, application_path):

        """Function to save the graph to file"""

        self.fig.savefig(os.path.join(application_path, "Graph {}.png".format(self.id)))

        return self.fig
