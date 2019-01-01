"""
This script will track changes in earned income.

Future versions will of the
class will be able to be used to track changes in any specific form of monthly
income or non-chash benefits. Additional features will include the ability to
track changes in any group of income types.
"""

__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"
__version__ = "1.0"

import pandas as pd

from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class Report:
    def __init__(self, file):
        """
        Initiate the data frame
        """
        self.entry_df = pd.read_excel(file, sheet_name="EntryData")
        self.exit_df = pd.read_excel(file, sheet_name="ExitInterimData")

    def return_most_recent_employment(self, df):
        """
        Return a data frame that only contains the most recent version of the
        Source of Income(141) column where the value of said column is equal to
        Earned Income(HUD)
        """
        # create a local copy of the dataframe with the values sorted by Client
        # Unique ID and Entry Exit Entry Date dropping duplicates by Client
        # Unique ID
        o_df = df[df["Source of Income(141)"] == "Earned Income (HUD)"].sort_values(
            by=["Client Unique Id", "Start Date(841)"],
            ascending=False
        ).drop_duplicates(subset="Client Unique Id", keep="first")

        # return the data frame
        return o_df

    def merge_dfs(self, e_df, ex_df):
        """
        Create a new dataframe by merging entry and exit dataframes then return
        three dataframes, one showing income gain, one showing income decrease,
        and one showing no income change.
        """
        # merge the entry dataframe and the exit dataframe on the client unique
        # id column filling nan values in the monthly amount fields with 0
        merged = e_df[[
            "Client Unique Id",
            "Client Uid",
            "Monthly Amount(142)"
        ]].merge(
            ex_df[["Client Unique Id", "Client Uid", "Monthly Amount(142)"]],
            on=["Client Unique Id", "Client Uid"],
            how="left",
            suffixes=("_entry", "_exit")
        ).fillna(0)

        # create the gain, loss, no_change df
        gain = merged[
            merged["Monthly Amount(142)_entry"] < merged["Monthly Amount(142)_exit"]
        ]
        loss = merged[
            merged["Monthly Amount(142)_entry"] > merged["Monthly Amount(142)_exit"]
        ]
        no_change = merged[
            merged["Monthly Amount(142)_entry"] == merged["Monthly Amount(142)_exit"]
        ]

        # return the dataframes
        return gain, loss, no_change

    def process(self):
        """
        Process the raw report and save the processed sheets to a new excel
        workbook
        """
        # create the three processed sheets using the merge_dfs and
        # return_most_recent_employment methods
        gain, loss, no_change = self.merge_dfs(
            self.return_most_recent_employment(self.entry_df),
            self.return_most_recent_employment(self.exit_df)
        )

        # create the writer object
        writer = pd.ExcelWriter(
            asksaveasfilename(
                title="Save the Income Change report",
                initialfile="Income Change Report(Processed)",
                defaultextension=".xlsx"
            ),
            engine="xlsxwriter"
        )

        # create the individual sheets
        gain.to_excel(writer, sheet_name="PTs with Income Growth", index=False)
        loss.to_excel(writer, sheet_name="PTs with Income Loss", index=False)
        no_change.to_excel(writer, sheet_name="PTs with No Income Change", index=False)
        self.entry_df.to_excel(writer, sheet_name="Raw Entry Data", index=False)
        self.exit_df.to_excel(writer, sheet_name="Raw Exit and Interim Data", index=False)

        # save the excel workbook and return True
        writer.save()
        return True

if __name__ == "__main__":
    file = askopenfilename(title="Open the Income Change report")
    a = Report(file)
    a.process()
