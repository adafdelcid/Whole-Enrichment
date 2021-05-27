 Author: Ada Del Cid
# GitHub: @adafdelcid
# May 2021

# GUI_Form_Enrichment: Graphical user interface for Formulation_Enrichment_by_Cell_Type.py
from tkinter import filedialog, Tk, StringVar, Label, Button, Entry, OptionMenu
from os import path
import sys

import Formulation_Enrichment_by_Cell_Type


class MyGUI:  # pylint: disable=too-many-instance-attributes
    # GUI for enrichment analysis by cell type

    def __init__(self, master):
        # Saves basic GUI buttons and data entries

        self.master = master
        master.geometry("600x700")
        master.title("Enrichment Analysis by Cell Type Tool")

        # string variables from user input
        self.fsp = StringVar()  # Formulation Sheet file Path
        self.ncp = StringVar()  # Normalized Counts file Path
        self.sc = StringVar()  # list of Sorted Cells
        self.dfp = StringVar()  # Destination Folder Path
        self.tbp = StringVar()  # Top/Bottom Percent
        self.nbc = StringVar()  # number of naked barcodes
        self.snl = StringVar()  # list of samples numbers for experiment
        self.fid = StringVar()  # file identifier
        self.rom = StringVar()  # boolean remove outlying mice
        self.r2t = StringVar()  # float of r^2 threshold
        self.rra = StringVar()  # boolean remove runaways
        self.op = StringVar()  # Outliers Percentile

        Label(master, text="Whole Enrichment Analysis", relief="solid", font=("arial", 16, "bold")).pack()

        Label(master, text="Formulation Sheet File Path", font=("arial", 12, "bold")).place(x=20, y=50)
        Button(master, text="Formulation sheet file", width=20, fg="green", font=("arial", 16),
               command=self.open_excel_file).place(x=370, y=48)

        Label(master, text="Normalized Counts CSV File Path", font=("arial", 12, "bold")).place(x=20, y=82)
        Button(master, text="Normalized counts file", width=20, fg="green", font=("arial", 16),
               command=self.open_csv_file).place(x=370, y=80)

        Label(master, text="List of Sorted Cells (separate by commas)", font=("arial", 12, "bold")).place(x=20, y=114)
        Entry(master, textvariable=self.sc).place(x=370, y=112)

        Label(master, text="Destination Folder Path", font=("arial", 12, "bold")).place(x=20, y=146)
        Entry(master, textvariable=self.dfp).place(x=370, y=144)

        Label(master, text="Top/Bottom Percent", font=("arial", 12, "bold")).place(x=20, y=178)
        Entry(master, textvariable=self.tbp).place(x=370, y=176)

        Label(master, text="Number of 'Naked' Barcodes", font=("arial", 12, "bold")).place(x=20, y=210)
        Entry(master, textvariable=self.nbc).place(x=370, y=208)

        Label(master, text="List of Sample Numbers", font=("arial", 12, "bold")).place(x=20, y=242)
        Entry(master, textvariable=self.snl).place(x=370, y=240)

        Label(master, text="File identifier (OPTIONAL)", font=("arial", 12, "bold")).place(x=20, y=274)
        Entry(master, textvariable=self.fid).place(x=370, y=272)

        # remove outlying mice
        Label(master, text="Remove outlying mice", font=("arial", 12, "bold")).place(x=20, y=304)
        options = ["Yes", "No"]

        rom_droplist = OptionMenu(master, self.rom, *options)
        self.rom.set("Yes or no")
        rom_droplist.config(width=15)
        rom_droplist.place(x=370, y=304)

        # threshold r^2 to remove mice
        Label(master, text="r^2 threshold (OPTIONAL: Default = 0.80)", font=("arial", 12, "bold")).place(x=20, y=334)
        Entry(master, textvariable=self.r2t).place(x=370, y=332)

        # remove runaways
        Label(master, text="Remove runaways", font=("arial", 12, "bold")).place(x=20, y=364)
        options = ["Yes", "No"]

        rra_droplist = OptionMenu(master, self.rra, *options)
        self.rra.set("Yes or no")
        rra_droplist.config(width=15)
        rra_droplist.place(x=370, y=362)

        # percentile by which to remove outliers
        Label(master, text="Outliers Percentile (OPTIONAL: Default = 99.9)", font=("arial", 12, "bold")).place(x=20,
                                                                                                               y=394)
        Entry(master, textvariable=self.op).place(x=370, y=392)

        Button(master, text="ENTER", width=16, fg="blue", font=("arial", 16), command=self.enrichment_analysis).place(
            x=150, y=424)
        Button(master, text="CANCEL", width=16, fg="blue", font=("arial", 16), command=exit1).place(x=300, y=424)

    def open_excel_file(self):
        # To open a file searcher and select a file
        self.fsp = filedialog.askopenfilename()

    def open_csv_file(self):
        # To open a file searcher and csv file
        self.ncp = filedialog.askopenfilename()

    def enrichment_analysis(self):  # pylint: disable=too-many-branches
        # pylint: disable=too-many-statements
        # Checks for any entry errors, returns list of errors or runs the enrichment analysis
        errors = False

        cell_types = string_to_list(self.sc.get())  # list of cell types
        fold_path = self.dfp.get()  # Destination Folder Path
        percent = self.tbp.get()  # Top/Bottom Percent
        num_bcs = self.nbc.get()  # Number of barcodes
        sample_num_list = string_to_list(self.snl.get())  # List of sample numbers
        remove_outlying_mice = self.rom.get()  # Option to remove outlying mice
        r2_threshold = self.r2t.get()  # r^2 threshold to remove mice
        remove_runaways = self.rra.get()  # Option to remove runaways
        percentile = self.op.get()  # Outliers Percentile

        # check for errors with formulation sheet
        if self.fsp == "PY_VAR0":
            color1 = "red"
            errors = True
        else:
            try:
                if ".xlsx" in self.fsp and path_exists(self.fsp):
                    color1 = "white"
                else:
                    color1 = "red"
                    errors = True
            except TypeError:
                color1 = "red"
                errors = True

        Label(self.master, text="Invalid formulation sheet file path", fg=color1, font=("arial", 12, "bold")).place(
            x=20, y=454)

        # check for errors with normalized counts csv file
        if self.ncp == "PY_VAR1":
            color2 = "red"
            errors = True
        else:
            try:
                if ".csv" in self.ncp and path_exists(self.ncp):
                    color2 = "white"
                else:
                    color2 = "red"
                    errors = True
            except TypeError:
                color2 = "red"
                errors = True

        Label(self.master, text="Invalid normalized counts file path", fg=color2, font=("arial", 12, "bold")).place(
            x=20, y=474)

        if cell_types == [""]:
            color3 = "red"
            errors = True
        else:
            color3 = "white"

        Label(self.master, text="Missing list of sorted cells", fg=color3, font=("arial", 12, "bold")).place(
            x=20, y=494)

        # check for errors with destination folder
        if path_exists(fold_path):
            color4 = "white"
        else:
            color4 = "red"
            errors = True

        Label(self.master, text="Invalid destination folder", fg=color4, font=("arial", 12, "bold")).place(x=20, y=514)

        # check for errors with top/bottom percent
        try:
            color5 = "white"
            percent = float(percent)
        except ValueError:
            color5 = "red"
            errors = True

        Label(self.master, text="Invalid top/bottom percent, enter a value between 0.1-99.9", fg=color5,
              font=("arial", 12, "bold")).place(x=20, y=534)

        # check if number of naked barcodes values to check for error here
        try:
            color6 = "white"
            num_bcs = int(num_bcs)
        except ValueError:
            color6 = "red"
            errors = True

        Label(self.master, text="Invalid number of naked barcodes", fg=color6, font=("arial", 12, "bold")).place(x=20,
                                                                                                                 y=554)
        # check list of sample number for error here
        if sample_num_list != [""]:
            color7 = "white"
        else:
            color7 = "red"
            errors = True

        Label(self.master, text="Missing list of sample numbers", fg=color7, font=("arial", 12, "bold")).place(x=20,
                                                                                                               y=574)

        # check remove outlying mice
        r_outlying_mice = False
        if remove_outlying_mice == "Yes or no":
            color8 = "red"
            errors = True
        else:
            color8 = "white"
            if remove_outlying_mice == "Yes":
                r_outlying_mice = True
            else:
                r_outlying_mice = False

        Label(self.master, text="Missing selection on mouse removal", fg=color8, font=("arial", 12, "bold")
              ).place(x=20, y=594)

        if r2_threshold == "":
            r2_threshold = 0.80
            color9 = "white"
        else:
            if r_outlying_mice:
                try:
                    color9 = "white"
                    r2_threshold = float(r2_threshold)
                except ValueError:
                    color9 = "red"
                    errors = True
            else:
                color9 = "white"

        Label(self.master, text="Enter valid r^2 threshold or leave box empty for default", fg=color9,
              font=("arial", 12, "bold")).place(x=20, y=634)

        # check remove outlying mice
        r_runaways = False
        if remove_runaways == "Yes or no":
            color10 = "red"
            errors = True
        else:
            color10 = "white"
            if remove_runaways == "Yes":
                r_runaways = True
            else:
                r_runaways = False

        Label(self.master, text="Missing selection on removal of runaways", fg=color10, font=("arial", 12, "bold")
              ).place(x=20, y=614)

        if percentile == "":
            percentile = 99.9
            color11 = "white"
        else:
            if r_runaways:
                try:
                    color11 = "white"
                    percentile = float(percentile)
                except ValueError:
                    color11 = "red"
                    errors = True
            else:
                color11 = "white"

        Label(self.master, text="Enter valid percentile or leave box empty for default", fg=color11,
              font=("arial", 12, "bold")).place(x=20, y=654)

        if not errors:
            Formulation_Enrichment_by_Cell_Type.run_enrichment_analysis(fold_path, self.fid.get(), self.fsp, self.ncp,
                                                                        cell_types, num_bcs, percent, sample_num_list,
                                                                        r_outlying_mice, r2_threshold, r_runaways,
                                                                        percentile)
            print("Enrichment analysis performed!")
            exit1()
        else:
            self.master.mainloop()


def exit1():
    # exit and close GUI

    sys.exit()


def string_to_list(string1):
    # Creates a list out of a string of items separated by commas

    string1 = remove_spaces(string1)
    list1 = list(string1.split(","))
    return list1


def remove_spaces(string1):
    # remove any unnecessary spaces

    return string1.replace(" ", "")


def path_exists(path1):
    # Checks if a directory path exists

    return path.exists(path1)


root = Tk()
my_gui = MyGUI(root)
root.mainloop()
