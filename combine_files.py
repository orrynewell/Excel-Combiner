"""
Excel/CSV File Combiner
Orry Newell
This script will create a graphic interface and allow the user to select files or folders with excel or csvs in order
to be combined.
27NOV2021
"""
import os
import tkinter.filedialog
from tkinter import *
import pandas as pd


def combine_files(in_lb, out_lb, db):
    """
    Combine Files - Actually combines the things, need to add the code for nested csvs in a folder.
    :param in_lb: The input listbox to get the files from
    :param out_lb: The output listbox that holds the destination to write the folder
    :param db: The dialog box for writing progress reports to
    :return: nothing
    """
    input_files = []
    for i in range(in_lb.size()):
        input_files.append(in_lb.get(i))
    out_path = out_lb.get(0)
    db.insert(tkinter.END, "Files Selected: ")
    for file in input_files:
        db.insert(tkinter.END, "\n\t{}".format(file))
    db.insert(tkinter.END, "\nOutput Path: {}\n".format(os.path.join(out_path, "combined_excel.xlsx")))
    full_df = pd.DataFrame()
    for file in input_files:
        db.insert(tkinter.END, "Processing {}...".format(file[file.rfind(r'/') + 1:]))
        if file.endswith('.xlsx'):
            excel_file = pd.ExcelFile(file)
            sheets = excel_file.sheet_names
            for sheet in sheets:
                temp_df = excel_file.parse(sheet_name=sheet)
                full_df = full_df.append(temp_df)
        elif file.endswith('.csv'):
            full_df = full_df.append(pd.read_csv(file), ignore_index=True)
        db.insert(tkinter.END, "Complete\n")
    full_df.to_excel(os.path.join(out_path, "combined_excel.xlsx"))
    db.insert(tkinter.END, "Process Complete - {}\n".format(os.path.join(out_path, "combined_excel.xlsx")))


def browse_input_files(lb):
    """
    Browse Input Files - Allows user to select multiple files to be combined
    :param lb: The input listbox that will hold the items.
    :return: nothing
    """
    filename = tkinter.filedialog.askopenfilenames(initialdir=r"C:\Users\orryn\OneDrive\Desktop\Work\Scripts",
                                                   title="Select a File")
    if filename:
        for file in filename:
            lb.insert(lb.size() + 1, file)


def browse_output_file(lb):
    """
    Browse Output Files - Gets the output location from the user.
    :param lb: The listbox containing the destination
    :return: nothing
    """
    filename = tkinter.filedialog.askdirectory(initialdir=r"C:\Users\orryn\OneDrive\Desktop\Work\Scripts",
                                               title="Select a File")
    if filename:
        if lb.size() == 0:
            lb.insert(0, filename)
        else:
            lb.delete(0, END)
            lb.insert(0, filename)


def delete_selected(lb):
    """
    Delete Selected - Allows the user to select multiple items in the listbox to be deleted
    :param lb: The input listbox containing source files
    :return: nothing
    """
    sel = lb.curselection()
    for index in sel[::-1]:
        lb.delete(index)


def make_window():
    """
    Creates the tkinter window and all its widgets
    :return: nothing
    """
    root = Tk()
    root.geometry('1000x600')
    root.title("File Combiner")
    input_listbox_lbl = Label(root,
                              text="Input Files")
    output_listbox_lbl = Label(root,
                               text="Output File")
    input_file_listbox = Listbox(root,
                                 height=10,
                                 width=85,
                                 activestyle="dotbox",
                                 selectmode=MULTIPLE)
    output_file_listbox = Listbox(root,
                                  height=10,
                                  width=62,
                                  activestyle="dotbox")
    input_explore_btn = Button(root,
                               text="Browse Files",
                               command=lambda: browse_input_files(input_file_listbox),
                               padx=40)
    input_delete_btn = Button(root,
                              text="Remove",
                              command=lambda: delete_selected(input_file_listbox),
                              padx=40)
    output_explore_btn = Button(root,
                                text="Browse Files",
                                command=lambda: browse_output_file(output_file_listbox),
                                padx=40)
    process_btn = Button(root,
                         text="Process",
                         command=lambda: combine_files(input_file_listbox, output_file_listbox, dialog_box),
                         padx=40)
    exit_btn = Button(root,
                      text="Exit",
                      command=root.destroy,
                      padx=40)
    dialog_box = Text(root,
                      heigh=10,
                      width=100,
                      bg="gray")

    input_listbox_lbl.grid(row=0, column=1, sticky=N, pady=20)
    output_listbox_lbl.grid(row=0, column=3, columnspan=2, sticky=N, pady=20)

    input_file_listbox.grid(row=1, column=0, columnspan=3, sticky=N, padx=20, pady=20)
    output_file_listbox.grid(row=1, column=3, columnspan=2, sticky=N, padx=20, pady=20)

    input_explore_btn.grid(row=2, column=0, columnspan=2, pady=2, padx=40)
    input_delete_btn.grid(row=2, column=1, columnspan=2, pady=2, padx=40)
    output_explore_btn.grid(row=2, column=3, columnspan=2, sticky=N, pady=2)

    dialog_box.grid(row=3, column=0, columnspan=5, pady=20)

    process_btn.grid(row=4, column=1, padx=20, pady=20)
    exit_btn.grid(row=4, column=3, padx=20, pady=20)

    root.mainloop()


def app():
    make_window()


if __name__ == "__main__":
    app()
