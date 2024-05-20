import sys
import pandas as pd
import re
from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext
from tkinter.filedialog import askopenfilename

excel_file = None
result = ("Welcome to the Sankey Chart Maker.\n\n"
          "This program will take an excel file, read the headers of a chart, "
          "and by your inputs create the code for a sankey chart.\n"
          "The inputs must be an exact match for the program to work properly.\n"
          "The format that will display will be in source[amount]target form.\n"
          "By choosing the single sources you will be able to create a custom result.\n"
          "Select an excel file to begin, or start creating custom results.")

"""def setup_logging():
    log_file = "output.log"
    sys.stdout = open(log_file, 'w')
    sys.stderr = open(log_file, 'w')
    print("Logging to", log_file)"""


def sankey_format(sheet, source, amount, target):
    df = pd.read_excel(sheet)

    head_df = pd.read_excel(sheet, nrows=0)

    headers = head_df.columns.tolist()

    arr = []

    if source in headers and amount in headers and target in headers:

        for i in range(len(df)):
            skip = False
            source1 = df.loc[[i], [source]].to_string(index=False, header=False)
            amount1 = df.loc[[i], [amount]].to_string(index=False, header=False)
            target1 = df.loc[[i], [target]].to_string(index=False, header=False)

            if source1 == "NaN" or amount1 == "NaN" or target1 == "NaN":
                skip = True

            if is_num(amount1):
                if not skip:
                    text = f"{source1}[{amount1}]{target1}\n"
                    arr.append(text)
            else:
                arr.append("The amount " + amount1 + "needs to be a valid number.")

    if source not in headers:
        arr.append("There are no headers with the name of: " + source + "\n")

    if amount not in headers:
        arr.append("There are no headers with the name of: " + amount + "\n")

    if target not in headers:
        arr.append("There are no headers with the name of: " + target + "\n")

    return ''.join(arr)


def on_file_clicked():
    global excel_file
    excel_file = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])


def on_submit_clicked():
    global result
    get_source = source.get()
    get_amount = amount.get()
    get_target = target.get()
    get_single_source = single_source.get()
    get_single_amount = single_amount.get()
    get_single_target = single_target.get()

    if excel_file is None:
        if get_single_amount and get_single_target and get_single_source:
            if is_num(get_single_amount):
                result = f"{get_single_source}[{get_single_amount}]{get_single_target}\n"
            else:
                result = "The amount needs to be a number."
        elif get_single_source or get_single_target or get_single_amount:
            result = "Please fill out all single fields"
        else:
            result = "Please chose an excel file."

    else:
        if get_single_amount and get_single_target and get_single_source:
            result = f"{get_single_source}[{get_single_amount}]{get_single_target}\n"
        elif get_single_source or get_single_target or get_single_amount:
            result = "Please fill out all single fields"
        else:
            result = sankey_format(excel_file, get_source, get_amount, get_target)

    result_label.delete('1.0', END)
    result_label.insert(END, result)


def on_single_source_change(event):
    if single_source.get().strip():
        if source.get().strip():
            result_label.delete('1.0', END)
            result_label.insert(END, "Please choose only one source field.")
        source.delete(0, END)
        source.config(state=DISABLED)

    else:
        source.config(state=NORMAL)


def on_single_amount_change(event):
    if single_amount.get().strip():
        if amount.get().strip():
            result_label.delete('1.0', END)
            result_label.insert(END, "Please choose only one amount field.")
        amount.delete(0, END)
        amount.config(state=DISABLED)

    else:
        amount.config(state=NORMAL)


def on_single_target_change(event):
    if single_target.get().strip():
        if target.get().strip():
            result_label.delete('1.0', END)
            result_label.insert(END, "Please choose only one target field.")
        target.delete(0, END)
        target.config(state=DISABLED)

    else:
        target.config(state=NORMAL)


def is_num(num):
    pattern = r'^[0-9]*\.?[0-9]+$'
    return bool(re.match(pattern, num))


if __name__ == '__main__':
    # setup_logging()
    root = Tk()
    root.title("Sankey")
    frm = ttk.Frame(root, padding=10)
    frm.grid()

    # Source
    ttk.Label(frm, text="Enter your source: ").grid(column=0, row=0)
    source = ttk.Entry(frm)
    source.grid(column=1, row=0)
    # Amount
    ttk.Label(frm, text="Enter your amount: ").grid(column=0, row=1)
    amount = ttk.Entry(frm)
    amount.grid(column=1, row=1)
    # Target
    ttk.Label(frm, text="Enter your target: ").grid(column=0, row=2)
    target = ttk.Entry(frm)
    target.grid(column=1, row=2)
    # Single Source
    ttk.Label(frm, text="Enter a single source(if applicable): ").grid(column=2, row=0)
    single_source = ttk.Entry(frm)
    single_source.grid(column=3, row=0)
    single_source.bind("<KeyRelease>", on_single_source_change)
    # Single Amount
    ttk.Label(frm, text="Enter a single amount(if applicable): ").grid(column=2, row=1)
    single_amount = ttk.Entry(frm)
    single_amount.grid(column=3, row=1)
    single_amount.bind("<KeyRelease>", on_single_amount_change)
    # Single Target
    ttk.Label(frm, text="Enter a single target(if applicable): ").grid(column=2, row=2)
    single_target = ttk.Entry(frm)
    single_target.grid(column=3, row=2)
    single_target.bind("<KeyRelease>", on_single_target_change)

    # Submit Button
    sub_button = ttk.Button(frm, text="Submit", command=on_submit_clicked)
    sub_button.grid(column=0, row=3)
    # Select File Button
    file_button = ttk.Button(frm, text="Choose a file", command=on_file_clicked)
    file_button.grid(column=1, row=3)

    # Results
    result_label = scrolledtext.ScrolledText(root, wrap=WORD, width=50, height=20)
    result_label.delete('1.0', END)
    result_label.insert(END, result)
    result_label.grid(column=0, row=4)

    root.mainloop()
