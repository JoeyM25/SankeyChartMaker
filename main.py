import tkinter
import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext


def sankey_format(sheet, source, amount, target):
    df = pd.read_excel(sheet)

    arr = []

    for i in range(len(df)):
        skip = False
        source1 = df.loc[[i], [source]].to_string(index=False, header=False)
        amount1 = df.loc[[i], [amount]].to_string(index=False, header=False)
        target1 = df.loc[[i], [target]].to_string(index=False, header=False)

        if source1 == "NaN" or amount1 == "NaN" or target1 == "NaN":
            skip = True

        if not skip:
            text = f"{source1}[{amount1}]{target1}\n"
            arr.append(text)

    return ''.join(arr)


def on_click():
    result = ""
    get_source = source.get()
    get_amount = amount.get()
    get_target = target.get()
    get_single_source = single_source.get()
    get_single_amount = single_amount.get()
    get_single_target = single_target.get()

    if not get_single_amount or not get_single_target or not get_single_source:
        result = sankey_format("testData.xlsx", get_source, get_amount, get_target)

    if not get_target or not get_amount or not get_source:
        result = f"{get_single_source}[{get_single_amount}]{get_single_target}\n"

    result_label.delete('1.0', END)
    result_label.insert(END, result)


if __name__ == '__main__':
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
    # Single Amount
    ttk.Label(frm, text="Enter a single amount(if applicable): ").grid(column=2, row=1)
    single_amount = ttk.Entry(frm)
    single_amount.grid(column=3, row=1)
    # Single Target
    ttk.Label(frm, text="Enter a single target(if applicable): ").grid(column=2, row=2)
    single_target = ttk.Entry(frm)
    single_target.grid(column=3, row=2)

    # Submit Button
    button = ttk.Button(frm, text="Submit", command=on_click)
    button.grid(column=0, row=3)
    # Results
    result_label = scrolledtext.ScrolledText(root, wrap=WORD, width=50, height=20)
    result_label.grid(column=0, row=4)

    root.mainloop()
