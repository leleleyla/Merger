import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from tkinter.ttk import Progressbar, Checkbutton

# Dictionary to store exclusions for each file
exclusions = {}


# Function to select CSV files
def add_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        csv_paths.append(file_path)
        listbox.insert(tk.END, file_path)
        exclusions[file_path] = []


# Function to remove selected CSV file
def remove_csv_file():
    try:
        selected_index = listbox.curselection()[0]
        file_path = csv_paths.pop(selected_index)
        del exclusions[file_path]
        listbox.delete(selected_index)
    except IndexError:
        messagebox.showwarning("Warning", "Please select a CSV file to remove.")


# Function to search for Lethargus CSV in a selected folder
def search_lethargus():
    folder = filedialog.askdirectory(title="Select Folder")
    if not folder:
        return

    results_folder = os.path.join(folder, 'results')
    if not os.path.isdir(results_folder):
        messagebox.showwarning("Results Folder Not Found", f"Results folder not located in {folder}.")
        return

    csv_file = os.path.join(results_folder, 'Lethargus_dataframe.csv')
    if not os.path.isfile(csv_file):
        messagebox.showwarning("Lethargus CSV Not Found", f"Lethargus_dataframe.csv not found in {results_folder}.")
        return

    csv_paths.append(csv_file)
    listbox.insert(tk.END, csv_file)
    exclusions[csv_file] = []


# Function to open a window for excluding worms for a selected CSV
def open_exclude_worms_window():
    try:
        selected_index = listbox.curselection()[0]
        file_path = csv_paths[selected_index]
    except IndexError:
        messagebox.showwarning("Warning", "Please select a CSV file to exclude worms.")
        return

    def save_exclusions():
        selected_worms = [w for w, var in worm_vars.items() if var.get()]
        exclusions[file_path] = selected_worms
        exclusion_window.destroy()

    def update_checkbuttons():
        worms = pd.read_csv(file_path, header=None, nrows=1).iloc[0, 1:].tolist()
        for worm in worms:
            if worm not in worm_vars:
                var = tk.BooleanVar()
                worm_vars[worm] = var
                tk.Checkbutton(exclusion_window, text=worm, variable=var).pack(anchor='w')

    exclusion_window = tk.Toplevel(root)
    exclusion_window.title("Exclude Worms")
    tk.Label(exclusion_window, text="Select worms to exclude:").pack(padx=10, pady=5)
    worm_vars = {}
    update_checkbuttons()
    tk.Button(exclusion_window, text="Exclude Selected", command=save_exclusions).pack(pady=10)


# Function to merge CSVs into Excel with excluded worms
def merge_csv_to_excel():
    if not csv_paths:
        messagebox.showwarning("Warning", "Please add at least one CSV file.")
        return

    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")],
                                               title="Save As")
    if not output_file:
        return

    # Initialize a list for dataframes
    dfs = []

    # Read and append CSVs
    for idx, file in enumerate(csv_paths):
        df = pd.read_csv(file, header=None)  # Read without header to process all data

        if df.empty:
            messagebox.showwarning("Warning", f"CSV file {file} is empty.")
            continue

        # Drop the first column
        df = df.iloc[:, 1:]

        # Reapply headers from the first row of the original DataFrame
        headers = pd.read_csv(file, header=None, nrows=1).iloc[0, 1:]  # Read headers (skip the first column)
        df.columns = headers

        # Filter out columns based on excluded worms
        if not df.empty:
            columns_to_include = [col for col in df.columns if col not in exclusions[file]]
            df = df[columns_to_include]

        dfs.append(df)
        progress['value'] = ((idx + 1) / len(csv_paths)) * 50
        root.update_idletasks()

    # Combine all dataframes
    combined_df = pd.concat(dfs, axis=1, ignore_index=False)

    # Remove duplicated rows if any
    combined_df = combined_df.loc[~combined_df.index.duplicated(keep='first')]

    # Write combined_df to Excel
    combined_df.to_excel(output_file, index=False)
    messagebox.showinfo("Success", f"Excel file created successfully as {output_file}")
    progress['value'] = 0  # Reset progress bar


# GUI Setup
root = tk.Tk()
root.title("CSV to Excel Merger")

# Label and Entry for output file name
tk.Label(root, text="Output Excel File Name:").grid(row=0, column=0, padx=10, pady=5, sticky='E')
output_entry = tk.Entry(root, width=30)
output_entry.grid(row=0, column=1, padx=10, pady=5, sticky='W')

# Buttons
tk.Button(root, text="Add CSV", command=add_csv_file).grid(row=1, column=0, padx=10, pady=5, sticky='W')
tk.Button(root, text="Lethargus Searcher", command=search_lethargus).grid(row=2, column=0, padx=10, pady=5, sticky='W')
tk.Button(root, text="Exclude Worms", command=open_exclude_worms_window).grid(row=1, column=1, padx=10, pady=5,
                                                                              sticky='W')
tk.Button(root, text="Remove CSV", command=remove_csv_file).grid(row=1, column=2, padx=10, pady=5, sticky='W')
tk.Button(root, text="Merge", command=merge_csv_to_excel).grid(row=3, column=0, columnspan=3, pady=20)

# Listbox for CSV file paths
csv_paths = []
listbox = tk.Listbox(root, width=50, height=10)
listbox.grid(row=1, column=3, columnspan=2, padx=10, pady=5, sticky='N')

# Progress Bar
progress = Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='determinate')
progress.grid(row=4, column=0, columnspan=3, pady=10)

root.mainloop()
