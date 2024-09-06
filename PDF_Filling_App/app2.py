import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def compare_excel_files(file1, file2):
    try:
        # Load the Excel files into DataFrames
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Check if the shapes of the DataFrames are the same
        if df1.shape != df2.shape:
            return False, f"Files have different shapes: {df1.shape} vs {df2.shape}"

        # Compare the two DataFrames
        comparison = df1 == df2

        # Check if all values are equal
        if comparison.all().all():
            return True, "All values in the files are equal."
        else:
            # Find the exact locations of differences
            differences = []
            for row in range(comparison.shape[0]):
                for col in range(comparison.shape[1]):
                    if not comparison.iloc[row, col]:
                        differences.append(f"Difference at row {row + 1}, column {col + 1}: {df1.iloc[row, col]} vs {df2.iloc[row, col]}")

            return False, differences
    except Exception as e:
        return False, str(e)


def select_file():
    file_path = filedialog.askopenfilename(title="Select an Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path


def on_compare_button_click():
    file1 = select_file()
    if not file1:
        messagebox.showwarning("Warning", "No file selected!")
        return

    file2 = select_file()
    if not file2:
        messagebox.showwarning("Warning", "No file selected!")
        return

    # Perform the comparison
    are_equal, result = compare_excel_files(file1, file2)

    # Show the result
    if are_equal:
        messagebox.showinfo("Result", result)
    else:
        if isinstance(result, list):
            differences = "\n".join(result)
        else:
            differences = result
        messagebox.showinfo("Result", f"Files are different:\n\n{differences}")


# Set up the main application window
app = tk.Tk()
app.title("Excel File Comparator")

# Create a button to compare files
compare_button = tk.Button(app, text="Compare Excel Files", command=on_compare_button_click)
compare_button.pack(pady=20)

# Run the application
app.mainloop()
