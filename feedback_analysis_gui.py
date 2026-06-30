
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import traceback

# Import your existing analysis script
# Make sure feedback_analysis.py is in the SAME folder as this file.
try:
    import feedback_analysis
except Exception as e:
    raise SystemExit(
        "Could not import feedback_analysis.py.\n"
        "Place feedback_analysis.py in the same folder as this launcher.\n\n"
        f"Error: {e}"
    )

def choose_file():
    file = filedialog.askopenfilename(
        title="Select Reviews File",
        filetypes=[
            ("Excel files","*.xlsx *.xls"),
            ("CSV files","*.csv"),
            ("All files","*.*")
        ]
    )
    if file:
        entry.delete(0, tk.END)
        entry.insert(0, file)

def run_analysis():
    infile = entry.get().strip()
    if not infile:
        messagebox.showerror("Error", "Please choose an Excel or CSV file.")
        return

    outdir = Path(infile).parent / "feedback_analysis_output"

    try:
        feedback_analysis.run_analysis(infile, outdir)
        messagebox.showinfo(
            "Done",
            f"Analysis complete!\n\nReport saved to:\n{outdir}"
        )
    except Exception:
        messagebox.showerror("Analysis Failed", traceback.format_exc())

root = tk.Tk()
root.title("Customer Feedback Analyzer")
root.geometry("600x180")

tk.Label(root,text="Review File").pack(anchor="w", padx=10, pady=(12,2))

frame = tk.Frame(root)
frame.pack(fill="x", padx=10)

entry = tk.Entry(frame)
entry.pack(side="left", fill="x", expand=True)

tk.Button(frame,text="Browse",command=choose_file).pack(side="left", padx=5)

tk.Button(
    root,
    text="Run Analysis",
    font=("Arial",11,"bold"),
    command=run_analysis,
    height=2
).pack(pady=20)

root.mainloop()
