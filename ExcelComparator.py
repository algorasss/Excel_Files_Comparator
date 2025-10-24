import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment

class ExcelDiffTool:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Diff Tool")
        self.master.geometry("500x300")

        self.files = []

        self.label = tk.Label(master, text="Drag and drop Excel files here\nor click 'Add Files'", bg="#f0f0f0")
        self.label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
 
        self.label.config(highlightbackground="grey", highlightthickness=2)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.drop)

        self.delete_files_btn = tk.Button(master, text="Clear Files", command=lambda: [self.files.clear(), self.label.config(text="")])
        self.delete_files_btn.pack(pady=5)

        self.add_btn = tk.Button(master, text="Add Files", command=self.add_files)
        self.add_btn.pack(pady=5)

        self.compare_btn = tk.Button(master, text="Compare", command=self.compare)
        self.compare_btn.pack(pady=5)


    def drop(self, event):
        self.label.config(highlightbackground="green", highlightthickness=3)       #TODO: upgrade visual feedback
        self.master.after(250, lambda: self.label.config(highlightbackground="gray", highlightthickness=2))

        paths = self.master.tk.splitlist(event.data)
        self.files.extend([p for p in paths if p.endswith('.xlsx')])
        self.label.config(text="\n".join(self.files))

    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        self.files.extend(paths)
        self.label.config(text="\n".join(self.files))


    def compare(self):
        if len(self.files) < 2:
            messagebox.showwarning("Warning", "Add at least 2 Excel files")
            return

        dfs = []
        for f in self.files:
            try:
                df = pd.read_excel(f, dtype=str, header=None)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {f}\n{e}")
                return
            dfs.append(df)

        # determine max rows and cols
        max_rows = max(df.shape[0] for df in dfs)
        max_cols = max(df.shape[1] for df in dfs)

        # create result workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Diff"

        red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        green_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")

        for r in range(max_rows):
            for c in range(max_cols):
                values = []
                for df in dfs:
                    try:
                        val = df.iat[r, c]
                        if pd.isna(val):
                            val = "—"
                    except:
                        val = "—"
                    values.append(str(val))
                cell = ws.cell(row=r+1, column=c+1)
                if all(v == values[0] for v in values):
                    cell.value = values[0]
                    cell.fill = green_fill
                else:
                    cell.value = "\n".join(f"{i+1}: {v}" for i,v in enumerate(values))
                    cell.alignment = Alignment(wrap_text=True)
                    cell.fill = red_fill

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            wb.save(save_path)
            messagebox.showinfo("Done", f"Diff saved to {save_path}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ExcelDiffTool(root)
    root.mainloop()