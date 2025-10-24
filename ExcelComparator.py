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

        list_frame = tk.Frame(master)
        list_frame.pack(fill="both", expand=False, padx=10, pady=(10, 5))

        self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=8)
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.drop)
        # self.label = tk.Label(master, text="Drag and drop Excel files here\nor click 'Add Files'", bg="#f0f0f0")
        # self.label.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
 
        # self.label.config(highlightbackground="grey", highlightthickness=2)

        # self.label.drop_target_register(DND_FILES)
        # self.label.dnd_bind('<<Drop>>', self.drop)

        # context menu
        self.menu = tk.Menu(master, tearoff=0)
        self.menu.add_command(label="Remove selected", command=self.remove_selected)
        self.menu.add_command(label="Clear all", command=self.clear_files)

        # bind right click
        self.listbox.bind("<Button-3>", self._on_right_click)   # Windows, Linux
        self.listbox.bind("<Button-2>", self._on_right_click)   # Mac sometimes

        # controls
        btn_frame = tk.Frame(master)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.add_btn = tk.Button(btn_frame, text="Add Files", command=self.add_files)
        self.add_btn.pack(side="left", padx=(0, 6))

        self.remove_btn = tk.Button(btn_frame, text="Remove selected", command=self.remove_selected)
        self.remove_btn.pack(side="left", padx=(0, 6))

        self.delete_files_btn = tk.Button(btn_frame, text="Clear Files", command=self.clear_files)
        self.delete_files_btn.pack(side="left", padx=(0, 6))

        self.compare_btn = tk.Button(master, text="Compare", command=self.compare)
        self.compare_btn.pack(pady=5)


    def update_listbox(self): 
        self.listbox.delete(0, tk.END)
        for p in self.files:
            self.listbox.insert(tk.END, p)

    def drop(self, event):
        #self.label.config(highlightbackground="green", highlightthickness=3)       #TODO: upgrade visual feedback
        #self.master.after(250, lambda: self.label.config(highlightbackground="gray", highlightthickness=2))

        paths = self.master.tk.splitlist(event.data)
        added = False
        for p in paths:
            if p.endswith('.xlsx') and p not in self.files:
                self.files.append(p)
                added = True
        if added:
            self.update_listbox()
        return "break"

    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        added = False
        for p in paths:
            if p.endswith('.xlsx') and p not in self.files:
                self.files.append(p)
                added = True
        if added:
            self.update_listbox()
        #self.files.extend(paths)
        #self.label.config(text="\n".join(self.files))

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        try:
            del self.files[idx]
        except IndexError:
            pass
        self.update_listbox()

    def clear_files(self):
        if messagebox.askyesno("Confirm", "Clear all files?"):
            self.files.clear()
            self.update_listbox()
    def _on_right_click(self, event):
        try:
            index = self.listbox.nearest(event.y)
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(index)
        except Exception:
            pass
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()


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