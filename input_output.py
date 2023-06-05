import tkinter as tk
from tkinter import filedialog
import pandas as pd
from tkinter import ttk
import xlwt

class GUI:
    def __init__(self, master):
        self.master = master
        master.title("tk")
        master.geometry("700x500")

        self.left_frame = tk.Frame(master)
        self.left_frame.pack(side="left", fill="both", expand=True)

        self.right_frame = tk.Frame(master)
        self.right_frame.pack(side="right", fill="both", expand=True)

        self.button = tk.Button(master, text="Import Excel File", command=self.import_file)
        self.button.pack(pady=10)

        self.combo = ttk.Combobox(master, state="readonly", width=30)
        self.combo.pack(pady=10)

        self.add_button = tk.Button(master, text="Add Data", command=self.add_data)
        self.add_button.pack(pady=10)

        self.export_combo = ttk.Combobox(master, state="readonly", width=5)
        self.export_combo["values"] = ["txt", "csv", "xls"]
        self.export_combo.set("txt")
        self.export_combo.pack(pady=10)

        self.export_button = tk.Button(master, text="Export Data", command=self.export_data)
        self.export_button.pack(pady=10)

        self.asd = Department()

    def import_file(self):
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(file_path)
            df[["Surname", "Name"]] = df["Name"].str.split(" ", n=1, expand=True)
            sections = df["Section"].unique().tolist()
            self.combo["values"] = sections
            self.combo.set("ENGR 102 01")
            self.df = df
            
            department_dict = dict(zip(df["Id"], df["Department"]))
            self.department_dict = department_dict
            
            self.combo.bind("<<ComboboxSelected>>", self.update_data)
            self.update_data(None)


    def export_data(self):
        file_type = self.export_combo.get()
        if file_type == "txt":
            file_extension = ".txt"
            file_options = {"filetypes": [("Text files", "*.txt")]}
        elif file_type == "csv":
            raise BaseException("File type is not supported")
        elif file_type == "xls":
            file_extension = ".xls"
            file_options = {"filetypes": [("Excel files", "*.xls")]}
        else:
            raise ValueError("Invalid file type")

        file_path = filedialog.asksaveasfilename(
            title="Export Data", defaultextension=file_extension, **file_options,initialfile=self.combo.get()+file_extension
        )
        if file_path:
            if file_type == "txt":
                with open(file_path, "w") as f:
                    for item in self.listbox2.get(0, tk.END):
                        name_parts = item.split(" ")
                        surname = name_parts[0]
                        first_names = " ".join(name_parts[1:-1])
                        id_ = name_parts[-1]
                        f.write(id_+" "+first_names+" "+surname+" "+self.department_dict[int(id_)] + "\n")
                  
            elif file_type == "csv":
                pass  
            elif file_type == "xls":
                wb = xlwt.Workbook()
                ws = wb.add_sheet("Sheet1")

                ws.write(0, 0, "ID")
                ws.write(0, 1, "Name")
                ws.write(0, 2, "Department")

                for i, item in enumerate(self.listbox2.get(0, tk.END)):
                    name_parts = item.split(" ")
                    surname = name_parts[0]
                    first_names = " ".join(name_parts[1:-1])
                    id_ = name_parts[-1]
                    ws.write(i+1, 0, int(id_))
                    ws.write(i+1, 1, first_names)
                    ws.write(i+1, 2, self.department_dict[int(id_)])

                wb.save(file_path)

    def update_data(self, event):
        section = self.combo.get()
        data = self.df.loc[self.df["Section"] == section, ["Surname", "Name", "Id"]]
        self.listbox.delete(0, tk.END)

        for row in data.itertuples(index=False):
            self.listbox.insert(tk.END, " ".join([str(val) for val in row]))

        self.listbox2.delete(0, tk.END)

    def add_data(self):
        curitem = self.listbox.curselection()
        for i in curitem:
            self.listbox2.insert(tk.END, self.listbox.get(i))

    def create_listbox_widget(self):
        self.listbox = tk.Listbox(self.left_frame, font=("Courier", 10), selectmode="multiple")
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.place(relheight=0.4, relwidth=1)

        self.listbox2 = tk.Listbox(self.right_frame, font=("Courier", 10), selectmode="multiple")
        self.listbox2.pack(side="right", fill="both", expand=True)
        self.listbox2.place(relheight=0.4, relwidth=1)

class Department:
    def __init__(self):
        self.department_dict = {}

    def add_department(self, name, id):
        self.department_dict[name] = id

    def get_department_id(self, name):
        return self.department_dict.get(name)
        

if __name__ == "__main__":
    root = tk.Tk()
    gui = GUI(root)
    gui.create_listbox_widget()
    root.mainloop()

