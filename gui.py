import petpace_sqlite_data_merge as pp
import tkinter as tk
from tkinter import ttk, filedialog
from datetime import date, time, datetime
from tkcalendar import Calendar, DateEntry
import os

class data_converter:
    def __init__(self, master):
        self.master = master
        master.title("Doggo data reformatter")
        self.somebuttons = tk.Frame(master)
        self.somebuttons.columnconfigure(0, weight=1)
        self.somebuttons.columnconfigure(1, weight=3)

        self.somebuttons.pack(fill='x')

        self.export_file = tk.StringVar()
        self.prox_file = tk.StringVar()
        self.datatable_file = tk.StringVar()

        self.open_export_button = ttk.Button(
            self.somebuttons, text="Load export file", command=self.get_export_file)
        self.open_export_button.grid(row=0, column=0, sticky=tk.W)
        self.export_label = ttk.Label(self.somebuttons, text=self.export_file.get())
        self.export_label.grid(row=0, column=1)

        self.open_prox_button = ttk.Button(
            self.somebuttons, text="Load proximity file", command=self.get_prox_file)
        self.open_prox_button.grid(row=1, column=0, sticky=tk.W)
        self.prox_label = ttk.Label(self.somebuttons, text=self.prox_file.get())
        self.prox_label.grid(row=1, column=1)

        self.open_datatable_button = ttk.Button(
            self.somebuttons, text="Load datatable file", command=self.get_datatable_file)
        self.open_datatable_button.grid(row=2, column=0, sticky=tk.W)
        self.data_label = ttk.Label(self.somebuttons, text=self.datatable_file.get())
        self.data_label.grid(row=2, column=1)

        self.somedates = tk.Frame(master)
        self.somedates.pack(fill='x')
        ttk.Label(self.somedates, text='Choose start date').grid(row=1, column=1)
        self.start_date = DateEntry(
            self.somedates, width=12, background='darkblue',
            foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.start_date.grid(row=1, column=2)
        ttk.Label(self.somedates, text='Choose start hour').grid(row=2, column=1)
        self.starthour = ttk.Spinbox(self.somedates, width=12, from_=0, to=23)
        self.starthour.grid(row=2, column=2)


        ttk.Label(self.somedates, text='Choose end date').grid(row=3, column=1)
        self.end_date = DateEntry(self.somedates, width=12, background='darkblue',
        foreground='white', borderwidth=2,date_pattern='dd/MM/yyyy')
        self.end_date.grid(row=3, column=2)
        self.endhour = ttk.Spinbox(self.somedates, width=12, from_=0, to=23)
        ttk.Label(self.somedates, text='Choose end hour').grid(row=4, column=1)
        self.endhour.grid(row=4, column=2)

        self.somemoredates = tk.Frame(master)
        self.somemoredates.pack(fill='x')
        ttk.Label(self.somemoredates, text='Charge time #1 start').grid(row=1, column=1)
        self.charge_date_1_start = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_1_start.grid(row=1, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=1, column=3)
        self.charge_date_1_end = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_1_end.grid(row=1, column=4)
        ttk.Label(self.somemoredates, text='hour start').grid(row=2, column=1)
        self.charge_hour_1_start = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_1_start.grid(row=2, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=2, column=3)
        self.charge_hour_1_end = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_1_end.grid(row=2, column=4)

        ttk.Label(self.somemoredates, text='Charge time #2 start').grid(row=3, column=1)
        self.charge_date_2_start = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_2_start.grid(row=3, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=3, column=3)
        self.charge_date_2_end = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_2_end.grid(row=3, column=4)
        ttk.Label(self.somemoredates, text='hour start').grid(row=4, column=1)
        self.charge_hour_2_start = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_2_start.grid(row=4, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=4, column=3)
        self.charge_hour_2_end = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_2_end.grid(row=4, column=4)

        ttk.Label(self.somemoredates, text='Charge time #3 start').grid(row=5, column=1)
        self.charge_date_3_start = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_3_start.grid(row=5, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=5, column=3)
        self.charge_date_3_end = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_3_end.grid(row=5, column=4)
        ttk.Label(self.somemoredates, text='hour start').grid(row=6, column=1)
        self.charge_hour_3_start = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_3_start.grid(row=6, column=2)
        self.charge_hour_3_start.set(20)
        ttk.Label(self.somemoredates, text='end').grid(row=6, column=3)
        self.charge_hour_3_end = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_3_end.grid(row=6, column=4)

        ttk.Label(self.somemoredates, text='Charge time #4 start').grid(row=7, column=1)
        self.charge_date_4_start = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_4_start.grid(row=7, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=7, column=3)
        self.charge_date_4_end = DateEntry(self.somemoredates, width=12, background='darkblue',
        foreground='white', borderwidth=2, date_pattern='dd/MM/yyyy')
        self.charge_date_4_end.grid(row=7, column=4)
        ttk.Label(self.somemoredates, text='hour start').grid(row=8, column=1)
        self.charge_hour_4_start = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_4_start.grid(row=8, column=2)
        ttk.Label(self.somemoredates, text='end').grid(row=8, column=3)
        self.charge_hour_4_end = ttk.Spinbox(self.somemoredates, width=12, from_=0, to=23)
        self.charge_hour_4_end.grid(row=8, column=4)


        self.save_xlsx = ttk.Button(
            self.master, text="Save Excel workbook", command=self.write_xlsx)
        self.save_xlsx.pack()
        self.starthour.set(20)
        self.endhour.set(20)
        self.charge_hour_1_start.set(20)
        self.charge_hour_1_end.set(8)
        self.charge_hour_2_start.set(20)
        self.charge_hour_2_end.set(8)
        self.charge_hour_3_start.set(20)
        self.charge_hour_3_end.set(8)
        self.charge_hour_4_start.set(20)
        self.charge_hour_4_end.set(8)

    def get_export_file(self):
        self.export_file.set(filedialog.askopenfilename(title='Export file'))
        self.export_label.config(text=os.path.split(self.export_file.get())[1])

    def get_prox_file(self):
        self.prox_file.set(filedialog.askopenfilename(title='Proximity file'))
        self.prox_label.config(text=os.path.split(self.prox_file.get())[1])


    def get_datatable_file(self):
        self.datatable_file.set(filedialog.askopenfilename(title='Datatable file'))
        self.data_label.config(text=os.path.split(self.datatable_file.get())[1])


    def write_xlsx(self):
        self.xlsx_file = filedialog.asksaveasfilename(title='compiled workbook',
                    filetypes=[("xlsx files", "*.xlsx")])

        start_date = datetime.combine(self.start_date.get_date(), time(int(self.starthour.get())))
        end_date = datetime.combine(self.end_date.get_date(), time(int(self.endhour.get())))
        charge1_start_date = datetime.combine(
            self.charge_date_1_start.get_date(),
            time(int(self.charge_hour_1_start.get())))
        charge1_end_date = datetime.combine(
            self.charge_date_1_end.get_date(),
            time(int(self.charge_hour_1_end.get())))
        charge2_start_date = datetime.combine(
            self.charge_date_2_start.get_date(),
            time(int(self.charge_hour_2_start.get())))
        charge2_end_date = datetime.combine(
            self.charge_date_2_end.get_date(),
            time(int(self.charge_hour_2_end.get())))
        charge3_start_date = datetime.combine(
            self.charge_date_3_start.get_date(),
            time(int(self.charge_hour_3_start.get())))
        charge3_end_date = datetime.combine(
            self.charge_date_3_end.get_date(),
            time(int(self.charge_hour_3_end.get())))
        charge4_start_date = datetime.combine(
            self.charge_date_4_start.get_date(),
            time(int(self.charge_hour_4_start.get())))
        charge4_end_date = datetime.combine(
            self.charge_date_4_end.get_date(),
            time(int(self.charge_hour_4_end.get())))
        skip_times = [
            (charge1_start_date, charge1_end_date),
            (charge2_start_date, charge2_end_date),
            (charge3_start_date, charge3_end_date),
            (charge4_start_date, charge4_end_date)]

        pp.main(
            self.export_file.get(), self.prox_file.get(), self.datatable_file.get(),
            self.xlsx_file, start_date, end_date, skip_times=skip_times)

if __name__ == '__main__':
    root = tk.Tk()
    my_gui = data_converter(root)
    root.mainloop()
