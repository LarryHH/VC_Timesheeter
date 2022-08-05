import os, sys
import os.path
from os import path
from pathlib import Path
from tkinter import Tk, Canvas, Entry, Button, PhotoImage, StringVar
import tkinter.font
import webbrowser

from datetime import date, datetime, timedelta
from timesheeter import Timesheeter

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

ASSETS_PATH = os.path.join(os.path.abspath(application_path), 'assets\\')
OUTPUT_PATH = os.path.dirname(sys.executable)
BUTTON_Y_OFFSET = 36

class WindowForm():
    
    @staticmethod
    def relative_to_assets(path: str) -> Path:
        return os.path.join(ASSETS_PATH, path)

    @staticmethod
    def first_monday_calc(year):
        """calculate when the first Monday of the year is to correctly align fortnightly period"""
        days = range(1,32)
        for day in days:
            if datetime(year, 1, day).weekday() == 0:
                first_monday = datetime(year, 1, day)
                next_monday = first_monday + timedelta(days=7)
                return (first_monday.date(), next_monday.date())

    def __init__(self, window, file, wb, date_sheet_name, date_cell) -> None:
        self.window = window
        self.canvas = Canvas(
            self.window,
            bg = "#E2D7D4",
            height = 700,
            width = 1000,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        self.canvas.place(x = 0, y = 0)

        self.bg_left = PhotoImage(
            file=WindowForm.relative_to_assets("bg_left.png"))
        self.canvas.create_image(
            200,
            350,
            image=self.bg_left
        )

        self.canvas.create_rectangle(
            400.0,
            0.0,
            1000.0,
            700.0,
            fill="#FBF9F9",
            outline="")

        self.window.title('VC Timesheeter')
        self.window.geometry("1000x700")
        self.window.configure(bg = "#E2D7D4")
        self.entry_image = PhotoImage(
            file=WindowForm.relative_to_assets("entry.png"))

        self.fill_canvas_left()
        res = self.prefill_form(file)
        self.fill_canvas_right(res, wb, date_sheet_name, date_cell)

        self.window.resizable(False, False)
        self.window.mainloop()

    def create_entry_image(self, x, y):
        self.canvas.create_image(
            x, y, image=self.entry_image
        )

    def create_entry_text_label(self, x, y, text):
        return self.canvas.create_text(
            x,
            y,
            anchor="nw",
            text=text,
            fill="#464749",
            font=("Arial Bold", 16 * -1)
        )

    @staticmethod
    def find_valid_output_path(output_location, year):
        i = 1
        while(path.exists(output_location)):
            if path.exists(output_location):
                output_location = OUTPUT_PATH + f"\\{year}_{i}"
                i += 1
        return output_location

    def prefill_form(self, file):
        template_file = file
        year = date.today().year
        first_monday, next_monday = WindowForm.first_monday_calc(year)
        output_location = OUTPUT_PATH + f"\\{year}"
        output_location = WindowForm.find_valid_output_path(output_location, year)

        return template_file, year, first_monday, next_monday, output_location

    def fill_canvas_left(self):
        self.canvas.create_text(
            84.0,
            116.0,
            anchor="nw",
            text="VIM&CO",
            fill="black",
            font=("EngraversGothic BT", 70 * -1)
        )
        self.canvas.create_text(
            84.0,
            186.0,
            anchor="nw",
            text="Timesheeter",
            fill="#464749",
            font=("EngraversGothic BT", 45 * -1)
        )
        self.canvas.create_text(
            34.0,
            304.0,
            anchor="nw",
            text="Produces a full, blank set of timesheets",
            fill="#464749",
            font=("Varela Round", 20 * -1)
        )
        self.canvas.create_text(
            34.0,
            340.0,
            anchor="nw",
            text="derived from an Excel template",
            fill="#464749",
            font=("Varela Round", 20 * -1)
        )
        self.canvas.create_text(
            34.0,
            377.0,
            anchor="nw",
            text="for the year of choice.",
            fill="#464749",
            font=("Varela Round", 20 * -1)
        )
        self.canvas.create_text(
            42.0,
            419.0,
            anchor="nw",
            text="◉ Enter <year> and <first monday>",
            fill="#464749",
            font=("Varela Round", 20 * -1)
        )
        self.canvas.create_text(
            42.0,
            450.0,
            anchor="nw",
            text="◉ Check if other fields are correct",
            fill="#464749",
            font=("Varela Round", 20 * -1)
        )
        self.canvas.create_text(
            26.0,
            658.0,
            anchor="nw",
            text="Larry Huynh, 2022",
            fill="#E2D7D4",
            font=("Varela Round", 16 * -1)
        )
        def cb_github_link():
            webbrowser.open_new(r"https://github.com/LarryHH/VC_Timesheeter")
        self.github_link = Button(
            borderwidth=0,
            highlightthickness=0,
            bg='#795c5f',
            activebackground='#795c5f',
            activeforeground='#795c5f',
            command=cb_github_link
        )        
        self.github_icon = PhotoImage(file=WindowForm.relative_to_assets("github_logo.png")) # make sure to add "/" not "\"
        self.github_link.config(image=self.github_icon)
        self.github_link.place(
            x=165.0,
            y=648.0
        )

    def fill_canvas_right(self, res, wb, date_sheet_name, date_cell):    
        template_file, year, first_monday, next_monday, output_location = res
        
        self.create_entry_image(700.0, 119)
        self.create_entry_image(700.0, 202)
        self.create_entry_image(700.0, 285)
        self.create_entry_image(700.0, 368)
        self.create_entry_image(700.0, 451)
        self.create_entry_image(700.0, 534)

        def cb_instruction_link():
            webbrowser.open_new(r"https://github.com/LarryHH/VC_Timesheeter/blob/master/docs/instructions.md")
        self.instruction_link = Button(
            borderwidth=0,
            highlightthickness=0,
            bg='#fbf9f9',
            activebackground='#fbf9f9',
            activeforeground='#fbf9f9',
            command=cb_instruction_link
        )        
        self.instruction_icon = PhotoImage(file=WindowForm.relative_to_assets("doubts-button_3.png")) # make sure to add "/" not "\"
        self.instruction_link.config(image=self.instruction_icon)
        self.instruction_link.place(
            x=950.0,
            y=20.0
        )

        self.entry_template = Entry(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            font=("VarelaRound Regular", 16 * -1)
        )
        self.entry_template.place(
            x=490,
            y=86+BUTTON_Y_OFFSET,
            width=400,
            height=20,
        )
        self.entry_template.insert('end', template_file)
        self.entry_template.configure(state='disabled')
        
        self.fm_options = tkinter.StringVar(self.window)
        self.fm_options.set(first_monday)
        self.entry_monday = tkinter.OptionMenu(self.window, self.fm_options, *[first_monday, next_monday])
        self.entry_monday.config(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            activebackground='#fbf9f9',
            font=("VarelaRound Regular", 16 * -1)
        )
        self.entry_monday.place(
            x=490,
            y=252.0+BUTTON_Y_OFFSET,
            width=400,
            height=20
        )

        self.entry_sheet = Entry(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            font=("VarelaRound Regular", 16 * -1)
        )
        self.entry_sheet.place(
            x=490,
            y=335.0+BUTTON_Y_OFFSET,
            width=400,
            height=20
        )
        self.entry_sheet.insert('end', date_sheet_name)
        
        self.entry_cell = Entry(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            font=("VarelaRound Regular", 16 * -1)
        )
        self.entry_cell.place(
            x=490,
            y=418.0+BUTTON_Y_OFFSET,
            width=400,
            height=20
        )
        self.entry_cell.insert('end', date_cell)

        self.entry_output = Entry(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            font=("VarelaRound Regular", 16 * -1)
        )
        self.entry_output.place(
            x=490,
            y=501.0+BUTTON_Y_OFFSET,
            width=400,
            height=20
        )
        self.entry_output.insert('end', output_location)

        sv_year = StringVar()
        sv_year.trace("w", lambda name, index, mode, sv=sv_year: self.update_year_callbacks(sv))
        self.entry_year = Entry(
            bd=0,
            bg="#FFFFFF",
            highlightthickness=0,
            font=("VarelaRound Regular", 16 * -1),
            textvariable=sv_year
        )
        self.entry_year.place(
            x=490,
            y=169+BUTTON_Y_OFFSET,
            width=400,
            height=20
        )
        self.entry_year.insert('end', year)

        self.entry_template_label = self.create_entry_text_label(490, 94.8326416015625, "Template File"),
        self.entry_year_label = self.create_entry_text_label(490, 176.0, "Year"),
        self.entry_date_sheet_label = self.create_entry_text_label(490, 342.0, "Date Sheet"),
        self.entry_date_cell_label = self.create_entry_text_label(490, 425.0, "Date Cell"),
        self.entry_output_label = self.create_entry_text_label(490, 508.0, "Output Location"),
        self.entry_fm_label = self.create_entry_text_label(490, 259.0, "First Monday")

        self.canvas.create_text(
            480.0,
            41.0,
            anchor="nw",
            text="Verify Fields:",
            fill="#464749",
            font=("VarelaRound Bold", 24 * -1)
        )

        self.button_go = Button(
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.exec_timersheeter(wb),
            relief="flat",
            text='GO!',
            bg='#efc4b8',
            fg='white',
            activebackground='#FDA48B',
            activeforeground='#FFFFFF'
        )
        self.button_go['font'] = tkinter.font.Font(family='VarelaRound Regular', size=30, weight='bold')
        self.button_go.place(
            x=475.0,
            y=604.0,
            width=450.0,
            height=50.0
        )

    def restore_labels(self):
        self.canvas.itemconfig(self.entry_template_label, text='Template File', fill='black')
        self.canvas.itemconfig(self.entry_year_label, text='Year', fill='black')
        self.canvas.itemconfig(self.entry_date_sheet_label, text='Date Sheet', fill='black')
        self.canvas.itemconfig(self.entry_date_cell_label, text='Date Cell', fill='black')
        self.canvas.itemconfig(self.entry_output_label, text='Output Location', fill='black')
        self.canvas.itemconfig(self.entry_fm_label, text='First Monday', fill='black')

    def update_year_callbacks(self, sv):
        # output_fp and fm
        self.entry_output.delete(0, 'end')
        year = self.entry_year.get()

        output_location = os.path.dirname(sys.executable) + f"\\{year}"
        output_location = WindowForm.find_valid_output_path(output_location, year)
        self.entry_output.insert('end', output_location)

        first_monday, next_monday = WindowForm.first_monday_calc(int(year))
        
        self.fm_options.set(first_monday)

        self.entry_monday['menu'].delete(0, 'end')

        for d in [first_monday, next_monday]:
            self.entry_monday['menu'].add_command(label=d, command= lambda d=d: self.fm_options.set(d))

    def button_processing(self):
        self.button_go['text'] = 'Processing...'
        self.button_go['font'] = tkinter.font.Font(family='VarelaRound Regular', size=24, weight='bold')
        self.button_go['cursor'] = 'X_cursor'
        self.button_go["state"] = "disabled"
        self.button_go.configure(bg="#d4d4d4", fg="black")
    
    def button_done(self, ts):
        while not ts.done:
            continue
        self.button_go['text'] = 'Done! Exit now'
        self.button_go['cursor'] = 'arrow'
        self.button_go["state"] = "disabled"
        self.button_go.configure(bg='#efc4b8', fg='#FFFFFF', disabledforeground='#FFFFFF')

    def exec_timersheeter(self, wb):
        ts = Timesheeter((wb, self.entry_template.get(), self.entry_year.get(), self.fm_options.get(), self.entry_sheet.get(), self.entry_cell.get(), self.entry_output.get()))
        errs = ts.is_good_to_go()
        if errs:
            for e in list(set(errs)):
                if e == 'FileExistsError' or e == 'OSError':
                    msg = f"Folder: {os.path.basename(os.path.normpath(self.entry_output.get()))} already exists in output location."
                    self.canvas.itemconfig(self.entry_output_label, text=msg, fill='red')
                if e == 'FileNotFoundError':
                    msg = f"No such existing folder."
                    self.canvas.itemconfig(self.entry_output_label, text=msg, fill='red')
        else:
            self.restore_labels()
            self.button_processing()

            self.canvas.after(500, lambda: ts.exec())
            self.canvas.after(500, lambda: self.button_done(ts))

if __name__ == "__main__":
    form = WindowForm(Tk(), None)
    form.window.mainloop()