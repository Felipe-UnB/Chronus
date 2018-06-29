import sys

if sys.version_info[0] < 3:
    from Tkinter import *
    from Tkinter import ttk
else:
    from tkinter import *
    from tkinter import ttk


class Chronus_GUI():
    '''
    Chronus graphical user interface
    '''

    def __init__(self, master):
        master.option_add('*tearOff', FALSE)
        menubar = Menu(master)

        master.config(menu=menubar, width=350, height=400)
        master.title('Chronus')

        # Menu creation
        self.file = Menu(menubar)
        self.show = Menu(menubar)
        self.help_ = Menu(menubar)

        # Adding menus to manubar
        menubar.add_cascade(menu=self.file, label='File')
        menubar.add_cascade(menu=self.show, label='Show')
        menubar.add_cascade(menu=self.help_, label='Help')

        # Adding commands to each menu
        self.file.add_command(label='New sample', command=lambda: self.NewSample(master))
        self.show.add_command(label='Show list of samples', command=lambda: print('Show list of samples'))
        self.help_.add_command(label='About', command=lambda: print('Chronus'))

        self.file.entryconfig('New sample', accelerator='Ctrl+N')

    def NewSample(self, master):
        NewSample_Window = Toplevel(master)
        NewSample_Window.config(width=350, height=400)

        cb_files_extension = ('ext')
        cb_delimmiter = ('tab')
        cb_supported_equipment = (
            'Thermo Finnigan Neptune_206F',
            'Thermo Finnigan Neptune_206MIC',
            'Spec2')

        def SelectFolder():
            if sys.version_info[0] < 3:
                from Tkinter import tkFileDialog as fDialog
                from Tkinter import ttk
            else:
                from tkinter import filedialog as fDialog
                from tkinter import ttk

            Directory = fDialog.askdirectory()

        button_SelectFolder = Button(NewSample_Window, text='Select folder', command=SelectFolder)

        files_extension = StringVar()
        delimmiter = StringVar()
        equipment = StringVar()

        combobox1 = ttk.Combobox(NewSample_Window, textvariable=files_extension)
        combobox1.config(values=cb_files_extension)
        combobox2 = ttk.Combobox(NewSample_Window, textvariable=delimmiter)
        combobox2.config(values=cb_delimmiter)
        combobox3 = ttk.Combobox(NewSample_Window, textvariable=equipment)
        combobox3.config(values=cb_supported_equipment)

        label1 = Label(NewSample_Window, text='Files extension')
        label2 = Label(NewSample_Window, text='Delimmiter')
        label3 = Label(NewSample_Window, text='Equipment')

        button_SelectFolder.grid(row=0, column=0)
        combobox1.grid(row=1, column=0)
        combobox2.grid(row=2, column=0)
        combobox3.grid(row=3, column=0)
        label1.grid(row=1, column=1)
        label2.grid(row=2, column=1)
        label3.grid(row=3, column=1)


def main():
    root = Tk()
    app = Chronus_GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
