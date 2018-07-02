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
        self.master = master
        self.master.option_add('*tearOff', FALSE)
        menubar = Menu(self.master)

        self.master.config(menu=menubar, width=350, height=400)
        self.master.title('Chronus')

        # Menu creation
        self.file = Menu(menubar)
        self.show = Menu(menubar)
        self.help_ = Menu(menubar)

        # Adding menus to manubar
        menubar.add_cascade(menu=self.file, label='File')
        menubar.add_cascade(menu=self.show, label='Show')
        menubar.add_cascade(menu=self.help_, label='Help')

        # Adding commands to each menu
        self.file.add_command(label='New sample', command=lambda: self.NewSample())
        self.show.add_command(label='Show list of samples', command=lambda: print('Show list of samples'))
        self.help_.add_command(label='About', command=lambda: print('Chronus'))

        # self.file.entryconfig('New sample', accelerator='ctrl+n')

    def NewSample(self):
        '''
        New window that allows the user to open a sample (with one or more analyses)
        :return:
        '''
        import ConvertFiles as CF
        NewSample_Window = Toplevel(self.master)
        NewSample_Window.config(width=350, height=400)
        NewSample_Window.title('New Sample')
        # https://stackoverflow.com/questions/11873138/how-to-set-a-child-windows-title

        cb_files_extension = ('ext')
        cb_delimmiter = ('tab')
        cb_supported_equipment = (
            'Thermo Finnigan Neptune_206F',
            'Thermo Finnigan Neptune_206MIC',
            'Spec2')

        files_extension = StringVar()
        delimmiter = StringVar()
        equipment = StringVar()
        directory = StringVar()
        directory_entry = StringVar()  # Used just to display the selected folder

        def ShowInfo_frame(self):

            frame_info = ttk.Frame(self.master, relief=RIDGE)
            frame_info.pack()

        def validate_directory(directory, window, directory_entry_widget):
            '''
            :param directory: string with the folder address to be checked
            :param directory_entry_widget: Entry widget used just to display the selected folder
            :param window: parent window which should be focused
            :return: a string with a validated path
            '''
            import os
            if sys.version_info[0] < 3:
                from Tkinter import messagebox
            else:
                from tkinter import messagebox

            if os.path.exists(directory) == False:
                messagebox.showwarning('Warning', 'Please, select a folder!')
                SelectFolder(directory_entry_widget)
                window.focus()
            else:
                window.focus()
                return directory

        def SelectFolder(directory_entry_widget):
            '''
            Asks the user for a folder path.
            :param directory_entry_widget: Entry widget used just to display the selected folder
            :return: a string with the path of the folder selected by the user
            '''
            if sys.version_info[0] < 3:
                from Tkinter import tkFileDialog as fDialog
                from Tkinter import ttk
            else:
                from tkinter import filedialog as fDialog
                from tkinter import ttk

            path = validate_directory(fDialog.askdirectory(), NewSample_Window, directory_entry_widget)
            directory_entry_widget.config(text=path)
            return path

        def Ok_button_action():
            self.Sample = CF.SampleData.LoadFiles(directory.get(),
                                                  files_extension.get(),
                                                  delimiter=delimmiter.get(),
                                                  equipment=equipment.get())

            ShowInfo_frame()

        def New_Sample_Object(folderpath, extension, delimiter, equipment):
            sample_object

        combobox1 = ttk.Combobox(NewSample_Window, textvariable=files_extension, values=cb_files_extension)
        # combobox1.config(values=cb_files_extension)
        combobox2 = ttk.Combobox(NewSample_Window, textvariable=delimmiter, values=cb_delimmiter)
        # combobox2.config(values=cb_delimmiter)
        combobox3 = ttk.Combobox(NewSample_Window, textvariable=equipment, values=cb_supported_equipment)
        # combobox3.config(values=cb_supported_equipment)

        label1 = Label(NewSample_Window, text='Files extension')
        label2 = Label(NewSample_Window, text='Delimmiter')
        label3 = Label(NewSample_Window, text='Equipment')
        label4 = Label(NewSample_Window)

        button_SelectFolder = Button(NewSample_Window,
                                     text='Select folder',
                                     command=lambda: directory.set(SelectFolder(label4)))
        button_Ok = Button(NewSample_Window,
                           text='Ok',
                           command=Ok_button_action)

        button_SelectFolder.grid(row=0, column=0)
        combobox1.grid(row=1, column=1, sticky=W)
        combobox2.grid(row=2, column=1, sticky=W)
        combobox3.grid(row=3, column=1, sticky=W)
        label1.grid(row=1, column=0)
        label2.grid(row=2, column=0)
        label3.grid(row=3, column=0)
        label4.grid(row=0, column=1)
        button_Ok.grid(row=4, column=0, rowspan=2)

        int_padding = 5
        ext_padding = 10

        for child in NewSample_Window.winfo_children():
            child.grid_configure(padx=ext_padding, pady=ext_padding, ipadx=int_padding, ipady=int_padding)


def main():
    root = Tk()
    app = Chronus_GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
