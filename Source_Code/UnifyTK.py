from AgilentUnify import AgilentUnify
from WatersUnify import WatersUnify
import tkinter as Tk
from tkinter import filedialog
import csv


class UnifyTK(Tk.Frame):
    """Runs the tk window that allows the user to interact with UnifyMB.

    Attributes
    ----------
    filename : string
        the filename of the file to unify.
    csv_file_in_list_format : list
        the csv file to unify, as a list of lists, produced by csv.reader
    files_to_unify_directory : string
        the directory that the files to unify are in.
    browse_directory_button : Tk.Button
        opens up to the directory the files to unify are in.
    unify_tk_console_log_view : Tk.text
        text box for displaying error messages and log messages in the GUI.
    """

    def __init__(self, parent, **kwargs):
        """
        Parameters
        ----------
        parent : __init__.py
            a TkInter frame
        kwargs
            no additional parameters are passed.
        """

        Tk.Frame.__init__(self, parent, **kwargs)
        self.filename = ""
        self.csv_file_in_list_format = []
        self.files_to_unify_directory = r"T:\ANALYST WORK FILES\Peter\CrystalMB\CSVFilesToUnify"
        self.browse_directory_button = Tk.Button(self,
                                                 text='Select File',
                                                 command=lambda: self.unify_tk_controller())
        self.unify_tk_console_log_view = Tk.Text(self, height=25, width=60)
        # Tk layout
        self.browse_directory_button.grid(row=0, column=0, sticky=Tk.W, padx=5, pady=5)
        self.unify_tk_console_log_view.grid(row=1, column=0, sticky=Tk.W, padx=5)
        self.unify_tk_console_log_view.config(state=Tk.DISABLED)

    def unify_tk_controller(self):
        """main function for the class. """

        self.browse_file_directory()
        self.open_and_read_in_csv_file()
        self.unify_pathway_decider()
        # clearing the parameters so the button can be used again with a new file
        self.filename = ''
        self.csv_file_in_list_format = []

    def browse_file_directory(self):
        """opens up to the directory the files to unify are in. """

        self.filename = filedialog.askopenfilename(initialdir=self.files_to_unify_directory)

    def open_and_read_in_csv_file(self):
        """Opens the csv file at the filename, reads it line by line, appends to self.csv_file_in_list_format.

         Raises
         ------
         UnicodeDecodeError
            file is not being recognized as a csv file.
        """

        try:
            with open(self.filename, newline='') as csv_file:
                csv_file_reader = csv.reader(csv_file)
                for item in csv_file_reader:
                    self.csv_file_in_list_format.append(item)
            self.write_log_to_text_box("csv file successfully read. \n\n", True)
        except UnicodeDecodeError:
            self.write_log_to_text_box("This file is not being recognized as a csv file. \n\n", True)

    def unify_pathway_decider(self):
        """Decides which of the two pathways to send the csv down - Waters, or Agilent.

        There are two different software packages controlling agilent ICP/MS instruments in the lab, so we have
        an optional keyword argument if we are taking data from the older of the two ICP/MS systems. The fields
        and data format provided are slightly different, and need separate but similar handling. """

        keyword_check = self.csv_file_in_list_format[0][1]
        if keyword_check == 'version':
            WatersUnify(self.csv_file_in_list_format).waters_unify_controller()
            self.write_log_to_text_box("Waters csv file detected.\n" +
                                       "sending down WatersUnify pathway. \n")
        elif keyword_check == 'Sample Type':
            AgilentUnify(self.csv_file_in_list_format).agilent_unify_controller()
            self.write_log_to_text_box("Agilent csv file detected (Melanie ICP/MS)\n" +
                                       "sending down AgilentUnify pathway. \n")
        elif keyword_check == 'Type':
            AgilentUnify(self.csv_file_in_list_format).agilent_unify_controller(old_icp=True)
            self.write_log_to_text_box("Agilent csv file detected (Harry ICP/MS)\n" +
                                       "sending down AgilentUnify pathway. \n")
        else:
            self.write_log_to_text_box("No valid unify pathway detected for this csv file. \n")

    def write_log_to_text_box(self, passed_text, with_delete=False):
        """clears any text in EasyTraxMakerLog, and then writes the passed text to the Text box."""

        self.unify_tk_console_log_view.config(state=Tk.NORMAL)
        if with_delete:
            self.unify_tk_console_log_view.delete('1.0', Tk.END)
        self.unify_tk_console_log_view.insert(Tk.END, passed_text)
        self.unify_tk_console_log_view.config(state=Tk.DISABLED)


root = Tk.Tk()
root.geometry('495x450')
UnifyTK(root, height=450, width=495).grid()
root.mainloop()


