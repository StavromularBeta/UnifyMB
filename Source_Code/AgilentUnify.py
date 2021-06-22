import csv
import os.path
import errno
import xlsxwriter


class AgilentUnify:
    """Converts Agilent csv data to the Unified Excel Format.

    Attributes
    ----------
    csv_file_in_list_of_list_format : list(list)
        csv file to be converted in a list of lists.
    csv_file_in_list_of_list_format_condensed : list(list)
        csv file to be converted in a list of lists. Any extra fields not required are removed.
    required_fields_index_dictionary : dict
        dictionary of required file fields, and their indexes. 0's are placeholders, indexes are found later.
    main_file_name : string
        the name of the excel file to be generated. Same as the batch name you'd use for the targetlynx file.
    """

    def __init__(self, csv_file_in_list_of_list_format):
        """
        Parameters
        ----------
        csv_file_in_list_of_list_format : list(list)
            csv file to be converted in a list of lists.
        """
        self.csv_file_in_list_of_list_format = csv_file_in_list_of_list_format
        self.csv_file_in_list_of_list_format_condensed = [['data source',
                                                           'filename',
                                                           'create date',
                                                           'create time',
                                                           'sample name',
                                                           'sample type',
                                                           'analyte name',
                                                           'analyte concentration',
                                                           'percent recovery']]
        self.required_fields_index_dictionary = {'Sample Name': 0,
                                                 'Sample Type': 0,
                                                 'Date and Time Acquired': 0,
                                                 'Data File Name': 0,
                                                 'Batch Name': 0,
                                                 'Analyte': 0,
                                                 'Mass': 0,
                                                 'Concentration': 0,
                                                 'Units': 0,
                                                 'Tune Step': 0,
                                                 }
        self.required_fields_index_dictionary_old_icp = {'Solution Label': 0,
                                                         'Type': 0,
                                                         'Date Time': 0,
                                                         }
        self.main_file_name = ""

    def agilent_unify_controller(self, old_icp=False):
        """The main controller function for AgilentUnify. """
        if old_icp:
            self.find_indexes_of_required_fields(old_icp)
            self.create_condensed_csv_file_with_only_relevant_fields(old_icp)
        else:
            self.find_indexes_of_required_fields()
            self.create_condensed_csv_file_with_only_relevant_fields()
        self.generate_excel_files()
        # have the option of generating excel files with xlsxwriter, can format somewhat
        # or make un-formatted csv files.
        # self.generate_csv_files()

    def find_indexes_of_required_fields(self, old_icp=False):
        """finds the indexes of the fields we need to create our unified excel format.

        This is somewhat unnecessary, because with Agilent instruments, we can always trim fields we don't need, and
        we can put fields in whatever order we want, so we could in theory already know these indexes. However, having
        this function in place allows us to be lazy - we don't need to care about the order of fields, and we can ignore
        any extra fields. """
        if old_icp:
            required_field_index_counter = 0
            for item in self.csv_file_in_list_of_list_format[0]:
                if item in self.required_fields_index_dictionary_old_icp.keys():
                    # then it is a required field
                    # add the index we are at to the dictionary
                    self.required_fields_index_dictionary_old_icp[item] = required_field_index_counter
                required_field_index_counter += 1
        else:
            required_field_index_counter = 0
            for item in self.csv_file_in_list_of_list_format[0]:
                if item in self.required_fields_index_dictionary.keys():
                    # then it is a required field
                    # add the index we are at to the dictionary
                    self.required_fields_index_dictionary[item] = required_field_index_counter
                required_field_index_counter += 1

    def create_condensed_csv_file_with_only_relevant_fields(self, old_icp=False):
        """takes only the required fields from the full csv file, and creates a condensed list of lists.

        currently the only Agilent instruments we are concerned with produce ICP/MS data, so this method
        is hard coded to that type of data. If we can later get Agilent data off of the older Agilent instruments
        in .csv format, this method will need to be broadened with extra conditions. """
        if old_icp:
            analytes_list = self.csv_file_in_list_of_list_format[0][6:]
            for item in self.csv_file_in_list_of_list_format[1:]:
                sample_date_and_time = item[self.required_fields_index_dictionary_old_icp['Date Time']].split(" ")
                sample_date = sample_date_and_time[0]
                sample_time = sample_date_and_time[1] + " " + sample_date_and_time[2]
                if len(self.main_file_name) > 0:
                    pass
                else:
                    self.main_file_name = "Harry_icp_" + sample_date[0]
                analyte_counter = 6
                for analyte in analytes_list:
                    condensed_line = ['Agilent Instruments: ICP',
                                      'no data file provided',
                                      sample_date,
                                      sample_time,
                                      item[self.required_fields_index_dictionary_old_icp['Solution Label']],
                                      item[self.required_fields_index_dictionary_old_icp['Type']],
                                      analyte,
                                      item[analyte_counter],
                                      " "]
                    self.csv_file_in_list_of_list_format_condensed.append(condensed_line)
                    analyte_counter += 1
        else:
            for item in self.csv_file_in_list_of_list_format[1:]:
                if item[self.required_fields_index_dictionary['Tune Step']] == '1':
                    tune_step = 'no gas'
                elif item[self.required_fields_index_dictionary['Tune Step']] == '2':
                    tune_step = 'H'
                elif item[self.required_fields_index_dictionary['Tune Step']] == '3':
                    tune_step = 'He'
                else:
                    tune_step = 'not found'
                sample_date = item[self.required_fields_index_dictionary['Date and Time Acquired']][0:10]
                sample_time = item[self.required_fields_index_dictionary['Date and Time Acquired']][11:]
                analyte_name = str(item[self.required_fields_index_dictionary['Analyte']] + ' ' +
                                   item[self.required_fields_index_dictionary['Mass']] + ' [' +
                                   tune_step + ']'
                                   )
                condensed_line = ['Agilent Instruments: ICP',
                                  item[self.required_fields_index_dictionary['Data File Name']],
                                  sample_date,
                                  sample_time,
                                  item[self.required_fields_index_dictionary['Sample Name']],
                                  item[self.required_fields_index_dictionary['Sample Type']],
                                  analyte_name,
                                  item[self.required_fields_index_dictionary['Concentration']],
                                  " "]
                self.csv_file_in_list_of_list_format_condensed.append(condensed_line)
            # could do this better, getting the main file name from the first item (individual data file name) minus
            # the last 3 numbers
            self.main_file_name = self.csv_file_in_list_of_list_format[1][5][:-2]

    def generate_excel_files(self):
        """generates excel file versions of the unified excel format, using xlsxwriter.

        allows us to add formatting to the file produced, which we can't do with the .csv version."""

        target = r'T:\ANALYST WORK FILES\Peter\CrystalMB\UnifiedExcelFiles\ '
        batch_name = str(self.main_file_name)
        filename = target[:-1] + batch_name + '.xlsx'
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        # set column width
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:I', 20)
        # header format for the first two columns
        header_cell_format_1 = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#2321a7'})
        header_cell_format_2 = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#3330db'})
        header_cell_format_3 = workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': '#29abdc'})
        odd_sample_format = workbook.add_format({'bg_color': "#c7d6db"})
        even_sample_format = workbook.add_format({'bg_color': "#e9edef"})
        row = 0
        for item in self.csv_file_in_list_of_list_format_condensed:
            if row == 0:
                worksheet.write(row, 0, item[0], header_cell_format_1)
                worksheet.write(row, 1, item[1], header_cell_format_2)
                worksheet.write(row, 2, item[2], header_cell_format_2)
                worksheet.write(row, 3, item[3], header_cell_format_2)
                worksheet.write(row, 4, item[4], header_cell_format_2)
                worksheet.write(row, 5, item[5], header_cell_format_2)
                worksheet.write(row, 6, item[6], header_cell_format_3)
                worksheet.write(row, 7, item[7], header_cell_format_3)
                worksheet.write(row, 8, item[8], header_cell_format_3)
            else:
                if row % 2 == 0:
                    worksheet.write(row, 0, item[0], even_sample_format)
                    worksheet.write(row, 1, item[1], even_sample_format)
                    worksheet.write(row, 2, item[2], even_sample_format)
                    worksheet.write(row, 3, item[3], even_sample_format)
                    worksheet.write(row, 4, item[4], even_sample_format)
                    worksheet.write(row, 5, item[5], even_sample_format)
                    worksheet.write(row, 6, item[6], even_sample_format)
                    worksheet.write(row, 7, item[7], even_sample_format)
                    worksheet.write(row, 8, item[8], even_sample_format)
                else:
                    worksheet.write(row, 0, item[0], odd_sample_format)
                    worksheet.write(row, 1, item[1], odd_sample_format)
                    worksheet.write(row, 2, item[2], odd_sample_format)
                    worksheet.write(row, 3, item[3], odd_sample_format)
                    worksheet.write(row, 4, item[4], odd_sample_format)
                    worksheet.write(row, 5, item[5], odd_sample_format)
                    worksheet.write(row, 6, item[6], odd_sample_format)
                    worksheet.write(row, 7, item[7], odd_sample_format)
                    worksheet.write(row, 8, item[8], odd_sample_format)

            row += 1
        workbook.close()

    def generate_csv_files(self):
        """generates csv file versions of the unified excel format, using python's built in csv library.

        files generated have no formatting. Can use generate_excel_files, which comes with visual formatting."""

        target = r'T:\ANALYST WORK FILES\Peter\CrystalMB\UnifiedExcelFiles\ '
        try:
            batch_name = str(self.main_file_name)
            filename = target[:-1] + batch_name + '.csv'
            filename = filename.replace('/', '-')
            with self.safe_open_w(filename) as f:
                csv_writer = csv.writer(f)
                for item in self.csv_file_in_list_of_list_format_condensed:
                    csv_writer.writerow(item)
        except OSError:
            pass

    def mkdir_p(self, path):
        """tries to make the directory."""

        try:
            os.makedirs(path)
        except OSError as exc:  # Python >2.5
            if exc.errno == errno.EEXIST and os.path.isdir(path):
                pass
            else:
                raise

    def safe_open_w(self, path):
        """ Open "path" for writing, creating any parent directories as needed. """

        self.mkdir_p(os.path.dirname(path))
        return open(path, 'w', newline='')

