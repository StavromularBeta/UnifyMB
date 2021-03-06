import csv
import os.path
import errno
import xlsxwriter


class WatersUnify:
    """Converts Waters csv data to the Unified Excel Format.

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
    data_source : string
        the source of the data - either will be UPLC/MS/MS, OR UPLC/UV. Uses the method name to figure it out.
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
        self.required_fields_index_dictionary = {'name16': 0,
                                                 'createdate': 0,
                                                 'createtime': 0,
                                                 'sampleid': 0,
                                                 'type': 0,
                                                 'name20': 0,
                                                 'analconc': 0,
                                                 'percrecovery': 0
                                                 }
        self.main_file_name = ""
        self.data_source = ""

    def waters_unify_controller(self):
        """The main controller function for WatersUnify. """

        self.find_indexes_of_required_fields()
        self.create_condensed_csv_file_with_only_relevant_fields()
        self.generate_excel_files()
        # have the option of generating excel files with xlsxwriter, can format somewhat
        # or make un-formatted csv files.
        # self.generate_csv_files()

    def find_indexes_of_required_fields(self):
        """finds the indexes of the fields we need to create our unified excel format.

        This could be hard-coded in because the fields are always in the same order regardless of instrument,
        but it's nice to have the option, so I'm leaving it in. """

        required_field_index_counter = 0
        for item in self.csv_file_in_list_of_list_format[0]:
            if item in self.required_fields_index_dictionary.keys():
                # then it is a required field
                # add the index we are at to the dictionary
                self.required_fields_index_dictionary[item] = required_field_index_counter
            required_field_index_counter += 1

    def create_condensed_csv_file_with_only_relevant_fields(self):
        """takes only the required fields from the full csv file, and creates a condensed list of lists.

        Waters XML files have an abundance of data and produce all the values we need separately, so we just have
        to re-organize them. With Agilent instruments, a little more massaging is needed. We just need functionality
        to filter out blank lines, which isn't necessary with the Agilent .csv output files. """
        for item in self.csv_file_in_list_of_list_format[1:]:
            if item[self.required_fields_index_dictionary['name16']] == '':
                # so we can ignore the blank lines potentially at the end of the file
                pass
            else:
                # creating a csv line
                condensed_line = [str("Waters Instruments: " + item[12]),
                                  item[self.required_fields_index_dictionary['name16']],
                                  item[self.required_fields_index_dictionary['createdate']],
                                  item[self.required_fields_index_dictionary['createtime']],
                                  item[self.required_fields_index_dictionary['sampleid']],
                                  item[self.required_fields_index_dictionary['type']],
                                  item[self.required_fields_index_dictionary['name20']],
                                  item[self.required_fields_index_dictionary['analconc']],
                                  item[self.required_fields_index_dictionary['percrecovery']]]
                self.csv_file_in_list_of_list_format_condensed.append(condensed_line)
        # could do this better, getting the main file name from the first item (individual data file name) minus
        # the last 3 numbers
        self.main_file_name = self.csv_file_in_list_of_list_format_condensed[1][1][:-4]

    def generate_excel_files(self):
        """generates excel file versions of the unified excel format, using xlsxwriter.

        allows us to add formatting to the file produced, which we can't do with the .csv version. the formats
        denote shared fields - all rows in the file will have the same values in header_cell_format_1 cells,
        all rows in a sample will have the same values in header_cell_format_2 cells, and the cells with
        header_format_3 change each line.  """

        target = r'T:\ANALYST WORK FILES\Peter\CrystalMB\UnifiedExcelFiles\ '
        batch_name = str(self.main_file_name)
        filename = target[:-1] + batch_name + '.xlsx'
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        # set column width
        worksheet.set_column('A:A', 50)
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
                worksheet.write(row, 5, item[5], header_cell_format_3)
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
        """generates csv file versions of the unified excel format, using python's built in csv library. """

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








