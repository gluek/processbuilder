#!python

# txtToXlsWriter
# 2014-09 jan gruis
# python 3.4.1


#Erstellt einen neunen Laufzettel aus einzelenen Textdateien
#Verwendete Module: XlsxWriter: https://pypi.python.org/pypi/XlsxWriter

import xlsxwriter
import sys

class TXTtoXLSConverter:
    """setting up control symbols"""
    __caption_char__ = '='   # indicates "headlines" for steps,   only as first character
    __new_file_char__ = '>'   # indicates a new file to open,      only as first character
    __empty_line_char__ = '_'   # indicates an empty line,           only as first character
    __ignore_line_char__ = '-'   # line will be ignored,              only as first character
    __cell_break_char__ = '|'   # indicates a cell-"break",          only useable w/o headlines and new file
    __format_char__ = '§'   # indicates a change in format,      only usable after headlines!

    __file_extension__ = 'txt'     # w/o dot

    __row__ = -2    # script increases the row before it writes something! so to be 0 the first line to write in we need -1
    __col__ = 0


    def convertTXTtoXLS(process_filename, excel_filename):

        """Goal is to write into an Excel Spreedsheet, so we start with this
        Create an new Excel file and add a worksheet."""
        workbook = xlsxwriter.Workbook(excel_filename)
        worksheet = workbook.add_worksheet()

        TXTtoXLSConverter.__build_formats__(workbook)

        # Open the "process" file
        end = TXTtoXLSConverter.__export_file__(worksheet, process_filename, TXTtoXLSConverter.__row__)
            # this function is recursive, will handle the file processing and write the excel file

        # Set a few parameters for a nice excel file
        # Freeze the first column
        worksheet.freeze_panes(0, 1)
        # worksheet.set_column(first_col, last_col, width, cell_format, options)
        worksheet.set_column(0, 0, 20)  # first col
        worksheet.set_column(1, 1, 1)
        worksheet.set_column(2, 2, 25)  # third col, wenn writing more cols, expand this
        # worksheet.set_row(row, height, cell_format, options)
        for x in range(end):
            worksheet.set_row(x, 15, None)

        # We opened a workbook and are nice enough to close it ;)
        workbook.close()
        return

    def __export_file__(worksheet, filename, row):
        """works through the file/s and writes into xlsx
        filename/filepath w/o extension"""

        file = open('.'.join((filename.split(".")[0], TXTtoXLSConverter.__file_extension__)), mode='r', encoding='UTF-8')
        for line in file:

            # caption
            if line[0] == TXTtoXLSConverter.__caption_char__:
                row = TXTtoXLSConverter.__write_caption__(worksheet, line[1:].strip(), row)

            # new file
            elif line[0] == TXTtoXLSConverter.__new_file_char__:
                row = TXTtoXLSConverter.__export_file__(worksheet, line[1:].strip(), row)

            # empty line
            elif line[0] == TXTtoXLSConverter.__empty_line_char__:
                row += 1  # skip one row

            # ignored line
            elif line[0] == TXTtoXLSConverter.__ignore_line_char__:
                pass  # do nothing

            # normal
            else:
                row = TXTtoXLSConverter.__write_line__(worksheet, line.strip(), row)

        file.close()
        return row

    def __write_caption__(worksheet, text, row, col=0):
        """writing caption cells and handling the formats"""
        global current_format
        if TXTtoXLSConverter.__format_char__ in text:
            text, format_name = text.split(TXTtoXLSConverter.__format_char__)
            text = text.strip()
            format_name = format_name.strip()
            if format_name in formats_dict.keys():
                current_format = formats_dict[format_name]
            else:
                current_format = formats_dict['default']

        row += 2  # always write in the next row + one line space for captions
        worksheet.write(row, col,   text.strip(), current_format)
        worksheet.write(row, col+2, text.strip(), current_format)
        return row

    def __write_line__(worksheet, text, row, col=0):
        """writing default cells"""

        if text.strip() == '':
            return row  # do not write empty inputlines, except wenn marked with empty_line_char (see above)
        row += 1    # always write in the next row
        cells = text.split(TXTtoXLSConverter.__cell_break_char__)
        for cell in cells:
            tmpCell = cell.strip().replace("<br>","\n")  # creates linebreaks at <br>
            worksheet.write(row, col, tmpCell, current_format)
            col += 2  # one free column between the first col and the information
            #print(cell.strip(), end="")
        #print(cells)
        return row

    def __build_formats__(workbook):
        """this has to be called after creating the workbook and worksheet"""

        # http://xlsxwriter.readthedocs.org/en/latest/working_with_formats.html
        #{'font_name':'Arial'}
        #{'font_size':10}
        #{'text_wrap':True} # Turn text wrapping on for text in a cell
        #{'pattern':1} # The most common pattern is 1 which is a solid fill of the background color.
        #{'bg_color':'green'} #{'bg_color':'#RRGGBB'} # The set_bg_color() method can be used to set the background colour of a pattern. Patterns are defined via the set_pattern() method. If a pattern hasn’t been defined then a solid fill pattern is used as the default.
        #{'fg_color':'green'} #{'fg_color':'#RRGGBB'} # The cell font color.

        default = { 'font_name': 'Arial',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'bottom': 3,
                    'right': 1}

        global format_default
        format_default = workbook.add_format(default)

        fcaption = default.copy()
        fcaption.update({'font_size': 16})
        global format_caption
        format_caption = workbook.add_format(fcaption)

        fheader = default.copy()
        fheader.update({'bg_color': '#FF0000','bold': True,'font_size': 12})
        global format_header
        format_header = workbook.add_format(fheader)

        flitho = default.copy()
        flitho.update({'bg_color': '#FFFF99'})
        global format_litho
        format_litho = workbook.add_format(flitho)

        fetch = default.copy()
        fetch.update({'bg_color': '#99CCFF'})
        global format_etch
        format_etch = workbook.add_format(fetch)

        fmetal = default.copy()
        fmetal.update({'bg_color': '#CC99FF'})
        global format_metal
        format_metal = workbook.add_format(fmetal)

        fanneal = default.copy()
        fanneal.update({'bg_color': '#CCFFFF'})
        global format_anneal
        format_anneal = workbook.add_format(fanneal)

        fpassivation = default.copy()
        fpassivation.update({'bg_color': '#00FF00'})
        global format_passivation
        format_passivation = workbook.add_format(fpassivation)

        fmeasurement = default.copy()
        fmeasurement.update({'bg_color': '#FFFF00'})
        global format_measurement
        format_measurement = workbook.add_format(fmeasurement)

        fimplant = default.copy()
        fimplant.update({'bg_color': '#CCCCCC'})
        global format_implant
        format_implant = workbook.add_format(fimplant)

        global formats_dict
        formats_dict = {'default': format_default,
                        'caption': format_caption,
                        'header': format_header,
                        'litho': format_litho,
                        'etch': format_etch,
                        'metal': format_metal,
                        'anneal': format_anneal,
                        'passivation': format_passivation,
                        'measurement': format_measurement,
                        'implant': format_implant}

        # set the default format to the first current format
        global current_format
        current_format = format_default

