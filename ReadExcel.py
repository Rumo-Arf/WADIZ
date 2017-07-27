from openpyxl import load_workbook


class ReadExcel:
    howmuch = list()
    numberof = list()

    path = str()

    def __init__(self, path):
        self.path = path
        workbook = load_workbook(path, data_only=True)
        worksheet = workbook.get_sheet_by_name(workbook.get_sheet_names()[0])

        # for row in worksheet.rows:
        #     for cell in row:
        #         test += str(cell.value) + " "
        #     values.append(test)
        #     test = ""

        # make list of NO.
        y_axis = 2
        cell = "A" + str(y_axis)

        while worksheet[cell].value is not None and worksheet[cell].value is not "":
            self.numberof.append(worksheet[cell].value)
            y_axis += 1
            cell = "A" + str(y_axis)

        self.population = worksheet[cell].value

        y_axis = 2
        cell = "I" + str(y_axis)

        while worksheet[cell].value is not None and worksheet[cell].value is not "":
            self.howmuch.append(worksheet[cell].value)
            y_axis += 1
            cell = "I" + str(y_axis)

        if len(self.numberof) != len(self.howmuch):
            print("error!", len(self.numberof), len(self.howmuch))
            return

    def get_population(self):
        return len(self.numberof)

    def get_excel_path(self):
        return self.path

    def get_howmuch(self):
        return self.howmuch

    def get_numberof(self):
        return self.numberof
