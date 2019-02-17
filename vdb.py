import os
import openpyxl


def add_coord(wb):
    sheet = wb['Add Coordinator']
    fname = sheet['D2'].value
    lname = sheet['D4'].value
    mobile_phone = sheet['D6'].value
    office_phone = sheet['D8'].value
    email = sheet['D10'].value
    office_add = sheet['D12'].value
    notes = sheet['D14'].value
    coord_import = [fname, lname, mobile_phone, office_phone, email,
                    office_add, notes]
    # print("coord_name: ", sheet['coord_name'].value)
    return coord_import


def save_coord(wb, coord_import):
    coord = wb['Coordinators']
    nrows = coord.max_row
    newrow = nrows + 1

    # save id
    cell = 'A' + str(newrow)
    coord[cell] = newrow - 1

    # save data
    for i in range(7):
        cell = chr(ord('B') + i) + str(newrow)
        coord[cell] = coord_import[i]


def main():
    filename = 'volunteering_database.xlsx'
    file = os.path.join('D:\\volunteerDB', filename)
    wb = openpyxl.load_workbook(file)
    coord_import = add_coord(wb)
    save_coord(wb, coord_import)
    wb.save(file)


if __name__ == '__main__':
    main()
