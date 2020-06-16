import os
import xlrd
import numpy
import xlwt


def pallet_types(listOfFiles):

    panel_cnt = 0
    ply_cnt = 0
    pan_ply_cnt = 0
    three_top = 0
    three_bottom = 0
    three_stringers = 0

    # Print the files
    for elem in listOfFiles:
        top_no = 0
        bottom_no = 0
        stringer_no = 0
        top_ply = False
        top_pan = False
        bottom_ply = False
        bottom_pan = False

        #verify file type and content are desired
        if ".xls" in elem or ".XLS" in elem:
            temp = xlrd.open_workbook(elem)
        else:
            continue

        sheet = temp.sheet_by_index(0)
        title = str(sheet.cell(1, 0).value)
        if "DATA" not in str.strip(title):
            continue

        # top deck board
        for row in range(5, 8, 1):

            if sheet.cell(row, 2).value != xlrd.empty_cell.value:
                top_no = 1 + top_no
                if sheet.cell(row, 2).value == "PLY":
                    top_ply = True
                if sheet.cell(row, 14).value == "Y":
                    top_pan = True

        # bottom deck board
        for row in range(9, 12, 1):

            if sheet.cell(row, 2).value != xlrd.empty_cell.value:
                bottom_no = 1 + bottom_no
                if sheet.cell(row, 2).value == "PLY":
                    bottom_ply = True
                if sheet.cell(row, 14).value == "Y":
                    bottom_pan = True

        # stringers
        for row in range(13, 16, 1):

            if sheet.cell(row, 2).value != xlrd.empty_cell.value:
                stringer_no = 1 + stringer_no

        if top_no >= 3:
            three_top = three_top + 1
        if bottom_no >= 3:
            three_bottom = three_bottom + 1
        if stringer_no >= 3:
            three_stringers = three_stringers + 1
        if top_ply == True or bottom_ply == True:
            ply_cnt = ply_cnt + 1
        if top_pan == True or bottom_pan == True:
            panel_cnt = panel_cnt + 1
        if top_ply == True and top_pan or bottom_ply == True and bottom_pan == True:
            pan_ply_cnt = pan_ply_cnt + 1

    print("three top => ", three_top)
    print("three bottom => ", three_bottom)
    print("three stringers => ", three_stringers)
    print("panels => ", panel_cnt)
    print("ply => ", ply_cnt)
    print("both ply and pan => ", pan_ply_cnt)


def labor_adjustments(listOfFiles):
    for elem in listOfFiles:

        # verify file type and content are desired
        if ".xls" in elem or ".XLS" in elem:
            temp = xlrd.open_workbook(elem)
        else:
            continue

        sheet = temp.sheet_by_index(0)
        title = str(sheet.cell(1, 0).value)
        if "DATA" not in str.strip(title):
            continue

        if sheet.cell(98,12).value != xlrd.empty_cell.value and sheet.cell(98,12).ctype == xlrd.XL_CELL_NUMBER\
        and sheet.cell(99,12).ctype == xlrd.XL_CELL_NUMBER and sheet.cell(100,12).ctype == xlrd.XL_CELL_NUMBER\
        and sheet.cell(102,12).ctype == xlrd.XL_CELL_NUMBER:

            #determine labor ratios
            if sheet.cell(100,12).value != xlrd.empty_cell.value and sheet.cell(98,12).value >0:
                effic_ratio = sheet.cell(100,12).value/sheet.cell(98,12).value
            else:
                effic_ratio = 0

            if sheet.cell(99, 12).value != xlrd.empty_cell.value and sheet.cell(98,12).value > 0:
                setup_ratio = sheet.cell(99,12).value/sheet.cell(98,12).value
            else:
                setup_ratio = 0

            if setup_ratio + effic_ratio >= 0 and sheet.cell(98,12).value >0:
                adjust_ratio = (sheet.cell(100,12).value+sheet.cell(99,12).value)/sheet.cell(98,12).value
            else:
                adjust_ratio = 0

            cost_pallet = 30*sheet.cell(102,12).value
            if sheet.cell(102,9).ctype == xlrd.XL_CELL_NUMBER:
                assemb_ratio = sheet.cell(102,9).value*30/cost_pallet
            # if cost_pallet>5 and adjust_ratio>1:
            #     print(elem)

            print(str(effic_ratio) + "," + str(setup_ratio) + "," + str(adjust_ratio) + "," + str(cost_pallet) + "," +
                  str(sheet.cell(20,3).value) + "," + str(sheet.cell(55,5).value) + "," + str(sheet.cell(108,3).value) + "," + str(sheet.cell(108,4).value) + "," + str(assemb_ratio))

def board_dims(listOfFiles):
    vals = numpy.zeros([250,4])
    for elem in listOfFiles:
        in_use = False
        i=0
        # verify file type and content are desired
        if ".xls" in elem or ".XLS" in elem:
            temp = xlrd.open_workbook(elem)
        else:
            continue

        sheet = temp.sheet_by_index(0)
        title = str(sheet.cell(1, 0).value)
        if "DATA" not in str.strip(title):
            continue
        if sheet.cell(38, 3).value != xlrd.empty_cell.value and sheet.cell(38,3).ctype == xlrd.XL_CELL_NUMBER \
            and sheet.cell(39,3).value != xlrd.empty_cell.value and sheet.cell(39,3).ctype == xlrd.XL_CELL_NUMBER \
            and sheet.cell(98, 5).ctype == xlrd.XL_CELL_NUMBER:

            thickness = sheet.cell(38,3).value
            gross_thickness = sheet.cell(39,3).value
            width = sheet.cell(38,4).value
            gross_width = sheet.cell(39,4).value

            for i in range(0,len(vals)):
                if abs(vals[i][0] - thickness) < 1/16 and vals[i][2] == width:
                    in_use = True
                elif vals[i][0] == 0 and in_use == False and sheet.cell(98,5).value > 0 and gross_thickness>thickness:
                    vals[i][0] = thickness
                    vals[i][1] = gross_thickness
                    vals[i][2] = width
                    vals[i][3] = gross_width
                    print(elem)
                    in_use = True

    print(vals)

def read_data(listOfFiles):
    nvals = numpy.zeros([len(listOfFiles), 20])
    vals = nvals.astype('U')
    i=0
    for elem in listOfFiles:
        in_use = False
        # verify file type and content are desired
        if ".xls" in elem or ".XLS" in elem:
            temp = xlrd.open_workbook(elem)
        else:
            continue

        sheet = temp.sheet_by_index(0)
        title = str(sheet.cell(1, 0).value)
        if "DATA" not in str.strip(title):
            continue
        name = elem
        vals[i][0] = name.replace("C:/Users/Greg/Documents/Burgess/COST ESTIMATES\\","")                               #filename
        if sheet.cell(24,1).ctype == 3:
            wrongval= sheet.cell(24, 1).value
            temp_datemode = temp.datemode
            y,m,d,h,min,s = xlrd.xldate_as_tuple(wrongval,temp_datemode)
            vals[i][1] = str("{2}-{1}-{0}".format(y,m,d))    #date
        else:
            vals[i][1] = str(sheet.cell(24, 1).value)
        vals[i][2] = str(sheet.cell(26, 2).value)       #customer
        vals[i][3] = str(sheet.cell(28, 6).value)       #size
        vals[i][4] = str(sheet.cell(28, 2).value)       #wood type
        try: vals[i][5] = float(sheet.cell(30, 1).value)       #quantity
        except ValueError: vals[i][5]=0

        try: vals[i][6] = float(sheet.cell(70, 12).value)     #total lumber cost/plt
        except ValueError: vals[i][6]=0

        try: vals[i][7] = float(sheet.cell(72, 12).value)     #total nail cost/plt
        except ValueError: vals[i][7]=0

        try: vals[i][8] = float(sheet.cell(102, 2).value)     #cant saw hr/plt
        except ValueError: vals[i][8] = 0

        try: vals[i][9] = float(sheet.cell(102, 3).value)      #ind. saw hr/plt
        except ValueError: vals[i][9] = 0

        try: vals[i][10] = float(sheet.cell(102, 4).value)     #trim hr/plt
        except ValueError: vals[i][10] = 0

        try: vals[i][11] = float(sheet.cell(102, 5).value)   #band re-saw hr/plt
        except ValueError: vals[i][11] = 0

        try: vals[i][12] = float(sheet.cell(102, 6).value)     #chamfer hr/plt
        except ValueError: vals[i][12] = 0

        try: vals[i][13] = float(sheet.cell(102, 7).value)     #notch hr/plt
        except ValueError: vals[i][13] = 0

        try: vals[i][14] = float(sheet.cell(102, 8).value)     #panel saw hr/plt
        except ValueError: vals[i][14] = 0

        try: vals[i][15] = float(sheet.cell(102, 9).value)     #assembly hr/plt
        except ValueError: vals[i][15] = 0

        try: vals[i][16] = float(sheet.cell(102, 10).value)    #package saw hr/plt
        except ValueError: vals[i][16] = 0

        try: vals[i][17] = float(sheet.cell(102, 12).value)   #total hr/plt
        except ValueError: vals[i][17] = 0

        try: vals[i][18] = float(sheet.cell(107, 12).value)    #labor cost/plt
        except ValueError: vals[i][18] = 0

        try: vals[i][19] = float(sheet.cell(112, 12).value)    #total pallet cost
        except ValueError: vals[i][19] = 0

        i=i+1
    return vals

def write_data(vals):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('MAIN')

    row = 0
    col = 0
    style = xlwt.XFStyle()
    style.num_format_str = "0.000000"
    for row in range(len(vals)):
        for col in range(len(vals[0])):
            sheet.write(row+1,col,vals[row][col],style)


    workbook.save('OLD ESTIMATES.xls')



def main():
    dirName = 'C:/Users/Greg/Documents/Burgess/COST ESTIMATES';


    # Get the list of all files in directory tree at given path
    listOfFiles = list()
    for (dirpath, dirnames, filenames) in os.walk(dirName):

        listOfFiles += [os.path.join(dirpath, file) for file in filenames]

    #pallet_types(listOfFiles);
    #labor_adjustments(listOfFiles);
    #board_dims(listOfFiles);
    vals = read_data(listOfFiles);
    write_data(vals);

if __name__ == '__main__':
    main()
