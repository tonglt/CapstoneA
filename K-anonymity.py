import xlrd





def main():
    count=0
    k=100
    for row in range(0, mysheet.nrows):
        temp=""
        i = 0
        #index =2
        str1 =""
        for col in range(0, mysheet.ncols):
            if mysheet.cell(row, col).value != None:
                temp+=str(mysheet.cell(row, col).value)+"\t"
                if str(mysheet.cell(row,col).value) == "T":
                    index=2
                    i = row+index
                    for i in range(i, mysheet.nrows):
                        if(str1 != mysheet.cell(i,col).value):
                            str1 = mysheet.cell(i,col).value
                            j=row+2
                            print(str1)
                            for j in range(j, mysheet.nrows):
                                print(str(mysheet.cell(j, col).value))
                                if (str1 == str(mysheet.cell(j, col).value)):
                                    count += 1
                                    print(count)
                                    print(k)
                            index = index + 1
                            if (k > count):
                                k = count
                            count=0
                #if(k>count):
                    #k=count
                #count =0
        print(temp)
    print("##################")
    print("k=%d" %(k))


if __name__ == '__main__':
    xlsfile = xlrd.open_workbook("/Users/tony/Downloads/data.xlsx")
    count = 0
    str1=""
    try:
        mysheet = xlsfile.sheet_by_name("sheet1")
    except:
        print("no sheet in %s named PC")

    print("%d rows, %d cols" % (mysheet.nrows, mysheet.ncols))
    main()
