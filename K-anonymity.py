import xlrd





def main():
    dictionary = {}
    len_count=0
    count=0
    k=100
    l=100
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
                                    dictionary[str(str1)] = str(count)
                                    print("kkk")
                                    print(dictionary)
                                    print(count)
                                    print(k)
                            index = index + 1
                            if (k > count):
                                k = count
                            count=0
                            len_count=0
                    if (len(dictionary) >= 1):
                        len_count = len(dictionary)
                        print("aaa")
                        print(len(dictionary))
                    if (l > len_count):
                        l = len_count
                    dictionary.clear()
                #if(k>count):
                    #k=count
                #count =0
        print(temp)
    print("##################")
    print("k=%d" %(k))
    print("l=%d" % (l))
    print (dictionary)
    print (len(dictionary))


if __name__ == '__main__':
    xlsfile = xlrd.open_workbook("/Users/tony/Downloads/data.xlsx")
    #len_count =0
    #count = 0
    str1=""
    try:
        mysheet = xlsfile.sheet_by_name("sheet1")
    except:
        print("no sheet in %s named PC")

    print("%d rows, %d cols" % (mysheet.nrows, mysheet.ncols))
    main()
