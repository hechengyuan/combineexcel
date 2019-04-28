import xlrd, xlwt, time, datetime
from os.path import join
from xlutils.copy import copy

class Converter:

    def __init__(self,list):
        self.list = list

    def convert_to_remote_style(self):
        local_list = self.list
        converted_list = []
        for i in local_list:
            if i[-1] == "2餐":
                lunch_data = [i[0],i[1],i[2],i[3],i[4],"午餐"]
                dinner_data = [i[0],i[1],i[2],i[3],i[4],"晚餐"]
                converted_list.append(lunch_data)
                converted_list.append(dinner_data)
            elif i[-1] in ["午餐","晚餐"]:
                converted_list.append(i)
            else:
                pass
        return converted_list

    def convert_to_local_style(self):
        remote_list = self.list
        converted_list = []
        namelist = []
        for i in remote_list:
            if i[0] in namelist:
                converted_list[namelist.index(i[0])]=[i[0],i[1],i[2],i[3],i[4],"2餐"]
            else:
                namelist.append(i[0])
                converted_list.append(i)
        return converted_list



class Data:

    def __init__(self):
        self.today = datetime.date.today()
        self.tomorrow = self.today + datetime.timedelta(days=1)
        self.xltomorrow = xlrd.xldate.xldate_from_date_tuple((self.tomorrow.year,self.tomorrow.month,self.tomorrow.day), 0)
        self.localpath = "订餐备注.xlsx"
        self.remotepath = "订餐表" + str(self.tomorrow) + ".xlsx"

    def get_local(self):
        table = xlrd.open_workbook(self.localpath).sheets()[0]
        namelist = table.col_values(0, start_rowx=0, end_rowx=None)
        departmentlist = table.col_values(1, start_rowx=0, end_rowx=None)
        bedlist = table.col_values(2, start_rowx=0, end_rowx=None)
        channellist = table.col_values(4, start_rowx=0, end_rowx=None)
        notelist = table.col_values(3, start_rowx=0, end_rowx=None)
        tomorrowlist = table.col_values(int(self.xltomorrow - 43550.0), start_rowx=0, end_rowx=None)
        all_list = zip(namelist,departmentlist,bedlist,channellist,notelist,tomorrowlist)
        order_list = []
        for i in all_list:
            if i[-1] in ['','周一','周二','周三','周四','周五','周六','周日']:
                pass
            elif i[0] == "姓名":
                pass
            else:
                order_list.append(list(i))
        return order_list

    def get_remote(self):
        table = xlrd.open_workbook(self.remotepath).sheets()[0]
        namelist = table.col_values(2, start_rowx=0, end_rowx=None)
        departmentlist = table.col_values(1, start_rowx=0, end_rowx=None)
        bed_and_weekday_list = table.col_values(3, start_rowx=0, end_rowx=None)
        bedlist = []
        for i in bed_and_weekday_list:
            bedlist.append(i.split("/")[0])
        channellist = table.col_values(12, start_rowx=0, end_rowx=None)
        notelist = table.col_values(13, start_rowx=0, end_rowx=None)
        tomorrowlist = table.col_values(11, start_rowx=0, end_rowx=None)
        all_list = zip(namelist,departmentlist,bedlist,channellist,notelist,tomorrowlist)
        order_list=[]
        for i in all_list:
            if i[0] == "姓名":
                pass
            else:
                order_list.append(list(i))
        return order_list


class File:

    def __init__(self,list):
        self.today = datetime.date.today()
        self.tomorrow = self.today + datetime.timedelta(days=1)
        self.xltomorrow = xlrd.xldate.xldate_from_date_tuple((self.tomorrow.year,self.tomorrow.month,self.tomorrow.day), 0)
        self.localpath = "订餐备注.xlsx"
        self.remotepath = "订餐表" + str(self.tomorrow) + ".xlsx"
        self.list = list

    def write_to_local(self):
        file = xlrd.open_workbook(self.localpath)
        table = file.sheets()[0]
        avaliable_file = copy(file)

        for i in self.list:
            index = self.list.index(i)
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 0, i[0])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 1, i[1])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 2, i[2])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 4, i[3])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 3, i[4])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, int(self.xltomorrow - 43550.0), i[5])

        avaliable_file.save('cookie/new_local.xls')

    def write_to_remote(self):
        file = xlrd.open_workbook(self.remotepath)
        table = file.sheets()[0]
        avaliable_file = copy(file)

        for i in self.list:
            index = self.list.index(i)
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 2, i[0])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 1, i[1])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 3, i[2])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 12, i[3])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 13, i[4])
            avaliable_file.get_sheet(0).write(table.nrows + index - 1, 11, i[5])

        avaliable_file.save("cookie/2.xls")


    def write_cancel(self):
        file = xlrd.open_workbook("cookie/2.xls")
        table = file.sheets()[0]
        avaliable_file = copy(file)
        namelist = table.col_values(2, start_rowx=0, end_rowx=None)

        for i in self.list:
            index1 = namelist.index(i[0],0)
            index2 = namelist.index(i[0],2)
            avaliable_file.get_sheet(0).write(index1, 11, i[4])
            avaliable_file.get_sheet(0).write(index2, 11, i[4])

        avaliable_file.save("cookie/3.xls")

    def write_remote_note(self):
        file = xlrd.open_workbook("cookie/3.xls")
        table = file.sheets()[0]
        avaliable_file = copy(file)

        for i in self.list:
            index = self.list.index(i,0)
            avaliable_file.get_sheet(0).write(index+1,13,i)

        avaliable_file.save("cookie/all_order.xls")
