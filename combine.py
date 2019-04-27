from module import Converter, Data, File

def first_step():
    data = Data()
    local_data = data.get_local()
    remote_coverted_data = Converter(data.get_remote()).convert_to_local_style()
    namelist = []
    for i in local_data:
        namelist.append(i[0])
    different_list = []
    for i in remote_coverted_data:
        if i[0] in namelist:
            pass
        else:
            different_list.append(i)
    File(different_list).write_to_local()

def second_step():
    data = Data()
    local_converted_data = Converter(data.get_local()).convert_to_remote_style()
    remote_data = data.get_remote()
    namelist = []
    for i in remote_data:
        namelist.append(i[0])
    different_list = []
    for i in local_converted_data:
        if i[0] in namelist:
            pass
        else:
            different_list.append(i)
    File(different_list).write_to_remote()

def third_step():
    data = Data()
    local_data = data.get_local()
    cancel_list = []
    for i in local_data:
        if i[-4] == "退餐":
            cancel_list.append(i)
        else:
            pass
    File(cancel_list).write_cancel()


if __name__ == '__main__':
    first_step()
    second_step()
    third_step()
