import csv
import numpy as np
import xlrd
import xlwt
from xlutils.copy import copy


def get_mean_var_std(arr):
    # 求均值
    arr_mean = np.mean(arr)
    # 求方差
    arr_var = np.var(arr)
    # 求标准差
    arr_std = np.std(arr, ddof=1)
    print("平均值为：%f" % arr_mean)
    print("方差为：%f" % arr_var)
    print("标准差为:%f" % arr_std)

    return arr_mean, arr_var, arr_std


# 新建表格
def excel_int(path, sheet_name):
    workbook = xlwt.Workbook()  # 新建一个工作簿
    workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    workbook.save(path)  # 保存工作簿
    print("新建表格成功，表格名称为：", path)


# 写入表头
def excel_write_title(path, titels):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for j in range(0, len(titels)):
        new_worksheet.write(0, j, str(titels[j]))  # 表格中写入数据（对应的行）
    new_workbook.save(path)  # 保存工作簿


# 向表格按列写入一维数组（列表）
def excel_write_array(path, value, column):
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, len(value)):
        # 向表格中写入数据（对应的列），初始位置加1（因为有表头）
        new_worksheet.write(i + 1, column, float(value[i]))
    new_workbook.save(path)  # 保存工作簿


def get_data(csvPath, datarow):
    angle_ = []
    with open(csvPath, 'r') as f:
        reader = csv.reader(f)
        for i in reader:
            angle_.append(float(i[datarow]))
    angle_ = np.array(angle_)
    return angle_


if __name__ == "__main__":
    # 读取.cvs表格中角度数值angle
    csvname = "台州实验-插秧机"
    csv_path = 'data/' + csvname + ".csv"
    data_row = -2  # angle数据都是在倒数第二列，所以这里为-2
    angle = get_data(csv_path, data_row)

    # 设置初始经验值
    minup = 2  # TODO  设为初始的变化阈值，可设小一点
    batch = 4  # TODO  表示第几个变化边界点，该点之前均为准备状态

    # 原始angle数据 做一次差分
    diff = np.diff(angle)
    # 将差分中  >T1:1  <T1:0
    diff_abs_over_minup = np.where(abs(diff) > minup, 1, 0)
    # 做二次差分，这里无论向上还是向下都用一对：1，0，0，0，-1，来确定边界点
    diff_diff = np.diff(diff_abs_over_minup)
    # 利用【-1】来确定边界点，且将前【batch】个变化设置成初始已知值
    num = int(np.argwhere(diff_diff == -1)[batch])

    # 将差分中  >T1:diff原始值  <T1:0
    diff2 = np.where(abs(diff) > minup, diff, 0)
    # 提取出初始已知值的数据段，计算非零变化值的绝对值平均值，将其设置成后面变化阈值T
    diff_known = diff2[0:num]
    # 删除diff_known数组中值为0的数值，保留非0数值单独为一个数组
    diff_know_nozero = diff_known.ravel()[np.flatnonzero(diff_known)]
    # 求一组数据的均值，方差，标准差
    mean, var, std = get_mean_var_std(abs(diff_know_nozero))
    # round((data,len)) 保留data数据小数点长度为len
    T_upSize = round((mean - std), 3)  # TODO  可以考虑后面：将新的变化大的数值填充进去，扩充数据集
    print("高度变化阈值 T_upSize：{}".format(T_upSize))

    state = [0] * num  # 初始化 state
    h = 3  # state 数值高度
    # 设置workingFlag初始状态
    if diff2[num] > 0:
        workingFlag = h
    else:
        workingFlag = -h

    # 判断num之后，diff[i]状态，实际上这里真实情况应是输入angle[i]  # TODO
    for i in range(num, len(diff)):
        if abs(diff[i]) > T_upSize:
            # print("State changes!")
            if diff[i] > 0:
                workingFlag = h  # 非工作状态
            else:
                workingFlag = -h  # 工作状态
        else:
            None
            # print(workingFlag)
        state.append(workingFlag)
    state = np.array(state)

    # 保存到excel
    excel_name = "result/" + csvname + "_batch" + str(batch) + "_minup" + str(minup) + "_T_upSize" + str(
        T_upSize) + ".xls"
    sheet_name = "patameters"
    title = ["angle", "state", "diff", "diff2", "diff_abs_over_minup", "diff_diff", "diff_know_nozero"]
    # 新建表格
    excel_int(excel_name, sheet_name)
    # 写入表头
    excel_write_title(excel_name, title)
    # 写入数据
    save_data = [angle, state, diff, diff2, diff_abs_over_minup, diff_diff, diff_know_nozero]
    for i in range(len(save_data)):
        excel_write_array(excel_name, save_data[i], i)
