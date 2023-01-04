######################## 上海、工作日（工作日日期需要自己输入）、午餐， 输出为整月全部餐补#########################################
import pandas as pd

month = "2022.08"
day_input = [31]

grant_month = []
count_month = []
for k in range(0, len(day_input)):
    day_1 = str(day_input[k])  # 【重要！】要确保day_1是个字符串，否则影响Step 2中的if 判断

    # date_1是“2022.08.1”的格式
    date_1 = '.'.join((month, day_1))
    # date_01是“2022.08.01”的格式
    if day_input[k] <= 9:
        date_01 = '.'.join((month, "0" + day_1))
    else:
        date_01 = '.'.join((month, day_1))

    # 打开"移动和5G打卡记录表格",记作table1
    table1 = pd.read_excel('81.xlsx')
    # 打开"公司内打卡记录表格",记作table2
    table2 = pd.read_excel('82.xlsx')

    #####################################  Step 1：判断该员工是否需要发放餐补  ##################################
    ######### 根据 移动和5G打卡记录表格， 知该工号员工当日是否正常出勤 #########

    # 筛选出某一天的所有数据
    table1 = table1[(table1['日期'] == date_01)]
    # 获取行数
    row1 = table1.shape[0]

    # 将“工号”列全变成str,方面后面排序
    for index in range(0, row1):
        table1.iloc[index, 0] = str(table1.iloc[index, 0])
    # 按“工号”、“打卡时间”列排序
    table1.sort_values(by=['工号', '打卡时间'], inplace=True, ascending=True)
    # print(table)

    # 找出不重复的所有工号，放入列表中
    unduplicate_ID = set(table1["工号"].tolist())
    unduplicate_ID = list(unduplicate_ID)

    # 将 当日有5G园区正常出勤记录的 员工工号 放入 attendance 中
    date_time = []
    attendance = []
    for i in range(0, len(unduplicate_ID)):
        for j in range(0, row1):
            if table1.iloc[j, 0] == unduplicate_ID[i]:
                date_time.append(table1.iloc[j, 3])
        for z in range(0, len(date_time)):
            if date_time[z].hour < 14:  # “当日有5G园区出勤记录” 定义为 “当天14：00前有打卡记录”
                attendance.append(unduplicate_ID[i])
                break
        date_time = []
    #print(attendance)

    ######### 根据 公司内打卡记录表格， 知该工号员工当日是否在榕桥路及上海以外食堂就餐 #########
    # 获取行数
    row2 = table2.shape[0]

    # 将“工号”列全变成str,方面后面查找匹配
    for index in range(0, row2):
        table2.iloc[index, 1] = str(table2.iloc[index, 1])

    # 同时满足“工号在attendance中” 且 “日期正确” 且 "就餐时间是8~14点之间" 且 （“在SH_CT_8中” 或 “在SH_CT_201中” 或...)时，说明员工当日在公司内用午餐（那就无餐补），则将工号放入nogrant里
    nogrant = []
    for per in range(0, len(attendance)):
        for num in range(0, row2):
            if table2.iloc[num, 1] == attendance[per] and str(table2.iloc[num, 4].day) == day_1 and 8 <= table2.iloc[num, 4].hour <= 14 and (
                    str(table2.iloc[num, 7]) == "SH_CT_8" or str(table2.iloc[num, 7]) == "SH_CT_82" or str(
                    table2.iloc[num, 7]) == "SH_CT_201" or str(table2.iloc[num, 7]) == "SH_CT_202" or str(
                    table2.iloc[num, 7]) == "SH_CT_204" or str(table2.iloc[num, 7]) == "SH_CT_205" or str(
                    table2.iloc[num, 7]) == "SH_CT_206" or str(table2.iloc[num, 7]) == "SH_CT_208" or str(
                    table2.iloc[num, 7]) == "SH_CT_210" or str(table2.iloc[num, 7]) == "WX_CTA_17" or str(
                    table2.iloc[num, 7]) == "WX_WXP2_CT_53" or str(table2.iloc[num, 7]) == "WX_CT_147" or str(
                    table2.iloc[num, 7]) == "WX_CT_150" or str(table2.iloc[num, 7]) == "WX_CTB_59" or str(
                    table2.iloc[num, 7]) == "TC_CT 301" or str(table2.iloc[num, 7]) == "TC_CT 302" or str(
                    table2.iloc[num, 7]) == "TC_CT 300" or str(table2.iloc[num, 7]) == "XA_CTNew_68" or str(
                    table2.iloc[num, 7]) == "XA_CT_62" or str(table2.iloc[num, 7]) == "CQ_CT_55" or str(
                    table2.iloc[num, 7]) == "CQ_CT_86" or str(table2.iloc[num, 7]) == "CQ_CT_132" or str(
                    table2.iloc[num, 7]) == "WH_CT_135" or str(table2.iloc[num, 7]) == "WH_CT_168" or str(
                    table2.iloc[num, 7]) == "SZ_CT_162" or str(table2.iloc[num, 7]) == "SZ_CT_176" or str(
                    table2.iloc[num, 7]) == "LZP_CT_143" or str(table2.iloc[num, 7]) == "LzP_CT_97"):
                nogrant.append(attendance[per])
                #break
    #print(nogrant)

    # grant = attendance - nogrant
    grant = list(set(attendance) - set(nogrant))
    #print(grant)

    # 将每一天的grant人员工号 放入 grant_month中，最终积累出当月每天grant的人
    grant_month.append(grant)
    #print(grant_month)

# 将当月每天grant人员变成一个列表（有重复）
grant_month = sum(grant_month, [])

# ID_count_dict为员工工号和对应的餐补次数（无重复）
ID_count_dict = {}
for key in grant_month:
    ID_count_dict[key] = ID_count_dict.get(key,0)+1

ID = list(ID_count_dict.keys())
count = list(ID_count_dict.values())


#####################################  Step 2：根据不同标准对grant中员工发放餐补  ##################################
for n in range(0, len(ID)):
    if ID[n][0] == "S" or ID[n][0] == "i":
        count[n] = 28 * count[n]     # 实习生&非生产外包： 28元/餐
    else:
        count[n] = 23 * count[n]     # 员工：23元/餐
fee = count

# 将两个list变成字典可以一一对应
ID_fee = dict(zip(ID, fee))
print(ID_fee)

# 将字典导出为excel
result_excel = pd.DataFrame()
result_excel["ID"] = ID
result_excel["fee"] = fee
result_excel.to_excel(r'SH(ID-Fee).xlsx')

# 程序完成，保存好了excel文件
print("Finshed")
