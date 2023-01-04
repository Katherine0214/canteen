######################## 苏州、工作日（工作日日期需要自己输入） 午餐+加班餐， 节假日（节假日日期需要自己输入）餐补  输出为整月全部餐补#########################################
import pandas as pd

month = "2022.08"
work_day_input = [22, 25]
off_day_input =[8,9,20,25]

# 打开"苏州研发中心打卡记录表格",记作table
table = pd.read_excel('3.xlsx')

######################################   【工作日餐补总合】 #########################################
grant_lunch_month = []
grant_dinner_month = []
count_month = []
for k in range(0, len(work_day_input)):
    day_1 = str(work_day_input[k])  # 【重要！】要确保day_1是个字符串，否则影响Step 2中的if 判断

    # date_1是“2022.08.1”的格式
    date_1 = '.'.join((month, day_1))

    # # 打开"苏州研发中心打卡记录表格",记作table
    # table = pd.read_excel('3.xlsx')

    #####################################  Step 1：判断该员工是否需要发放餐补  ##################################
    ######### 根据 苏州研发中心打卡记录表格， 知该工号员工当日是否正常出勤 #########

    # 筛选出某一天的有大门打卡记录的所有员工号，大门打卡设备为：SZ_GATE_170、SZ_GATE_171...
    row1 = table.shape[0]

    ID_gate = []
    for num in range(0, row1):
        if str(table.iloc[num, 4].day) == day_1 and (
                str(table.iloc[num, 7]) == "SZ_GATE_170" or str(table.iloc[num, 7]) == "SZ_GATE_171" or str(table.iloc[num, 7]) == "SZ_GATE_172" or str(table.iloc[num, 7]) == "SZ_GATE_173" or str(table.iloc[num, 7]) == "SZ_GATE_169" or str(table.iloc[num, 7]) == "SZ_GATE_134"):
            ID_gate.append(str(table.iloc[num, 1]))
    #print(ID_gate)

    # 找出不重复的所有工号，放入列表中
    unduplicate_ID_gate = set(ID_gate)
    unduplicate_ID_gate = list(unduplicate_ID_gate)
    #print(unduplicate_ID_gate)

    # 将 当日有苏州研发中心正常出勤记录(当天14：00前有打卡记录)的 员工工号 放入 attendance 中
    date_time = []
    attendance = []
    for i in range(0, len(unduplicate_ID_gate)):
        for j in range(0, row1):
            if str(table.iloc[j, 1]) == unduplicate_ID_gate[i] and str(table.iloc[j, 4].day) == day_1:
                # date_time中是unduplicate_ID某个人的所有考勤时间
                date_time.append(table.iloc[j, 4])
        for z in range(0, len(date_time)):
            if date_time[z].hour < 14:  # “当日有苏州研发中心的考勤记录” 定义为 “当天14：00前有打卡记录”
                # 如果unduplicate_ID某个人的考勤时间小于14点，就算正常出勤，将其工号放在attendance中
                attendance.append(unduplicate_ID_gate[i])
                break
        date_time = []
    #print(attendance)

    ######### 根据 苏州研发中心打卡记录表格， 知该工号员工当日是否在榕桥路及上海以外食堂就餐 #########
    # 将“工号”列全变成str,方面后面查找匹配
    for index in range(0, row1):
        table.iloc[index, 1] = str(table.iloc[index, 1])

    #####################################  （一）：判断该员工是否需要发放午餐餐补  ##################################
    # 同时满足“工号在attendance中” 且 “日期正确” 且 "就餐时间是8~14点之间" 且 （“SZ_CT_162” 或 “SZ_CT_176” 或...)时，说明员工当日在公司内用餐（那就无餐补），则将工号放入nogrant里
    nogrant_lunch = []
    for per in range(0, len(attendance)):
        for num in range(0, row1):
            if table.iloc[num, 1] == attendance[per] and str(table.iloc[num, 4].day) == day_1 and 8 <= table.iloc[num, 4].hour <= 14 and (
                    str(table.iloc[num, 7]) == "SH_CT_8" or str(table.iloc[num, 7]) == "SH_CT_82" or str(
                    table.iloc[num, 7]) == "SH_CT_201" or str(table.iloc[num, 7]) == "SH_CT_202" or str(
                    table.iloc[num, 7]) == "SH_CT_204" or str(table.iloc[num, 7]) == "SH_CT_205" or str(
                    table.iloc[num, 7]) == "SH_CT_206" or str(table.iloc[num, 7]) == "SH_CT_208" or str(
                    table.iloc[num, 7]) == "SH_CT_210" or str(table.iloc[num, 7]) == "WX_CTA_17" or str(
                    table.iloc[num, 7]) == "WX_WXP2_CT_53" or str(table.iloc[num, 7]) == "WX_CT_147" or str(
                    table.iloc[num, 7]) == "WX_CT_150" or str(table.iloc[num, 7]) == "WX_CTB_59" or str(
                    table.iloc[num, 7]) == "TC_CT 301" or str(table.iloc[num, 7]) == "TC_CT 302" or str(
                    table.iloc[num, 7]) == "TC_CT 300" or str(table.iloc[num, 7]) == "XA_CTNew_68" or str(
                    table.iloc[num, 7]) == "XA_CT_62" or str(table.iloc[num, 7]) == "CQ_CT_55" or str(
                    table.iloc[num, 7]) == "CQ_CT_86" or str(table.iloc[num, 7]) == "CQ_CT_132" or str(
                    table.iloc[num, 7]) == "WH_CT_135" or str(table.iloc[num, 7]) == "WH_CT_168" or str(
                    table.iloc[num, 7]) == "SZ_CT_162" or str(table.iloc[num, 7]) == "SZ_CT_176" or str(
                    table.iloc[num, 7]) == "LZP_CT_143" or str(table.iloc[num, 7]) == "LzP_CT_97"):
                nogrant_lunch.append(attendance[per])
                break
    #print(nogrant_lunch)

    # grant_lunch = attendance - nogrant_lunch
    grant_lunch = list(set(attendance) - set(nogrant_lunch))
    #print(grant_lunch)

    # 将每一天的grant人员工号 放入 grant_month中，最终积累出当月每天grant的人
    grant_lunch_month.append(grant_lunch)
    #print(grant_lunch_month)

    #####################################  （二）：判断该员工是否需要发放加班晚餐餐补  ##################################
    # 同时满足“工号在attendance中” 且 “日期正确” 且 "大门打卡时间晚于20点" 且 （“SZ_GATE_” 或 “SZ_GATE_” 或...)时，说明员工当日在公司加班，则将工号放入overtime里
    overtime = []
    for per in range(0, len(attendance)):
        for num in range(0, row1):
            if table.iloc[num, 1] == attendance[per] and str(table.iloc[num, 4].day) == day_1 and table.iloc[
                num, 4].hour >= 20 and (str(table.iloc[num, 7]) == "SZ_GATE_170" or str(table.iloc[num, 7]) == "SZ_GATE_171" or str(table.iloc[num, 7]) == "SZ_GATE_172" or str(table.iloc[num, 7]) == "SZ_GATE_173" or str(table.iloc[num, 7]) == "SZ_GATE_169" or str(table.iloc[num, 7]) == "SZ_GATE_134"):
                overtime.append(attendance[per])
                break
    # print(overtime)

    # 同时满足“工号在overtime中” 且 “日期正确” 且 "就餐时间晚于20点" 且 （“SZ_CT_162” 或 “SZ_CT_176” 或...)时，说明员工当日在公司内用晚餐（那就有餐补），则将工号放入grant_dinner里
    grant_dinner = []
    for per in range(0, len(overtime)):
        for num in range(0, row1):
            if table.iloc[num, 1] == overtime[per] and str(table.iloc[num, 4].day) == day_1 and table.iloc[
                num, 4].hour >= 20 and (
                    str(table.iloc[num, 7]) == "SH_CT_8" or str(table.iloc[num, 7]) == "SH_CT_82" or str(
                table.iloc[num, 7]) == "SH_CT_201" or str(table.iloc[num, 7]) == "SH_CT_202" or str(
                table.iloc[num, 7]) == "SH_CT_204" or str(table.iloc[num, 7]) == "SH_CT_205" or str(
                table.iloc[num, 7]) == "SH_CT_206" or str(table.iloc[num, 7]) == "SH_CT_208" or str(
                table.iloc[num, 7]) == "SH_CT_210" or str(table.iloc[num, 7]) == "WX_CTA_17" or str(
                table.iloc[num, 7]) == "WX_WXP2_CT_53" or str(table.iloc[num, 7]) == "WX_CT_147" or str(
                table.iloc[num, 7]) == "WX_CT_150" or str(table.iloc[num, 7]) == "WX_CTB_59" or str(
                table.iloc[num, 7]) == "TC_CT 301" or str(table.iloc[num, 7]) == "TC_CT 302" or str(
                table.iloc[num, 7]) == "TC_CT 300" or str(table.iloc[num, 7]) == "XA_CTNew_68" or str(
                table.iloc[num, 7]) == "XA_CT_62" or str(table.iloc[num, 7]) == "CQ_CT_55" or str(
                table.iloc[num, 7]) == "CQ_CT_86" or str(table.iloc[num, 7]) == "CQ_CT_132" or str(
                table.iloc[num, 7]) == "WH_CT_135" or str(table.iloc[num, 7]) == "WH_CT_168" or str(
                table.iloc[num, 7]) == "SZ_CT_162" or str(table.iloc[num, 7]) == "SZ_CT_176" or str(
                table.iloc[num, 7]) == "LZP_CT_143" or str(table.iloc[num, 7]) == "LzP_CT_97"):
                grant_dinner.append(attendance[per])
                break
    # print(grant_dinner)

    # 将每一天的grant_dinner人员工号 放入 grant_dinner_month中，最终积累出当月每天grant的人
    grant_dinner_month.append(grant_dinner)
    # print(grant_dinner_month)


# 将当月每天grant人员变成一个列表（有重复）
grant_lunch_month = sum(grant_lunch_month, [])
grant_dinner_month = sum(grant_dinner_month, [])

# lunch_ID_count_dict为员工工号和对应的午餐餐补次数（无重复）
lunch_ID_count_dict = {}
for key in grant_lunch_month:
    lunch_ID_count_dict[key] = lunch_ID_count_dict.get(key, 0) + 1

ID_lunch = list(lunch_ID_count_dict.keys())
count_lunch = list(lunch_ID_count_dict.values())


# dinner_ID_count_dict为员工工号和对应的加班晚餐餐补次数（无重复）
dinner_ID_count_dict = {}
for key in grant_dinner_month:
    dinner_ID_count_dict[key] = dinner_ID_count_dict.get(key, 0) + 1

ID_dinner = list(dinner_ID_count_dict.keys())
count_dinner = list(dinner_ID_count_dict.values())


#####################################  Step 2：根据不同标准对grant中员工发放餐补  ##################################
#####################################  （一）：对grant_lunch中员工发放午餐餐补  ##########################
for n in range(0, len(ID_lunch)):
    if ID_lunch[n][0] == "S" or ID_lunch[n][0] == "i":
        count_lunch[n] = 25 * count_lunch[n]     # 实习生&非生产外包： 25元/餐
    else:
        count_lunch[n] = 20 * count_lunch[n]     # 员工：20元/餐
fee_lunch = count_lunch

# # 将两个list变成字典可以一一对应
# lunch_ID_fee = dict(zip(ID_lunch, fee_lunch))
# print("lunch_ID_fee",lunch_ID_fee)


#####################################  （二）：对grant_dinner中员工发放加班晚餐餐补  ##########################
for n in range(0, len(ID_dinner)):
    if ID_dinner[n][0] == "S" or ID_dinner[n][0] == "i":
        count_dinner[n] = 18 * count_dinner[n]     # 实习生&非生产外包： 18元/餐
    else:
        count_dinner[n] = 18 * count_dinner[n]     # 员工：18元/餐
fee_dinner = count_dinner

# # 将两个list变成字典可以一一对应
# dinner_ID_fee = dict(zip(ID_dinner, fee_dinner))
# print("dinner_ID_fee",dinner_ID_fee)



#####################################  Step 3：将整月的 午餐 和 加班晚餐 整合在一起  ###########################
# 将午餐和加班晚餐的ID和fee分别合并到两个列表中
workDay_ID_month = ID_lunch + ID_dinner
workDay_fee_month = fee_lunch + fee_dinner
#print(workDay_ID_month, workDay_fee_month)

# 用集合表示工号列表，去重复；然后再将集合变成列表，稳定不会改变
workDay_unduplicate_ID_month = set(workDay_ID_month)
workDay_unduplicate_ID_month = list(workDay_unduplicate_ID_month)
#print("workDay_unduplicate_ID_month:",workDay_unduplicate_ID_month)


# 函数：在source列表中找出elmt的所在位置
def find_repeat(source, elmt):  # The source may be a list or string.
    elmt_index = []
    s_index = 0
    e_index = len(source)
    while (s_index < e_index):
        try:
            temp = source.index(elmt, s_index, e_index)
            elmt_index.append(temp)
            s_index = temp + 1
        except ValueError:
            break
    return elmt_index

# 遍历workDay_unduplicate_ID_month中的工号，在原始workDay_ID_month的工号列表中找到相同工号的位置，并在对应次数列表workDay_fee_month中取出次数并累加
# 然后将该工号workDay_unduplicate_ID_month 和 累加后的费用workDay_fee_month 分别放入新的两个列表中（也可得到一个字典）
workDay_accum_fee_month = []
for x in range(0, len(workDay_unduplicate_ID_month)):
    repeat_location = find_repeat(workDay_ID_month, workDay_unduplicate_ID_month[x])
    accum_fee = 0
    for y in range(0, len(repeat_location)):
        accum_fee = accum_fee + workDay_fee_month[repeat_location[y]]
    workDay_accum_fee_month.append(accum_fee)

#print("workDay_unduplicate_ID_month:", workDay_unduplicate_ID_month, "workDay_accum_fee_month:", workDay_accum_fee_month)

# 将两个list变成字典可以一一对应
workDay_ID_fee = dict(zip(workDay_unduplicate_ID_month, workDay_accum_fee_month))
print("workDay_ID_fee",workDay_ID_fee)



######################################   【节假日餐补总合】 #########################################
#####################################  Step 1：计算员工当日在公司出勤多长时间  ##################################
# 将“工号”列全变成str,方面后面排序
for index in range(0, table.shape[0]):
    table.iloc[index, 1] = str(table.iloc[index, 1])

# 将数据存成规范的二维数组表
df = pd.DataFrame(table)
# 取出第一行的字段名
first_row = df[0:0]
# 筛选出所有大门打卡的相关记录
for row in first_row:
    onlyGate = df[(df['设备名称'] == 'SZ_GATE_170') | (df['设备名称'] == 'SZ_GATE_171') | (df['设备名称'] == 'SZ_GATE_172') | (df['设备名称'] == 'SZ_GATE_173') | (df['设备名称'] == 'SZ_GATE_169') | (df['设备名称'] == 'SZ_GATE_134')]
#print(onlyGate)


OffDay_ID_month = []
OffDay_count_month = []
for i in range(0, len(off_day_input)):
    day_1 = str(off_day_input[i])  # 【重要！】要确保day是个字符串，否则影响Step 2中的if 判断
    # date_1是“2022.08.1”的格式
    date_1 = '.'.join((month, day_1))

    # 在仅有大门打卡记录的表格中 筛选出某一天所有来的员工工号
    OffDay_ID = []
    row = onlyGate.shape[0]
    for i in range(0, row):
        if str(onlyGate.iloc[i, 4].day) == day_1:
            OffDay_ID.append(onlyGate.iloc[i, 1])

    # 找出不重复的所有工号，放入列表中
    OffDay_unduplicate_ID = set(OffDay_ID)
    OffDay_unduplicate_ID = list(OffDay_unduplicate_ID)
    #print(OffDay_unduplicate_ID)

    # 针对某个工号，求得其当日总工时
    OffDay_time = []
    OffDay_count = []
    row_gate = onlyGate.shape[0]

    for j in range(0, len(OffDay_unduplicate_ID)):
        for k in range(0, row_gate):
            if onlyGate.iloc[k, 1] == OffDay_unduplicate_ID[j] and str(onlyGate.iloc[k, 4].day) == day_1:
                OffDay_time.append(onlyGate.iloc[k, 4])
        # 针对某个工号，求得其当日总工时 (以小时为单位)
        period_hour = (OffDay_time[len(OffDay_time) - 1] - OffDay_time[0]) / pd.Timedelta(1, 'H')  # 这里是只看最晚 - 最早
        # 工作时间每满4H，算1餐；当日最多2餐
        if 0 <= period_hour < 4:
            OffDay_count.append(0)
        elif 4 <= period_hour < 8:
            OffDay_count.append(1)
        else:
            OffDay_count.append(2)
        OffDay_time = []

    # 将每一天的OffDay_ID和OffDay_count append到OffDay_month的列表中
    OffDay_ID_month.append(OffDay_unduplicate_ID)
    OffDay_count_month.append(OffDay_count)


# 将两个列表降低维度
OffDay_ID_month = sum(OffDay_ID_month, [])
OffDay_count_month = sum(OffDay_count_month, [])
#print("OffDay_ID_month:",OffDay_ID_month, "OffDay_count_month:",OffDay_count_month)

# 用集合表示工号列表，去重复；然后再将集合变成列表，稳定不会改变
OffDay_unduplicate_ID_month = set(OffDay_ID_month)
OffDay_unduplicate_ID_month = list(OffDay_unduplicate_ID_month)
#print("OffDay_unduplicate_ID_month:",OffDay_unduplicate_ID_month)

# 函数：在source列表中找出elmt的所在位置
def find_repeat(source, elmt):  # The source may be a list or string.
    elmt_index = []
    s_index = 0
    e_index = len(source)
    while (s_index < e_index):
        try:
            temp = source.index(elmt, s_index, e_index)
            elmt_index.append(temp)
            s_index = temp + 1
        except ValueError:
            break
    return elmt_index

# 遍历OffDay_unduplicate_ID_month中的工号，在原始OffDay_ID_month的工号列表中找到相同工号的位置，并在对应次数列表OffDay_count_month中取出次数并累加
# 然后将该工号OffDay_unduplicate_ID_month 和 累加后的次数OffDay_accum_count_month 分别放入新的两个列表中（也可得到一个字典）
OffDay_accum_count_month = []
for x in range(0, len(OffDay_unduplicate_ID_month)):
    repeat_location_offDay = find_repeat(OffDay_ID_month, OffDay_unduplicate_ID_month[x])
    accum_count = 0
    for y in range(0, len(repeat_location_offDay)):
        accum_count = accum_count + OffDay_count_month[repeat_location_offDay[y]]
    OffDay_accum_count_month.append(accum_count)

#print("OffDay_unduplicate_ID_month:", OffDay_unduplicate_ID_month, "OffDay_accum_count_month:", OffDay_accum_count_month)



#####################################  Step 2：根据不同标准对OffDay_unduplicate_ID_month中员工发放餐补  ##################################
# OffDay_unduplicate_ID_month是有餐补的员工的list，OffDay_fee是对应费用的list
for n in range(0, len(OffDay_unduplicate_ID_month)):
    if OffDay_unduplicate_ID_month[n][0] == "S" or OffDay_unduplicate_ID_month[n][0] == "i":
        OffDay_accum_count_month[n] = 18 * OffDay_accum_count_month[n]     # 实习生&非生产外包： 18元/餐
    else:
        OffDay_accum_count_month[n] = 18 * OffDay_accum_count_month[n]     # 员工：18元/餐
OffDay_accum_fee_month = OffDay_accum_count_month

# 将两个list变成字典可以一一对应
OffDay_ID_fee = dict(zip(OffDay_unduplicate_ID_month, OffDay_accum_fee_month))
print("OffDay_ID_fee",OffDay_ID_fee)


######################################  【工作日、节假日餐补总合】 #########################################
# 将工作日餐补和节假日餐补的ID和fee分别合并到两个列表中
WorkOffDay_ID_month = workDay_unduplicate_ID_month + OffDay_unduplicate_ID_month
WorkOffDay_fee_month = workDay_accum_fee_month + OffDay_accum_fee_month
#print(WorkOffDay_ID_month, WorkOffDay_fee_month)

# 用集合表示工号列表，去重复；然后再将集合变成列表，稳定不会改变
WorkOffDay_unduplicate_ID_month = set(WorkOffDay_ID_month)
WorkOffDay_unduplicate_ID_month = list(WorkOffDay_unduplicate_ID_month)
#print("WorkOffDay_unduplicate_ID_month:",WorkOffDay_unduplicate_ID_month)


# 遍历WorkOffDay_unduplicate_ID_month中的工号，在原始WorkOffDay_ID_month的工号列表中找到相同工号的位置，并在对应次数列表WorkOffDay_fee_month中取出费用并累加
# 然后将该工号WorkOffDay_unduplicate_ID_month 和 累加后的费用WorkOffDay_fee_month 分别放入新的两个列表中（也可得到一个字典）
WorkOffDay_accum_fee_month = []
for x in range(0, len(WorkOffDay_unduplicate_ID_month)):
    repeat_location_WorkOffDay = find_repeat(WorkOffDay_ID_month, WorkOffDay_unduplicate_ID_month[x])
    accum_fee_WorkOff = 0
    for y in range(0, len(repeat_location_WorkOffDay)):
        accum_fee_WorkOff = accum_fee_WorkOff + WorkOffDay_fee_month[repeat_location_WorkOffDay[y]]
    WorkOffDay_accum_fee_month.append(accum_fee_WorkOff)

#print("WorkOffDay_unduplicate_ID_month:", WorkOffDay_unduplicate_ID_month, "WorkOffDay_accum_fee_month:", WorkOffDay_accum_fee_month)

# 将两个list变成字典可以一一对应
WorkOffDay_ID_fee = dict(zip(WorkOffDay_unduplicate_ID_month, WorkOffDay_accum_fee_month))
print("WorkOffDay_ID_fee",WorkOffDay_ID_fee)


########################## 【导出数据】 ###################
# 将字典导出为excel
result_excel = pd.DataFrame()
result_excel["ID"] = WorkOffDay_unduplicate_ID_month
result_excel["fee"] = WorkOffDay_accum_fee_month
result_excel.to_excel(r'SZ_WorkOffDay(ID-Fee).xlsx')

print(result_excel)


# 程序完成，保存好了excel文件
print("Finshed")
