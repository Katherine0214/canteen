# canteen
>本算法包用于实现Excel表格数据分析（用于发放餐补）。

>用于测试的Excel表格暂没上传，可见本地。

### 🛠️ 一、进入环境（本地已配置好）
运行```conda activate env_canteen```

### 👾 二、SH（工作日午餐）  
### ```oneMonth_SH.py```
#### 输入： 

(1) ```移动和5G打卡记录表格.xlsx```  （81是测试表格）

(2) ```公司内打卡记录表格.xlsx```  （82是测试表格）

(3) ```月份 month = "2022.08" ```  （这个只需每月改一次）

(4) ```当月工作日的日期 day = [1,2,3,7,8,10,15]```  [ ]内的内容需要自己输入，注意用英文，隔开


#### 输出： 

```工号-餐补 对应的dict，并保存为 “SH(ID-Fee).xlsx” ```  (当程序中看到打印出“Finshed”即为完成)

### 🎲 三、SZ（工作日午餐+加班晚餐+节假日餐）   
### ```oneMonth_SZ_WorkOffDay.py```
#### 输入： 

(1) ```公司内打卡记录表格.xlsx```   （3是测试表格）

(2) ```月份 month = "2022.08"```  (这个只需每月改一次)

(3) ```当月工作日的日期 work_day = [1,2,3,7,8,10,15]```  [ ]内的内容需要自己输入，注意用英文，隔开

(4) ```当月周末和节假日的日期 off_day = [1,2,3,7,8,10,15]```  [ ]内的内容需要自己输入，注意用英文，隔开

#### 输出： 

```工号-餐补 对应的dict，并保存为 “SZ_WorkOffDay(ID-Fee).xlsx”```


