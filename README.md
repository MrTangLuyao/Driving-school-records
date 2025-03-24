# Driving-school-records
驾校教练使用的学员上课安排的规划程序<br>

## 如何记录：<br>

本软件使用excel作为数据记录，template内涵排序以及查询公式，可能需要**Microsoft 365最新版**才能顺利运行。<br>

### 运行 OKPASSGui 来记录你的安排：<br>

![image](https://github.com/user-attachments/assets/1d5e4367-35ae-4c46-a87d-39fd0b84bace)<br>


### 点击添加记录来添加一条记录：
使用如下格式添加一条记录
#### 需要注意 日期之间使用/分割 时间段的 - 左右两边有空格， 需保证**格式完全正确**！！！
![image](https://github.com/user-attachments/assets/4cd4b00a-fce5-45f5-a4d8-9b264778f282)<br>

### 点击修改记录来修改某位学员的记录<br>

![image](https://github.com/user-attachments/assets/4bd95c1c-8b68-4315-b2bd-21862ebe5755)<br>

输入学员姓名后点击**OK**<br>

![image](https://github.com/user-attachments/assets/303d59b7-9e7f-43ce-817a-a42603856e3b)<br>

找到你需要修改的那条记录，选中并点击 **选择** 按钮<br>

![image](https://github.com/user-attachments/assets/94def7d5-2600-4caf-8ec3-a5d97d268441)<br>

在编辑页面可以更改你需要的时间以及日期，编辑结束点击**保存**。<br>

### 点击归档文件可以将这个excel表格，需要注意，由于excel公式限制，一个表格只能记录99条信息，这也是为何需要设计归档功能<br>

点击归档按钮<br>

![image](https://github.com/user-attachments/assets/d3f0e0e7-7f00-44f0-bf54-25a9f9ad552c)<br>

点击 **是（Y）以确认**

![image](https://github.com/user-attachments/assets/5d128919-4c9e-419d-8dd4-8ff6cb6b3fe8)<br>


归档后会文件会被归档到当前文件夹下的OldRecords目录下，并加上今天的时间来进行归类

![image](https://github.com/user-attachments/assets/747774c2-cb7f-4302-9882-82a8933cf8a7)<br>

### 修改完成后点击退出系统

![image](https://github.com/user-attachments/assets/d95a29c5-0871-44cb-ac14-edbcc5ec4016)<br>

### 日志功能
![image](https://github.com/user-attachments/assets/d689a388-b480-4021-8022-7eb55800e4ad)<br>


软件会记录所有的操作日志，但不会保存到文件中，点击退出系统日志将会被清除。<br>

### 关于 OKPASSCmd<br>
与OKPASSGui**功能相同**，但是为命令行来交互，可能不直观但是稳定性更高。建议在Gui版本出错时使用。<br><br><br>



## 如何查看将来的日程安排？<br>

双击 **records.xlsx** 查看已经记录的数据<br>

界面如下：
![image](https://github.com/user-attachments/assets/b4c32858-7327-4dd5-bb22-0b6c1d503b03)<br>

内容分为四个表：**Record SortSheet Today Lookup**<br>

![image](https://github.com/user-attachments/assets/795e0470-bce5-447a-ae37-2a6948c3dfcb)<br>

第一个表Record为记录数据，由记录程序记录，也可以手动修改需要的值。<br>

第二个表SortSheet为运算数据，用于排序以及筛选，用户不需要访问SortSheet，也不要删除SortSheet<br>

**第三个表Today为今天的安排，用于查看今日日程安排：**
![image](https://github.com/user-attachments/assets/8b32db8e-8e0c-42a8-99cc-138926f1bbfd)<br>

**第四个表Lookup用于查询特定日期的日程安排：**
![image](https://github.com/user-attachments/assets/608ee4f5-9a61-4831-8d5b-c24a25368b2a)
使用方法：在LookFor后面填写你需要查询的日期，格式DD/MM/YYYY <br><br>

## 注意：
1- Record的D行应该被隐藏，从中填写的是排序种子（SortPeer），此栏不应该展示也不应该被用户修改<br>
2- 用户不应该更改除了Record表以外其他表格信息，否则可能导致运行异常<br>

## 文件结构：
![image](https://github.com/user-attachments/assets/5d65a74a-d731-4fe8-8fea-1759819edd0a)<br>
记录程序：OKPASSGui 或者 OKPASSCmd<br>
模板excel文件：于template目录内，用户不应该修改或者删除此文件夹<br>
归档的excel文件：OldRecords<br>
正在修改以及编辑的日程表文件：Records.xlsx<br><br>


# 已知的Bug：
在确实records.xlsx文件时，如启动OKPASSGui版本，程序可以顺利新建默认的Records.xlsx但是程序第一次启动会报错，后续运行正常。如用户出现类似报错无需紧张，重新运行软件即可。












