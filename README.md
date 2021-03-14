课表PDF转ICS
==

简介
---
ICS文件可以导入至WIN10日历、手机各种内置日历应用等。  
导入完成后，同学们可以在设备上快速查看课表，并在上课前收到及时的提醒。  


本项目代码可以将从学校新教务系统下载的PDF课表文件转换成ICS日历文件。


安装环境
---
**Python 3.6以上**  
**Python第三方库pdfplumber**



因为通过pip安装pdfplumber会自动安装依赖包pdfminer.six，而pdfminer.six会导致读取课表内容失败，将pdfminer.six换成pdfminer后便没有问题。所以**须按顺序执行下述命令**。
>`pip3 install pdfplumber`  
>`pip3 uninstall pdfminer.six`  
>`pip3 install pdfminer`  

使用前准备
---
```python
# 需要在脚本中修改以下信息
# 严格意义上此处的开学日是指该学期第一周的星期一

Year_Begin = '202x'  # 开学年份

Month_Begin = 'xx'  # 开学月份

Day_Begin = 'xx'  # 开学日

Pdf_Path = 'xxx(20xx-20xx-x)课表.pdf'  # 相对路径或绝对路径

Trigger_Time = 20  # 提前多少分钟提醒（默认为20分钟，此数值不建议调成太大）
```


Table.py
---
本脚本适用于表格形式的PDF课表文件

List.py
---
本脚本适用于列表形式的PDF课表文件

