# Import-excel-config-to-WinCC-trend-and-data-control
**中文**

 将配置好的Excel文件中的变量和属性应用到WinCC的趋势控件和数据控件，以加快组态速度避免重复劳动。 
 将所有文件拷贝到新建项目的GraCS文件夹，并将Start.pdl设置为启动画面，然后启动WinCC运行时。
 
 提供的功能：
 1. 读取配置好的Excel文件，同时应用到WinCC的趋势控件和数据控件
 2. 筛选趋势控件和数据控件的内容
 3. 筛选包含和（或）不包含用户输入文本的趋势控件和数据控件的内容，支持正则表达式
 
 
 需要注意：Excel文件必须为.xls格式（即Excel 2003及之前的格式，这样可以在Windows 7和更新的操作系统上，不需要安装额外的软件或微软Office）
 
------------
**English**

Apply the tags and attributes in the configured Excel file to the trend control and data control of WinCC to speed up the configuration and avoid duplication of labor.
 Copy all files to the GraCS folder of the new project and set Start.pdl as the start picture, then start WinCC Runtime.
 
 Features provided:
  1. Read the configured Excel file and apply it to the trend control and data control of WinCC at the same time
  2. Filter the content of trend control and data control
  3. Filter the content of trend control and data control that contain and/or not contain user input texts, support regular expressions
  
  
 Note: the excel file must be in .xls format (i.e. the format of Excel 2003 and before, so that it can be used on Windows 7 and newer operating systems without installing additional software or Microsoft Office)
