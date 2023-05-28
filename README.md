# Import-excel-config-to-WinCC-trend-and-data-control
**中文**

 将配置好的Excel文件中的变量和属性应用到WinCC的趋势控件和数据控件，以加快组态速度避免重复劳动。 
 
 当打开趋势或数据页面时，会检测当前界面的控件是否有内容，如果没有内容则加载Excel文件中定义好的标签、标签注释和对应的归档记录。
 
 当内容加载完成后，默认显示前三项归档。界面下方提供了筛选功能，可以根据标签注释来筛选需要显示的内容。
 
 当数据比较多时，加载速度较慢，建议自行增加提示窗口，告知操作员正在加载数据，避免切换界面导致加载不完整。
 
 例子使用方法：
 将所有文件拷贝到新建项目的GraCS文件夹，并将Start.pdl设置为启动画面，然后启动WinCC运行时。
 
 提供的功能：
 1. 读取配置好的Excel文件，同时应用到WinCC的趋势控件和数据控件
 2. 筛选趋势控件和数据控件的内容
 3. 筛选包含和（或）不包含用户输入文本的趋势控件和数据控件的内容，支持正则表达式
 
 
 需要注意：Excel文件必须为.xls格式（即Excel 2003及之前的格式，这样可以在Windows 7和更新的操作系统上，不需要安装额外的软件或微软Office）
 
------------
**English**

Apply the tags and properties in the configured Excel file to the trend control and data control of WinCC to speed up the configuration and avoid duplication of labor.

When opening a trend or data page, it will detect whether the control on the current interface have content, and if there is no content, load the tags, tag comments and corresponding archive records defined in the Excel file.
  
When the content is loaded, the first three archives are displayed by default. The filter function is provided at the bottom of the interface, and the content to be displayed can be filtered according to the label comments.
  
When there is a lot of data, the loading speed is slow. It is recommended to add a prompt window to inform the operator that the data is being loaded, so as to avoid incomplete loading caused by switching pictures.
 
Example usage:
  
Copy all files to the GraCS folder of the new project and set Start.pdl as the start picture, then start WinCC Runtime.
 
 
 Features provided:
  1. Read the configured Excel file and apply it to the trend control and data control of WinCC at the same time
  2. Filter the content of trend control and data control
  3. Filter the content of trend control and data control that contain and/or not contain user input texts, support regular expressions
  
  
 Note: the excel file must be in .xls format (i.e. the format of Excel 2003 and before, so that it can be used on Windows 7 and newer operating systems without installing additional software or Microsoft Office)
