# Excel-Data-Processing-Tools




## Excel 文件数据合并工具

---

### 简介

本项目提供了一个Python脚本，用于合并指定文件夹下的所有Excel文件数据合并到一个新的工作簿，并应用统一的字体样式。
使用方法
安装依赖
在使用本工具之前，需要确保已安装以下Python库：

- openpyxl
可以使用以下命令安装这些依赖：
```Bash
pip install openpyxl
```

运行脚本
将脚本保存为 merge_excel.py。
修改脚本中的 directory 变量，设置为包含Excel文件的文件夹路径。
修改脚本中的 output_file 变量，设置为合并后输出的Excel文件名。
在命令行中运行脚本：

```Bash
python merge_excel.py
```



### 代码说明

💯主要功能

- 定义默认字体样式：define_default_font 函数定义了一个默认的字体样式。
- 应用默认样式到工作簿：apply_default_font_to_workbook 函数将默认字体样式应用到整个工作簿的所有单元格。
- 创建模板工作簿：create_template_workbook 函数创建一个具有默认样式的模板工作簿。
- 合并Excel文件：combine_excel_files 函数遍历指定文件夹下的所有Excel文件，将其合并到一个新的工作簿中，并应用默认样式。

🫨注意事项

- 确保指定的文件夹路径存在且包含有效的Excel文件。
- 输出文件名应以 .xlsx 结尾。
- 脚本会忽略 UserWarning 类型的警告，以减少不必要的输出。
- 希望本工具能帮助您高效地合并Excel文件！如果有任何问题或建议，可以留言哦。
- 程序使用 
