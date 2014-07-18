csv2xls
=======

csv file to excel file

依赖库: 
  xlwt  可以通过pip或者easy_install安装

应用:
  1， 初始化一个Excel实例
    obj = Excel()
    可以指定字体类型，以及是否粗体等样式，放在dict中
    eg:
      font_style = dict(bold=True, height=240)
      obj = Excel(**font_style)
    
    可以设置的样式有: bold, height, italic, struck_out, outline, shadow, 
                   colour_index, _weight, escapement, underline, 
                   family, charset, name
    样式详解见：https://github.com/python-excel/xlwt/blob/master/xlwt/Formatting.py
    
  2, 将csv写到xls文件里
    obj.csv_to_xls(save_path, filename, csvfile=[])
    参数说明:
    save_path为生成的xls文件存放的目录, 如果该目录不存在，则会自动创建该目录。
    filename为生成的xls文件的名字, 最后生成的xls文件的路径为 save_path + filename.
    csvfile为csv文件的地址, 可以指定多个csv文件。不同csv文件的内容将会在新的xls文件不同的sheet里展示。
