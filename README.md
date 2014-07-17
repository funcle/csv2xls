csv2xls
=======

csv file to xls

依赖库: 
  xlwt  可以通过pip或者easy_install安装

应用：
  1， 初始化一个Excel实例
    obj = Excel()
    这里可以指定字体类型，以及是否粗体 obj = Excel(font_name='xx', bold=True)
    
  2, 将csv写到xls文件里
    obj.csv_to_xls(save_path, filename, csvfile=[])
    参数说明：
      save_path为生成的xls文件存放的目录, 如果该目录不存在，则会自动创建该目录。
      filename为生成的xls文件的名字, 最后生成的xls文件的路径为 save_path + filename。
      csvfile为csv文件的地址, 可以指定多个csv文件。不同csv文件的内容将会在新的xls文件不同的sheet里展示。
