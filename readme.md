# 1. 参数介绍：

    mode = "点名册"     # ['点名册', '个人成绩'] 两种模式
    id_column_name = "学号"  # Excel表中的一个关键词，用于索引获取数据
    name_column_name = "姓名"  # # Excel表中的一个关键词，用于索引用于索引获取数据
    images_dir = "E:\Pycharm_code\My_code"  # 图片文件夹路径，最好是按文件创建时间递增排序
    excel_path = "E:\Pycharm_code\My_code\\111.xlsx"   # Excel表的路径

# 2. 注意事项：
1. 保证试卷（图片）的顺序与Excel表中人名顺序一致
2. 图片必须是竖直方向，否则需要原地进行旋转（不可另存为，因为会改变创建时间）
3. 如果存在缺考，将缺考人在Excel表中直接删除
