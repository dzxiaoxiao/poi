## 使用说明：

- 创建ExportExcel对象

- 用ExportExcel实例对象的方法<method>createExcel</method>创建Excel对象

- 画表格之前可以对表格进行设置样式
> <method>createTableHeaderFont</method> 获取表头字体样式对象，对表头内容进行自定义字体样式 <br/>
> <method>createTableBodyFont</method> 获取表体字体样式对象，对表体内容进行自定义字体样式
> <method>setAddBorder</method> 设置表格是否添加边框
> <method>setAddTableHeaderBorder</method> 设置表格的表头是否添加边框

- 组装表头结构[对于多级表头会进行合并居中，如果是map或者实体类，field字段一定要匹配。否则取不到数据就会抛出异常]
> TableHeader 表头对象 注：field支持多级取值 例如：a.b[n].c | a.b.c | a[n].b.c

- 将表格填充至Excel
> <method>drawTable</method> 这个方法提供了两个实现方式：
> 第一种：以追加的形式将表格填充至Excel，两个表格之间默认间隔两行。
> 第二种：以指定下标的形式将表格填充至Excel指定位置。

- 将Excel写出至指定磁盘路径

- 这里额外提供了一个方法<method>setCellBackGround</method>可以设置指定区域内所有单元格的背景色
