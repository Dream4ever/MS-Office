# 编辑

- 横向或者纵向选定若干个单元格，直接输入想要的数值，然后按`Ctrl+Enter`，就会用该值填充所有单元格了。

# 查找

- 想知道某个单元格的内容，是公式计算后得到的值，还是直接输入的值？用 `ISFORMULA()` 函数来检查就行了。

```
=ISFORMULA(A2)
```

参考链接：[How to tell if an Excel cell has a formula or is hardcoded?](https://superuser.com/a/1037251/432588)

- 在A列中查找包含B列各单元格中的字符串的单元格，如存在则返回A列对应单元格所在行的C列单元格的值。
  - `MATCH("*"&B2&"*",A:A,0)`是在A列中，查找包含B2单元格字符串的单元格。
  - `INDEX(C:C,MATCH("*"&B2&"*",A:A,0))`，则根据之前的查找结果，获取对应行在C列单元格中的值。

- 要获取字符串在单元格中首次出现的位置，用`FIND`函数即可：`FIND("John", A2)`。

- 要判断是否有重复的行，用下面的公式即可。该公式判断第2行至第10行，是否有A、B、C列的值均相同的行。

```
 =IF(SUMPRODUCT(($A$2:$A$10=A2)*1,($B$2:$B$10=B2)*1,($C$2:$C$10=C2)*1)>1,"Duplicates","No duplicates")
 ```
 
 参考链接：[How To Find And Select Duplicate Rows In A Range In Excel?](https://www.extendoffice.com/documents/excel/1352-excel-find-duplicate-rows.html)

- 返回最后一个日期单元格其右侧单元格的数值：=VLOOKUP(I3, A2:B520, 2, TRUE)。I3为目标日期单元格日期对应的数值，A列为日期，B列为其它数据，TRUE表示如果能找到目标日期单元格，则返回该行某一列，否则返回上一行某一列。
- 返回最后一个非空单元格的位置：=LOOKUP(1,0/(B2:B65536<>""),ROW(B2:B65536)) （03好像不支持整列）
- 返回最后一个非空单元格的数值：=LOOKUP(1,0/(B2:B65536<>""),B2:B65536)
- 返回最后一个数值单元格的位置：=LOOKUP(9.9E+307,A:A,ROW(A:A))
- 返回最后一个数值单元格的数值：=LOOKUP(9.9E+307,A:A)
- 返回最后一个文本单元格的位置：=LOOKUP(REPT("座",255),B:B,ROW(B:B))
- 返回最后一个文本单元格的数值：=LOOKUP(REPT("座",255),B:B)

# 生成

- 生成一组分布在指定区间，并且有指定均值的随机数：

```
=NORMINV(RAND(),MEAN,standard_dev)
```

参考链接：[Generate Random Number By Given Certain Mean And Standard Deviation In Excel](https://www.extendoffice.com/documents/excel/2472-excel-random-number-mean-standard-deviation.html)

# 转换

- 将2010.08.03格式的数据转换为日期格式，由于Excel将2010-08-03格式的数据识别为日期格式，所以将 . 替换成 - 即可。

# 比较

- Excel需要比较两个单元格内容是否相同，用A1=B1这种公式即可。

# 统计

- 将每日数据记录表按月求和并生成曲线
> 1. 对日期列和数据列创建数据透视表
> 2. 日期作为行，数据列作为值
> 3. 右键对日期列创建组，选择月和年，即生成各月和各年的统计数据。

- 自定义排序的结果如果不是自己想要的，可能是主要关键字和各个次要关键字的顺序设置错了。

# 图表

- 一个簇状柱形图上需要显示两个数据系列而且不能重叠，可按下面的方式设置。

|客户|销量|销售额|
|--|--|--|
|A|1| |
|B|2| |
|A| |3|
|B| |4|

- 对簇状条形图手动修改了数据标签的值之后，再更改图所对应的实际数据，数据标签的值也不会变。这样可以使图中的各个数据条按照自己需要的比例显示。
- 新建数据透视表时，提示“数据透视表字段名无效，必须使用组合为带有标志列列表的数据。如果要更改数据透视表字段的名称，必需建入字段的新名称。”，这时将没有内容的标题行写上内容即可解决。

# 浏览

- 工作表较多，需要切换的时候，在底部的左右箭头图标上右击，在弹出的“激活”窗口中切换即可。

# 其它

- 每一个Excel文件称为一个工作簿(workbook)，一个工作簿中的每个sheet称为工作表(worksheet)。
