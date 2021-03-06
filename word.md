# 正则表达式

## 常见内容的正则表达式

```
^? 任意字符
^# 任意数字
^$ 任意字母
^p 段落标记
^l 软回车
^t 格式标记
^d 域
^w 空白区 (空格、不间断空格、以及任意顺序的格式标记)
^f 脚注标记
^e 尾注标记
^b 分节符
^& 原查找内容
^c 剪贴板内容
```

来源： [如何快速地排版一篇排版很乱的Word？](https://www.zhihu.com/question/24709866)

## 应用实例

- 清除多个软回车：软回车是 Shift + Enter 产生的，在替换命令中表示为 ^l 。
先执行此步，是为了保证后面两步替换空格时，覆盖所有的情况。如果此处的软回车不替换，下面的两步就会漏掉软回车前后的空格。

- 替换行首的连续空格（普通空格，全角空格等）

```
^p^w
^p
```

- 替换行尾的连续空格

```
^w^p
^p
```

- 替换连续多个换行符（至少两个）为单个换行符（使用通配符）

```
(^13){2,}
^13
```

- 替换数字列表后的全角点号为半角点号（使用通配符）

```
([0-9]@)([．.])([! 0-9])
\1. \3
```

- 替换内含至少一个空格的中英文括号对为内含四个空格的中文括号对（使用通配符）

```
([（\(])( @)([）\)])
（    ）
```

- 替换`字母A~E+中英文点号+无空格`为`字母A~E+英文点号+一个空格`（使用通配符）

```
([A-E])([ ．.])([! ])
\1. \3
```

- 替换`字母A~E+中英文点号+至少一个空格`为`字母A~E+英文点号+一个空格`（使用通配符）

```
([A-E])([ ．.])( )@
\1. 
```

- 搜索方括号中包含1~3位阿拉伯数字的情况（即参考文献序号）

```
\[([0-9]{1,3})\]
```

# 特殊内容查找

## 计算文档中批注数量

点击左上角“快速访问工具栏”右侧的下拉三角箭头，再点击“其他命令”，在弹出的对话框中，按照下图的方法将指定的命令添加到“快速访问工具栏”中。

然后定位到文档中最后一个批注那里，看批注的编号即可。

![](https://raw.githubusercontent.com/Dream4ever/Pics/master/word-add-print-review-edit-mode.png)

参考资料

- [Numbers for Referencing Review Comments in Word](https://superuser.com/a/720801/432588)

# MathType

编辑论文公式的时候发现自动编号出故障了，用`Modify Break`把第一章的MathType分隔符由`Next Chapter`改为了`Chapter Number`，这样指定了第一章之后，后面的各章编号就正常了。
用MathType的`Right-numbered Equation`功能来插入带编号的公式时，分割形式为`(1.1)`，需要将其改成`(1-1)`这样的编号，在MathType选项卡的`Equation Numbers`选项下，点击`(1) Insert Number`选项右边的倒三角，选择`Format`选项，将`Seperator`这一项右边的内容从点号`.`改成短横线`-`即可。
MathType插入的用大括号括起来的两个公式，下面的公式太长，为了美观通过手动输入回车换行，但是上面的公式和括号的上端之间有一段距离，在MathType中选中公式（不要选中左边的大括号），在格式中设置为`对齐到顶端`或`对齐到底端`即可。

# 疑难杂症

有同事的部分 Word 文档在打开时报错，经研究，发现是 Word 文件被锁定了，按照 [How to resolve the problem "Word experienced an error trying to open the file" when opening a Word 2007/2010 file (Easy Fix Article)](https://support.microsoft.com/en-us/help/2749199/how-to-resolve-the-problem-word-experienced-an-error-trying-to-open-th) 这篇文章中所说的方法，将文档解锁，就可以正常打开了。
