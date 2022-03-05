# Docx_Search
Docx paragraph searching \replaceing；table searching\copy table contents……

I am going to use Chinese only as the main lang under this line.

github resourcehere: [kasusa/Docx_Search](https://github.com/kasusa/Docx_Search)

---

## 简介

本库主要适用于docx文件的修改、段落搜索、表格搜索等，填补了官方未提供的“搜索”功能的空白。

功能主要包括：

- word中所有图片提取为bitmap
- 搜索段落
- 搜索表格
- 删除段落
- 获取单元格（cell）的内容
- 批量替换
- 保存到桌面 out文件夹

## 依赖

本库是基于开源*非商用的docx库编写的。引用如下两个库，可以在这里下载：[xceedsoftware/DocX](https://github.com/xceedsoftware/DocX)，这个项目提供了一些非常实用的例子，建议新手先看他们的例子来学习c#操作word的逻辑。

```
using Xceed.Document.NET;
using Xceed.Words.NET;
```

## 创建对象

我把官方的document对象又封装了一遍，这样使用函数的时候更加直观。

```cs
myutil tempo;//我的word对象
//docx文件的路径
string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @$"\Sample\方案\方案A.docx";
if (System.IO.Directory.Exists(tempo_path))
{
    //初始化对象
    tempo = new myutil(a);
}
```

## 搜索段落

搜索段落如同word中一样，使用string作为关键字进行搜索，实现下来整体速度还不错。

有的时候一篇word中搜索一个关键词可能有好几个结果，这种时候返回list的就派上用场了。

有的时候可能要操作的段落不好定位，但是可以确定他在某个标题的下面，index搜索就派上用场了，返回的index+1差不多就是需要的段落位置了。

注意：xceed.docx的实现比较完善，可以搜索到目录、表格中的段落，这是python-docx很难实现的。

我写了几种类型的搜索函数：

```cs
Find_Paragraph_for_p(string v) 					//搜索后返回paragraph
Find_Paragraph_for_plist(string v) 				//搜索后返回paragraph列表
Find_Paragraph_for_text(string v) 				//搜索后返回string（段落的文字内容）
Find_Paragraph_for_i(string v) 					//搜索后返回段落的index
Find_Paragraph_for_ilist(string v) 				//搜索后返回段落的index list
```

举一个实际的例子：

```cs
//搜索“本报告记录编号：”，返回段落的text
tmpstr = doc.Find_Paragraph_for_text("本报告记录编号：");
//获取到：本报告记录编号：P2021XXXXX-GB01。
//处理切分一下字符串，获取到我需要的部分
tmpstr = myutil.get_string_after(tmpstr, "本报告记录编号：", "P2021XXXXX".Length);//结果：P2021XXXXX
```

再举一个例子：

```cs
//搜索总体评价（一级标题），获取其后面段落的内容
int i = doc.Find_Paragraph_for_ilist("总体评价")[0] + 1;
a = doc.document.Paragraphs[i].Text;
```

搜索的时候如果得到的结果不是很正常，就要看看word中是不是又多个能搜到的位置了。

## 搜索表格

搜索表格是一个我自创的功能，他是利用表头（首行）的文字内容对整个word中的所有表标题进行检索,如果某一个表格的第一行和给的参数相同，就认为是查找到了这张表。

经过我的实际测试，速度还算可以。

```cs
//寻找表头文字为“序号	机房名称	物理位置	重要程度”的第一张表
//注意这里我没有删除空格和tab，这是为了从word中复制过来方便，实际在函数中会把空格、tab字符都删除后进行对比
var table1 = doc.findTableList("序号	机房名称	物理位置	重要程度")[0];
```

下面是我实现的一个表格复制函数。从word1中把t1复制到t2，可以选择是否包含表头。

```cs
#region 表格复制函数
/// <param name="t1head">table1 表头</param>
/// <param name="i1">table1 所在index</param>
/// <param name="t2head">table2 表头</param>
/// <param name="i2">table 2 所在index</param>
void CopyTable(string t1head, string t2head, int i1 = 0, int i2 = 0)
{
    bool toremove = false;
    var table1 = doc.findTableList(t1head)[i1];
    ConsoleWriter.WriteColoredText("table 报告中 ↑", ConsoleColor.Green);
    var table2 = tempo.findTableList(t2head)[i2];
    ConsoleWriter.WriteColoredText("table 模板中 ↑", ConsoleColor.Green);
    //如果t1比t2更宽，增加一列临时列
    if (table1.ColumnCount > table2.ColumnCount)
    {
        table2.InsertColumn();
        toremove = true;
    }
    //如果t1比t2更窄，直接给t2瘦身
    else if (table1.ColumnCount < table2.ColumnCount)
    {
        table2.RemoveColumn(table2.ColumnCount-1);
    }
    //删除所有空的内容行
    while (table2.RowCount > 1)
    {
        table2.RemoveRow(table2.RowCount - 1);
    }
    //从内容行数开始复制
    for (int i = 1; i < table1.RowCount; i++)
    {
        Xceed.Document.NET.Row row = table1.Rows[i];

        table2.InsertRow(row);
    }
    //删除多复制过来的列
    if (toremove)
    {
        table2.RemoveColumn(table2.ColumnCount - 1);
    }
    ConsoleWriter.WriteColoredText("复制表完毕;", ConsoleColor.Yellow);

}

/// <summary>
/// 复制表t1到表t2（包含表头）
/// </summary>
/// <param name="t1head">table1 表头</param>
/// <param name="i1">table1 所在位数</param>
/// <param name="t2head">table2 表头</param>
/// <param name="i2">table 2 所在位数</param>
void CopyTable_withHead(string t1head, string t2head, int i1 = 0, int i2 = 0)
{
    bool toremove = false;
    var table1 = doc.findTableList(t1head)[i1];
    ConsoleWriter.WriteColoredText("table 报告中 ↑", ConsoleColor.Green);
    var table2 = tempo.findTableList(t2head)[i2];
    ConsoleWriter.WriteColoredText("table 模板中 ↑", ConsoleColor.Green);
    //如果t1比t2更宽，增加一列临时列
    if (table1.ColumnCount > table2.ColumnCount)
    {
        table2.InsertColumn();
        toremove = true;
    }
    //如果t1比t2更窄，直接给t2瘦身
    else if (table1.ColumnCount < table2.ColumnCount)
    {
        table2.RemoveColumn(table2.ColumnCount - 1);
    }
    //删除所有空的内容行
    while (table2.RowCount > 1)
    {
        table2.RemoveRow(table2.RowCount - 1);
    }
    //从内容行数开始复制
    for (int i = 0; i < table1.RowCount; i++)
    {
        Xceed.Document.NET.Row row = table1.Rows[i];

        table2.InsertRow(row);
    }
    //删除多复制过来的列
    if (toremove)
    {
        table2.RemoveColumn(table2.ColumnCount - 1);
    }
    //删除顶部的原始行（表头总是有问题服了）
    table2.RemoveRow(0);
    ConsoleWriter.WriteColoredText("复制表完毕;", ConsoleColor.Yellow);

}

#endregion
```

## 文字批量替换

文字批量替换是我参考官方文档的例子写出来的，因为功能很常用所以把他包含到了库中。

```csharp
doc._replacePatterns.Add("P2021xxxxx", "P202100001");//向字典里面添加怕【被替换】、【替换成】的字符串
doc._replacePatterns.Add("AAAAA", "可口可乐公司");

doc.ReplaceTextWithText_all_noBracket();//自动搜索文档中所有在字典里面的内容并替换
```

