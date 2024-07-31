# 批量查找与替换Word文档内容

本人在ZSHC实习时无意碰到一难题：如何查找海量word文件中的关键词，并替换它们。后在网上冲浪，寻得解决方法，有感而发，故写下此文。希望能为大家提供一些帮助。本人能力有限，如有差错，敬请指正。                    *by Dinooosor*

## 一、需求场景

在工作中我们可能会遇到需要对大量word文件中的文字内容进行批量查找替换的情况。本文借鉴了https://blog.csdn.net/chinajavafan/article/details/135761133文章的内容，并结合自己的实践经验，整理出解决方法。

## 二、实践操作

1.使用AnyTXT Searcher软件（下载链接：https://anytxt.net/），搜索存在关键词的文档，并将需要处理的文件放入同一文件夹中

![这是图片](E:\Git-learn\使用VBA批量处理多文档内容\figure\0.jpg)

2.新建一个word,选择“宏”选项

![这是图片](E:\Git-learn\使用VBA批量处理多文档内容\figure\1.jpg)

3.进入“宏”选项，点击创建，输入宏名“批量修改”（可以自己取）

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\2.jpg)

4.将宏代码粘贴进去，注意开头的格式修改（新建宏会自动带Sub 与 End Club,注意最后代码中只能保留一个），并保存。



```visual basic
Private Sub CommandButton1_Click()
'关闭屏幕刷新
 
Application.ScreenUpdating = False
 
'定义变量
 
Dim myFile$, myPath$, i%, myDoc As Object, myAPP As Object, txt$, Re_txt$
 
'设置应用对象，建立临时进程
 
Set myAPP = New Word.Application
 
'显示选择文件夹对话框
 
With Application.FileDialog(msoFileDialogFolderPicker)
 
    .Title = "选择目标文件夹"
 
    If .Show = -1 Then
 
        '读取选择的文件路径
 
        myPath = .SelectedItems(1)
 
    Else
 
        Exit Sub
 
    End If
 
End With
 
'文件夹目录和文件完整路径
 
myPath = myPath & ""
 
myFile = Dir(myPath & "\*.docx") '这里用的是docx,所以适用的文档只能是docx
 
'获取被替换的文字
 
txt = InputBox("需要替换的文字:")
 
'获取替换文件
 
Re_txt = InputBox("替换成:")
 
'显示打开文档
 
myAPP.Visible = True '是否显示打开文档
 
'循环处理文件夹中的全部文件
 
Do While myFile <> "" '文件不为空
 
    '打开文件
 
    Set myDoc = myAPP.Documents.Open(myPath & "\" & myFile)
 
    '判断文件是否受保护，仅对未受保护的文件有效
 
    If myDoc.ProtectionType = wdNoProtection Then
 
        '查找替换
 
        With myDoc.Content.Find
 
            .Text = txt
 
            .Replacement.Text = Re_txt
 
            .Forward = True
 
            .Wrap = 2
 
            .Format = False
 
            .MatchCase = False
 
            .MatchWholeWord = False
 
            .MatchByte = True
 
            .MatchWildcards = False
 
            .MatchSoundsLike = False
 
            .MatchAllWordForms = False
 
            .Execute Replace:=2
 
        End With
 
    End If
 
    '设置文件窗口状态，避免再次打开时被隐藏
 
    Application.WindowState = wdWindowStateNormal
 
    '保存并关闭文档
 
    myDoc.Save
 
    myDoc.Close
 
    myFile = Dir
 
Loop
 
 '关闭临时进程
 
myAPP.Quit
 
'打开屏幕更新
 
Application.ScreenUpdating = True
 
'输出提示信息
 
MsgBox ("全部替换完毕!")
End Sub
```

>注意在上述代码中：myFile = Dir(myPath & "\*.**docx**") '里用的是docx,作者猜测只能适用于docx，对于处理doc类型文件，可改为：myFile = Dir(myPath & "\*.**doc**") '，实践证明可行。
>
>PS：作者不知道怎么同时查找docx、doc文档www，希望懂的大佬教一下

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\3.jpg)

5.点击运行

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\4.jpg)

6.作者本人在运行时跑出这一东西，点击禁用宏后仍能使用（希望懂的大佬解答一些www）

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\6.jpg)

7.选择对应的文件夹（在第一步我们创建的文件夹），点击确定

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\7.jpg)

8.选择需要替换的内容

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\8.jpg)

9.选择需要替换成的内容

![图片2](E:\Git-learn\使用VBA批量处理多文档内容\figure\9.jpg)

10.显示替换成功

> PS：本次实验样本数量较小，大概30余篇，尚未在大样本环境下实践，故不能保证在海量文档下该方法完全有效。





Excel中有冻结窗格
