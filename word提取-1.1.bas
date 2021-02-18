Attribute VB_Name = "模块11"
'word提取 1.1
'修复name不存在报错的问题及其他问题；增加共享权力人身份证等信息填入；保留两位小数 20.8.13
'导出房屋照片 20.8.11

Public path As String
Public source As String

Public Function dad(name As String, num As String, i As Integer)

    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\00.档案袋.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "I-O"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10) & Sheet1.Cells(i, 11) & Sheet1.Cells(i, 12) & Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
      
    With oword.Selection.Find
        .ClearFormatting
        .Text = "B"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function zjtzs(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\01.指界通知书.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AA-5"
        .Replacement.ClearFormatting
        .Replacement.Text = Val(Sheet1.Cells(i, 27)) - 5
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AA"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    If Val(Sheet1.Cells(i, 28)) - 12 > 0 Then
        With oword.Selection.Find
            .ClearFormatting
            .Text = "AB-12正负判断"
            .Replacement.ClearFormatting
            .Replacement.Text = "下"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    Else
        With oword.Selection.Find
            .ClearFormatting
            .Text = "AB-12正负判断"
            .Replacement.ClearFormatting
            .Replacement.Text = "上"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    End If
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AB"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 28)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "□√本人/□单位法定代表人或负责人 /□√农民集体推举的指界人－ＡＹ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 51)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AY"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 51)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "J"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "K"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｋ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "L"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｌ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｍ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "Y"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 25)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "Z"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function zjwts(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\02.指界委托书.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BF"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 58)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BI"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 61)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "D"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 4)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "F"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 6)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "K"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "L"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function fwzp(name As String, num As String, i As Integer)
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\12.房屋照片.docx")
    
    Dim table As table
    Set table = doc.Tables(1)
    
    MyPath = path & name & "(" & num & ")\立面照\"   ' 指定路径 最后必须加上\
    myname = Dir(MyPath, vbDirectory)   ' 找寻第一项
    Dim j As Integer
    j = 1
    Do While myname <> ""   ' 开始循环
        ' 跳过当前的目录及上层目录
        If myname <> "." And myname <> ".." Then
            ' 使用位比较来确定 MyName代表一目录
            If InStr(myname, "png") Then
                t_row = table.Rows.Count
                If j > t_row Then
                    table.Rows(t_row).Cells(1).Select
                    oword.Selection.InsertRowsBelow 1
                    table.Rows(j).Cells(1).Range = Split(myname, ".png")(0)
                    table.Rows(j).Cells(2).Range.InlineShapes.AddPicture Filename:=MyPath + myname, LinkToFile:=False, SaveWithDocument:=True
                Else
                    table.Rows(j).Cells(1).Range = Split(myname, ".png")(0)
                    table.Rows(j).Cells(2).Range.InlineShapes.AddPicture Filename:=MyPath + myname, LinkToFile:=False, SaveWithDocument:=True
                End If
                j = j + 1
            End If
        End If
        myname = Dir    ' 查找下一个目录
    Loop
    
    doc.Save
    doc.Close
    
End Function
 
Public Function hzsms(name As String, num As String, i As Integer)
    
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\03.户主声明书.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "HIJ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 8) & Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "D"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 4)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "J"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "K"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "L"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function qqdj(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\05.农村宅基地使用权及房屋所有权确权登记申请表.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "O-S"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 15) & Sheet1.Cells(i, 16) & Sheet1.Cells(i, 17) & Sheet1.Cells(i, 18) & Sheet1.Cells(i, 19)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "I-M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10) & Sheet1.Cells(i, 11) & Sheet1.Cells(i, 12) & Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "G里：“、”+1 " '若不是原文档替换字符串，请修改这 todo
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 77)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    Dim arr
    Dim sarr
    Dim gyqlr As String
    Dim cur As Integer
    Dim mytable As table
    
    cur = 6
    gyqlr = Sheet1.Cells(i, 7)
    If gyqlr <> "" Then
    For Each mytable In doc.Tables
        arr = Split(gyqlr, "、")
        For j = 0 To UBound(arr)
            If InStr(arr(j), " ") Then
                sarr = Split(arr(j), " ")
                If Len(sarr) = 3 Then
                    mytable.Cell(cur, 2).Range = sarr(0)
                    mytable.Cell(cur, 4).Range = sarr(1)
                    mytable.Cell(cur, 3).Range = sarr(2)
                End If
                cur = cur + 1
            End If
        Next
    Next
    End If
    
    doc.Save
    doc.Close
    
End Function

Public Function bdcqdc(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\06.不动产权籍调查表.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "N-S"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 14) & Sheet1.Cells(i, 15) & Sheet1.Cells(i, 16) & Sheet1.Cells(i, 17) & Sheet1.Cells(i, 18) & Sheet1.Cells(i, 19)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AF"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 32)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AG"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 33)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AH"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 34)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AI"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 35)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AK"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 37)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AL"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 38)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AO"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 41)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AP"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 42)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AQ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 43)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AW"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 49)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BE"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 57)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BF"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 58)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BK"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 62)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BL"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 64)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With

    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "A"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 1)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "F"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 6)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "U"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 21)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function fwxxdc(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\10.房屋基本信息调查表.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "I-M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10) & Sheet1.Cells(i, 11) & Sheet1.Cells(i, 12) & Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "V-X"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 22) & "年" & Sheet1.Cells(i, 23) & "月" & Sheet1.Cells(i, 24) & "日"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AT"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 46)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AU"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 47)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AV"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 48)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AX"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 50)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
     With oword.Selection.Find
        .ClearFormatting
        .Text = "AW"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 49)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AP"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 42)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AM"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 39)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AN"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 40)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AH"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 34)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AI"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 35)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＳ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 45)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＣ－ＡＥ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 29) & Sheet1.Cells(i, 30) & Sheet1.Cells(i, 31)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BM"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 65)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BN"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 66)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BO"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 67)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BP"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 68)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BQ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 69)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＲ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 70)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＳ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 71)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "B"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "F"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 6)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "T"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 20)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function fwaqcls(name As String, num As String, i As Integer)

    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\13.房屋安全承诺书.doc")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "V-X"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 22) & "年" & Sheet1.Cells(i, 23) & "月" & Sheet1.Cells(i, 24) & "日"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＡ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 53)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＱ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 69)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｃ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｄ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 4)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｅ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｊ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｋ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｌ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｍ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｔ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 20)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function sjqr(name As String, num As String, i As Integer)

    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\14.农村宅基地及房屋三级确认审核表.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "G内、+1"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 77)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AF"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 32)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AI"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 35)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AG"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 33)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AJ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 36)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BQ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 69)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "K"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "J"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "L"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "V"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 22)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "W"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 23)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function bdcdj(name As String, num As String, i As Integer)

    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\15.不动产登记申请审批表.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "I-M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10) & Sheet1.Cells(i, 11) & Sheet1.Cells(i, 12) & Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｎ－Ｓ＋ＡＺ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 14) & Sheet1.Cells(i, 15) & Sheet1.Cells(i, 16) & Sheet1.Cells(i, 17) & Sheet1.Cells(i, 18) & Sheet1.Cells(i, 19) & Sheet1.Cells(i, 52)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AP"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 42)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AG"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 33)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AW"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 33)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AI"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 33)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "AJ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 36)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BL+/BO"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 64) & "+/" & Sheet1.Cells(i, 67)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
        
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "F"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 6)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "U"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 21)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function zjd(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\16.宅基地权属来源证明（无证，审批材料使用）.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＡ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＧ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＪ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "BQ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "G内、+1"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 77)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "C"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "E"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "K"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "L"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "V"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 22) & "年"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "W"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 23) & "月"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｗ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 23) & "月"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｘ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 24) & "日"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｙ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 25)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｚ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "T"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 20)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Public Function bdccl(name As String, num As String, i As Integer)
    
    Dim oword As Word.Application
    On Error Resume Next
    Set oword = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set oword = CreateObject("Word.Application")
    End If
    
    Dim doc As Word.Document
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\18.不动产测量报告（以村为单位出报告）.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＵ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 73)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＡ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＳ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 45)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＦ"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 32)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＡＩ"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 35)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＶ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 74)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＷ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 75)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "ＢＸ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 76)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "I-M"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 9) & Sheet1.Cells(i, 10) & Sheet1.Cells(i, 11) & Sheet1.Cells(i, 12) & Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｙ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 25)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｚ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "Ｍ"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Sub wordtq()
    
    path = "C:\Users\PWB\Desktop\chengguo\" '源文件路径
    source = "C:\Users\PWB\Desktop\米易房地一体模板-0422\" '目标路径
    
    If Dir(path, vbDirectory) = "" Then
        MkDir (path)
    End If

    Dim name As String
    Dim num As String
    Dim i As Integer
    
    For i = 6 To Sheet1.UsedRange.Rows.Count
    
        name = Sheet1.Cells(i, 3).Value
        num = Sheet1.Cells(i, 2).Value
        If Not (name = "" Or name = " ") Then
        
            If Dir(path & name & "(" & num & ")", vbDirectory) = "" Then
                MkDir (path & name & "(" & num & ")")
            End If
            
            FileCopy source & "00.档案袋.docx", path & name & "(" & num & ")\00.档案袋.docx"
            FileCopy source & "01.指界通知书.docx", path & name & "(" & num & ")\01.指界通知书.docx"
            FileCopy source & "02.指界委托书.docx", path & name & "(" & num & ")\02.指界委托书.docx"
            FileCopy source & "03.户主声明书.docx", path & name & "(" & num & ")\03.户主声明书.docx"
            FileCopy source & "05.农村宅基地使用权及房屋所有权确权登记申请表.docx", path & name & "(" & num & ")\05.农村宅基地使用权及房屋所有权确权登记申请表.docx"
            FileCopy source & "06.不动产权籍调查表.docx", path & name & "(" & num & ")\06.不动产权籍调查表.docx"
            FileCopy source & "10.房屋基本信息调查表.docx", path & name & "(" & num & ")\10.房屋基本信息调查表.docx"
            FileCopy source & "12.房屋照片.docx", path & name & "(" & num & ")\12.房屋照片.docx"
            FileCopy source & "13.房屋安全承诺书.doc", path & name & "(" & num & ")\13.房屋安全承诺书.doc"
            FileCopy source & "14.农村宅基地及房屋三级确认审核表.docx", path & name & "(" & num & ")\14.农村宅基地及房屋三级确认审核表.docx"
            FileCopy source & "15.不动产登记申请审批表.docx", path & name & "(" & num & ")\15.不动产登记申请审批表.docx"
            FileCopy source & "16.宅基地权属来源证明（无证，审批材料使用）.docx", path & name & "(" & num & ")\16.宅基地权属来源证明（无证，审批材料使用）.docx"
            FileCopy source & "17.违约缺席定界通知书（人不在本地的使用）.docx", path & name & "(" & num & ")\17.违约缺席定界通知书（人不在本地的使用）.docx"
            FileCopy source & "18.不动产测量报告（以村为单位出报告）.docx", path & name & "(" & num & ")\18.不动产测量报告（以村为单位出报告）.docx"
            
            dad name, num, i
            zjtzs name, num, i
            zjwts name, num, i
            hzsms name, num, i
            qqdj name, num, i
            bdcqdc name, num, i
            fwxxdc name, num, i
            fwaqcls name, num, i
            sjqr name, num, i
            bdcdj name, num, i
            zjd name, num, i
            bdccl name, num, i
            fwzp name, num, i
            
        End If
        
    Next

End Sub
