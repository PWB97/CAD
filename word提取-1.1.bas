Attribute VB_Name = "ģ��11"
'word��ȡ 1.1
'�޸�name�����ڱ�������⼰�������⣻���ӹ���Ȩ�������֤����Ϣ���룻������λС�� 20.8.13
'����������Ƭ 20.8.11

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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\00.������.docx")
    
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\01.ָ��֪ͨ��.docx")
    
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
            .Text = "AB-12�����ж�"
            .Replacement.ClearFormatting
            .Replacement.Text = "��"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    Else
        With oword.Selection.Find
            .ClearFormatting
            .Text = "AB-12�����ж�"
            .Replacement.ClearFormatting
            .Replacement.Text = "��"
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
        .Text = "���̱���/����λ���������˻����� /����ũ�����ƾٵ�ָ���ˣ�����"
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
        .Text = "��"
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
        .Text = "��"
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
        .Text = "��"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\02.ָ��ί����.docx")
    
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\12.������Ƭ.docx")
    
    Dim table As table
    Set table = doc.Tables(1)
    
    MyPath = path & name & "(" & num & ")\������\"   ' ָ��·�� ���������\
    myname = Dir(MyPath, vbDirectory)   ' ��Ѱ��һ��
    Dim j As Integer
    j = 1
    Do While myname <> ""   ' ��ʼѭ��
        ' ������ǰ��Ŀ¼���ϲ�Ŀ¼
        If myname <> "." And myname <> ".." Then
            ' ʹ��λ�Ƚ���ȷ�� MyName����һĿ¼
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
        myname = Dir    ' ������һ��Ŀ¼
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\03.����������.docx")
    
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\05.ũ��լ����ʹ��Ȩ����������ȨȷȨ�Ǽ������.docx")
    
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
        .Text = "G�������+1 " '������ԭ�ĵ��滻�ַ��������޸��� todo
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
        arr = Split(gyqlr, "��")
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\06.������Ȩ�������.docx")
    
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\10.���ݻ�����Ϣ�����.docx")
    
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
        .Replacement.Text = Sheet1.Cells(i, 22) & "��" & Sheet1.Cells(i, 23) & "��" & Sheet1.Cells(i, 24) & "��"
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
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 45)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "���ã�����"
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
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 70)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\13.���ݰ�ȫ��ŵ��.doc")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "V-X"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 22) & "��" & Sheet1.Cells(i, 23) & "��" & Sheet1.Cells(i, 24) & "��"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 53)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 69)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 3)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 4)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 5)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 10)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 11)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 12)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 13)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\14.ũ��լ���ؼ���������ȷ����˱�.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "G�ڡ�+1"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\15.�������Ǽ�����������.docx")
    
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
        .Text = "�Σ��ӣ�����"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\16.լ����Ȩ����Դ֤������֤����������ʹ�ã�.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
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
        .Text = "G�ڡ�+1"
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
        .Replacement.Text = Sheet1.Cells(i, 22) & "��"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "W"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 23) & "��"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 23) & "��"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 24) & "��"
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 25)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
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
    Set doc = oword.Documents.Open(path & name & "(" & num & ")\18.�������������棨�Դ�Ϊ��λ�����棩.docx")
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 73)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 27)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 45)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 32)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "����"
        .Replacement.ClearFormatting
        .Replacement.Text = Round(Val(Sheet1.Cells(i, 35)), 2)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 74)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 75)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "�£�"
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
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 25)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    With oword.Selection.Find
        .ClearFormatting
        .Text = "��"
        .Replacement.ClearFormatting
        .Replacement.Text = Sheet1.Cells(i, 26)
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    
    doc.Save
    doc.Close
    
End Function

Sub wordtq()
    
    path = "C:\Users\PWB\Desktop\chengguo\" 'Դ�ļ�·��
    source = "C:\Users\PWB\Desktop\���׷���һ��ģ��-0422\" 'Ŀ��·��
    
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
            
            FileCopy source & "00.������.docx", path & name & "(" & num & ")\00.������.docx"
            FileCopy source & "01.ָ��֪ͨ��.docx", path & name & "(" & num & ")\01.ָ��֪ͨ��.docx"
            FileCopy source & "02.ָ��ί����.docx", path & name & "(" & num & ")\02.ָ��ί����.docx"
            FileCopy source & "03.����������.docx", path & name & "(" & num & ")\03.����������.docx"
            FileCopy source & "05.ũ��լ����ʹ��Ȩ����������ȨȷȨ�Ǽ������.docx", path & name & "(" & num & ")\05.ũ��լ����ʹ��Ȩ����������ȨȷȨ�Ǽ������.docx"
            FileCopy source & "06.������Ȩ�������.docx", path & name & "(" & num & ")\06.������Ȩ�������.docx"
            FileCopy source & "10.���ݻ�����Ϣ�����.docx", path & name & "(" & num & ")\10.���ݻ�����Ϣ�����.docx"
            FileCopy source & "12.������Ƭ.docx", path & name & "(" & num & ")\12.������Ƭ.docx"
            FileCopy source & "13.���ݰ�ȫ��ŵ��.doc", path & name & "(" & num & ")\13.���ݰ�ȫ��ŵ��.doc"
            FileCopy source & "14.ũ��լ���ؼ���������ȷ����˱�.docx", path & name & "(" & num & ")\14.ũ��լ���ؼ���������ȷ����˱�.docx"
            FileCopy source & "15.�������Ǽ�����������.docx", path & name & "(" & num & ")\15.�������Ǽ�����������.docx"
            FileCopy source & "16.լ����Ȩ����Դ֤������֤����������ʹ�ã�.docx", path & name & "(" & num & ")\16.լ����Ȩ����Դ֤������֤����������ʹ�ã�.docx"
            FileCopy source & "17.ΥԼȱϯ����֪ͨ�飨�˲��ڱ��ص�ʹ�ã�.docx", path & name & "(" & num & ")\17.ΥԼȱϯ����֪ͨ�飨�˲��ڱ��ص�ʹ�ã�.docx"
            FileCopy source & "18.�������������棨�Դ�Ϊ��λ�����棩.docx", path & name & "(" & num & ")\18.�������������棨�Դ�Ϊ��λ�����棩.docx"
            
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
