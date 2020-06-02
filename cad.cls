Public Sub capture()

Dim group As AcadGroup
Dim flag As Boolean
Dim entity As AcadEntity
Dim text As AcadText
Dim max_x As Double
Dim min_x As Double
Dim max_y As Double
Dim min_y As Double
Dim polyline As AcadPolyline
Dim line As AcadLine
Dim cor As Variant
Dim ss As AcadSelectionSet
'ReDim gpCode(0) As Integer
'ReDim dataValue(0) As Variant
'Dim groupCode As Variant, dataCode As Variant
Dim pointsArray(0 To 23) As Double
Dim l As Double
Dim h As Double
Dim a As Double
Dim c_x As Double
Dim c_y As Double
Dim name As String
Dim path As String
Dim leng As Integer

path = "C:\Users\PWB\Desktop\cad\" '保存路径

a = 1.3 '放大系数

'gpCode(0) = 8
'dataValue(0) = "JDZ"

'groupCode = gpCode
'dataCode = dataValue

For Each group In ThisDrawing.Groups
    
    flag = False
    '查找带有判断中文名字group
    For Each entity In group
        If entity.ObjectName = "AcDbText" Then
            Set text = entity
            name = text.TextString
            '中文字符且长度大于1
            If Len(name) > 1 And ((AscW(Mid(name, 1, 1)) > -40870 And AscW(Mid(name, 1, 1)) < -19967) Or (AscW(Mid(name, 1, 1)) < 40870 And AscW(Mid(name, 1, 1)) > 19967)) Then
                flag = True
                Exit For
            End If
        End If
    Next
    
    '找到并处理
    If flag = True Then
    
        max_x = 0
        max_y = 0
        min_x = 99999999
        min_y = 99999999
        
        '选择区域
        For Each entity In group
            If entity.ObjectName = "AcDbPolyline" Then
                cor = entity.Coordinate(0)
                If max_x < cor(0) Then
                    max_x = cor(0)
                End If
                If max_y < cor(1) Then
                    max_y = cor(1)
                End If
                If min_x > cor(0) Then
                    min_x = cor(0)
                End If
                If min_y > cor(1) Then
                    min_y = cor(1)
                End If
                cor = entity.Coordinate(1)
                If max_x < cor(0) Then
                    max_x = cor(0)
                End If
                If max_y < cor(1) Then
                    max_y = cor(1)
                End If
                If min_x > cor(0) Then
                    min_x = cor(0)
                End If
                If min_y > cor(1) Then
                    min_y = cor(1)
                End If
            End If
            If entity.ObjectName = "AcDbLine" Then
                cor = entity.StartPoint
                If max_x < cor(0) Then
                    max_x = cor(0)
                End If
                If max_y < cor(1) Then
                    max_y = cor(1)
                End If
                If min_x > cor(0) Then
                    min_x = cor(0)
                End If
                If min_y > cor(1) Then
                    min_y = cor(1)
                End If
                cor = entity.EndPoint
                If max_x < cor(0) Then
                    max_x = cor(0)
                End If
                If max_y < cor(1) Then
                    max_y = cor(1)
                End If
                If min_x > cor(0) Then
                    min_x = cor(0)
                End If
                If min_y > cor(1) Then
                    min_y = cor(1)
                End If
            End If
        Next
        
        '选择集
        Set ss = ThisDrawing.SelectionSets.Add("ss")
        '新建保存路径
        If Dir(path & name, vbDirectory) = "" Then
            MkDir (path & name)
        End If
    
        c_x = (min_x + max_x) / 2
        c_y = (min_y + max_y) / 2
        l = (max_x - min_x) / 2 * a '放大指定倍数
        h = (max_y - min_y) / 2 * a
        
        pointsArray(0) = c_x - l: pointsArray(1) = c_y + h: pointsArray(2) = -99999999
        pointsArray(3) = c_x + l: pointsArray(4) = c_y + h: pointsArray(5) = -99999999
        pointsArray(6) = c_x - l: pointsArray(7) = c_y - h: pointsArray(8) = -99999999
        pointsArray(9) = c_x + l: pointsArray(10) = c_y - h: pointsArray(11) = -99999999

        pointsArray(12) = c_x - l: pointsArray(13) = c_y + h: pointsArray(14) = 99999999
        pointsArray(15) = c_x + l: pointsArray(16) = c_y + h: pointsArray(17) = 99999999
        pointsArray(18) = c_x - l: pointsArray(19) = c_y - h: pointsArray(20) = 99999999
        pointsArray(21) = c_x + l: pointsArray(22) = c_y - h: pointsArray(23) = 99999999
        
        ss.SelectByPolygon acSelectionSetFence, pointsArray
        
        c = ss.Count - 1
        
        ReDim ocollection(0 To c) As AcadEntity
        
        For Index = 0 To c
        Set ocollection(Index) = ss.Item(Index)
        Next
        
        '创建文档
        Dim ACADapp As AcadApplication
        
        Set ACADapp = GetObject(, "AutoCAD.Application")
        
        Dim curdoc As AcadDocument
        
        Set curdoc = ACADapp.Documents(ACADapp.ActiveDocument.name)
        
        Dim newdoc As AcadDocument
        Set newdoc = ACADapp.Documents.Add
        '复制选择区域
        curdoc.CopyObjects ocollection, newdoc.ModelSpace
        
        ss.Delete
        '保存
        ACADapp.ActiveDocument.SaveAs (path & name & "\宗地草图.dwg")
        
        ACADapp.ActiveDocument.Close
        
    End If
    
Next

End Sub
