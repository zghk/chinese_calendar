Sub 替换所有对象为图片()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim picturePath As String
    
    ' 设置图片路径 - 请修改为你的实际图片路径
    picturePath = "D:\Green\脚本\日历\休.svg"
    
    ' 确认图片文件存在
    If Dir(picturePath) = "" Then
        MsgBox "找不到指定的图片文件！", vbExclamation
        Exit Sub
    End If
    
    ' 遍历当前工作簿的所有工作表
    For Each ws In ThisWorkbook.Worksheets
        ' 如果工作表中有形状对象
        If ws.Shapes.Count > 0 Then
            ' 从后向前遍历所有形状（这样删除对象时不会影响索引）
            For i = ws.Shapes.Count To 1 Step -1
                ' 获取当前形状对象
                Set shp = ws.Shapes(i)
                
                ' 记录原始位置和大小
                Dim Left As Double: Left = shp.Left
                Dim Top As Double: Top = shp.Top
                Dim Width As Double: Width = shp.Width
                Dim Height As Double: Height = shp.Height
                
                ' 删除原始对象
                shp.Delete
                
                ' 插入新图片并设置位置和大小
                ws.Shapes.AddPicture _
                    Filename:=picturePath, _
                    LinkToFile:=False, _
                    SaveWithDocument:=True, _
                    Left:=Left, _
                    Top:=Top, _
                    Width:=Width, _
                    Height:=Height
            Next i
        End If
    Next ws
    
    MsgBox "所有对象已替换完成！", vbInformation
End Sub

