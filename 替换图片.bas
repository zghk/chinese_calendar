Sub �滻���ж���ΪͼƬ()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim picturePath As String
    
    ' ����ͼƬ·�� - ���޸�Ϊ���ʵ��ͼƬ·��
    picturePath = "D:\Green\�ű�\����\��.svg"
    
    ' ȷ��ͼƬ�ļ�����
    If Dir(picturePath) = "" Then
        MsgBox "�Ҳ���ָ����ͼƬ�ļ���", vbExclamation
        Exit Sub
    End If
    
    ' ������ǰ�����������й�����
    For Each ws In ThisWorkbook.Worksheets
        ' ���������������״����
        If ws.Shapes.Count > 0 Then
            ' �Ӻ���ǰ����������״������ɾ������ʱ����Ӱ��������
            For i = ws.Shapes.Count To 1 Step -1
                ' ��ȡ��ǰ��״����
                Set shp = ws.Shapes(i)
                
                ' ��¼ԭʼλ�úʹ�С
                Dim Left As Double: Left = shp.Left
                Dim Top As Double: Top = shp.Top
                Dim Width As Double: Width = shp.Width
                Dim Height As Double: Height = shp.Height
                
                ' ɾ��ԭʼ����
                shp.Delete
                
                ' ������ͼƬ������λ�úʹ�С
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
    
    MsgBox "���ж������滻��ɣ�", vbInformation
End Sub

