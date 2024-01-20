Attribute VB_Name = "���Ÿ���ɫ��1"
Sub InsertPictures���Ÿ���ɫ��1()
    '�������
    
    '�����С���ɫ�С�ͼƬ�����С�ͼƬ·����ͼƬ����
    Dim No_name, Color_name, Piccolumn, Pic_dir, Product_name As String

    Dim picPath As String
    
    '��Ԫ���ȡ��߶�
    Dim Pic_ColWidth, Pic_RowHeight As Integer
    
    'ͼƬ�Ŀ�ȡ��߶�
    Dim picWidth, picHeight As Integer
    
    '��ʼ����i���������һ���к�
    Dim i, MaxRowCount As Integer
    
    '��Ԫ�������š���ɫ
    Dim No_Range, Color_Range As Range
    
    Dim prevNoRange As Variant
    Dim picShapeRange As ShapeRange
    

    '�ݴ���
    On Error Resume Next

    '�ر���Ļ���£������ٶ�
    'Application.ScreenUpdating = False
    
    '���û���������
    No_name = InputBox("��������������е�����", "����������:", "A")
    '���Ϊ�����˳�
    If No_name = "" Then Exit Sub

    '������ɫ������
    Color_name = InputBox("��������ɫ�����е�����", "��ɫ������:", "C")
    '���Ϊ�����˳�
    If Color_name = "" Then Exit Sub
    
    
    '����ͼƬ������
    Piccolumn = InputBox("������ͼƬ���������е����ơ�", "ͼƬ����������", "E")
    '���Ϊ�����˳�
    If Piccolumn = "" Then Exit Sub
    
    '����ͼƬ�ļ���·��
    Pic_dir = InputBox("������ͼƬ�ļ���ŵ��ļ���·����", "ͼƬ�ļ���·��", "D:\User\Desktop\�½��ļ��� (2)\����-��ɫ")
    '���Ϊ�����˳�
    If Pic_dir = "" Then Exit Sub

    '��ͼƬ�ļ���·�����û��б��\�������
    If Right(Pic_dir, 1) <> "\" Then Pic_dir = Pic_dir & "\"
    
    '���õ�Ԫ��Ŀ��
    Pic_ColWidth = 10
    Pic_RowHeight = 60
    
    '��ȡ������������һ���к�
    MaxRowCount = Cells(Rows.Count, No_name).End(xlUp).Row
    'MsgBox MaxRowCount
    
    '�����п�
    Columns(Piccolumn).ColumnWidth = Pic_ColWidth
    'MsgBox "�п�" & Pic_ColWidth
    
    '�����и�
    Rows("2:" & MaxRowCount).RowHeight = Pic_RowHeight
    'MsgBox "�и�" & Pic_RowHeight
    
    '����2�����ÿ�ʼ���ͼƬ���к��ǵ�2��
    For i = 2 To MaxRowCount
                
        ' ��ʼ�� prevNoRange Ϊ��һ�� No_Range ��ֵ��������ڣ�
        If Not IsEmpty(Cells(i, No_name).Value) Then
            prevNoRange = Cells(i, No_name).Value
        End If
        
        ' ��ȡ No_Range �� Color_Range ��ֵ
        Set No_Range = Cells(i, No_name)
        Set Color_Range = Cells(i, Color_name)
                
        ' ��� No_Range �Ƿ�Ϊ�գ���ʹ����һ��ֵ
        If No_Range.Value = "" Then
            No_Range.Value = prevNoRange
        Else
            prevNoRange = No_Range.Value
        End If
        'MsgBox prevNoRange & Color_Range
        
        '��ȡͼƬ�ļ������ȶ�ȡjpg��ʽ����û�У����ȡjpeg��ʽ������Ȼû�У�����ڶ�ȡpng��ʽ
        picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".jpg"
        'MsgBox picPath

        If Dir(picPath) = "" Then
            '��ȡjpeg��ʽͼƬ
            picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".jpeg"
            '��ȡpng��ʽͼƬ
            If Dir(picPath) = "" Then picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".png"
        End If
        
        '����ļ��Ƿ����
        If Dir(picPath) <> "" Then
            'MsgBox "�ļ�����" & picPath
            
            '��ȡ����ͼƬ�ĵ�Ԫ�����
            Set No_Range = Cells(i, Piccolumn)
            
            'ѡ�е�Ԫ��
            No_Range.Select

            '����һ�����Եȱ�����
            ActiveSheet.Pictures.Insert(picPath).Select
            Set picShapeRange = Selection.ShapeRange
            
            '��ȡͼƬ���
            picWidth = picShapeRange.Width
            picHeight = picShapeRange.Height
            'MsgBox "��" & picWidth & "��" & picHeight
            
            '����ͼƬ�Ŀ�ߣ���ͼƬ���з��ã���˻���Ҫ����ͼƬ�ı߾�
            If picWidth >= picHeight Then
                picShapeRange.Width = No_Range.Width - 2
                picShapeRange.Left = picShapeRange.Left + 1
                picShapeRange.Top = No_Range.Top + (No_Range.Height - picShapeRange.Height) / 2
            Else
                picShapeRange.Height = No_Range.Height - 2
                picShapeRange.Top = picShapeRange.Top + 1
                picShapeRange.Left = No_Range.Left + (No_Range.Width - picShapeRange.Width) / 2
            End If
            
            '������
            'Set pic = ActiveSheet.Shapes.AddPicture(picPath, False, True, No_Range.Left, No_Range.Top, -2, -2)
            'pic.Height = No_Range.Height
            'pic.Width = No_Range.Width  '(No_Range.Width - pic.Width) / 2 + No_Range.Left
            
            '��������������ȫ��䵥Ԫ��
            'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (No_Range.Left + 1), (No_Range.Top + 1), (No_Range.Width - 1), (No_Range.Height - 1)).Fill.UserPicture picPath
            
        End If
           
    Next i
    ActiveSheet.Shapes.SelectAll
    
    '���þ��ζ����ޱ߿�
    Selection.ShapeRange.Line.Visible = msoFalse
    
    Application.ScreenUpdating = True

End Sub
