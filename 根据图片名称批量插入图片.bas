Attribute VB_Name = "����1"
Sub InsertPicture����1()
        'copy right 2020 by billy
    '�������
    Dim cellcolumn, Piccolumn As String
    
    Dim picDir, picPath As String
    
    Dim i, MaxRowCount As Integer
    
    Dim picColWidth, picRowHeight As Integer
    
    Dim picWidth, picHeight As Integer
    
    Dim SrcRange, picRange As Range
    
    Dim picShapeRange As ShapeRange
    
    
    '�ݴ���
    On Error Resume Next
    
    '�ر���Ļ���£������ٶ�
    'Application.ScreenUpdating = False
    
    '���ÿ��������
    cellcolumn = InputBox("�������������е����ƣ�ͼƬ���ļ�������һ�У���", "���������", "A")

    '���ò���ͼƬ���ڵڼ���
    Piccolumn = InputBox("������ͼƬ����������е����ƣ�ͼƬ����������һ�У���", "ͼƬ������", "F")

    'ͼƬ��ŵ��ļ���·��,�磺E:\FX_Image\
    picDir = InputBox("������ͼƬ�ļ���ŵ��ļ���·����", "ͼƬ·��", "D:\User\Desktop\�½��ļ��� (2)\����")

    '���������ǣ����˳�
    If cellcolumn = "" Or Piccolumn = "" Or picDir = "" Then Exit Sub
    
    '��ͼƬ·���ļ������û��б��\�������
    If Right(picDir, 1) <> "\" Then picDir = picDir & "\"
    
    'ͼƬ��Ԫ��Ŀ��
    picColWidth = 10
    picRowHeight = 60

    '��ȡ������������һ���к�
    MaxRowCount = Cells(Rows.Count, cellcolumn).End(xlUp).Row

    '�����п�
    Columns(Piccolumn).ColumnWidth = picColWidth
        'MsgBox "�п�" & picColWidth
    
    '�����и�
    Rows("2:" & MaxRowCount).RowHeight = picRowHeight
    
    '����2�����ÿ�ʼ���ͼƬ���к��ǵ�2��
    For i = 2 To MaxRowCount
        
        'ͼƬ�ļ������ڵĵ�Ԫ�����
        Set SrcRange = Cells(i, cellcolumn)
        
        '��ȡͼƬ�ļ������ȶ�ȡjpg��ʽ����û�У����ȡjpeg��ʽ������Ȼû�У�����ڶ�ȡpng��ʽ
        picPath = picDir & SrcRange & ".jpg"
        '����ļ��Ƿ����
        If Dir(picPath) = "" Then
            '��ȡjpeg��ʽͼƬ
            picPath = picDir & SrcRange & ".jpeg"
            '��ȡpng��ʽͼƬ
            If Dir(picPath) = "" Then picPath = picDir & SrcRange & ".png"
        End If
        
        If Dir(picPath) <> "" Then
        
            '��ȡ����ͼƬ�ĵ�Ԫ�����
            Set picRange = Cells(i, Piccolumn)
            
            'ѡ�е�Ԫ��
            picRange.Select
            

            
            '����һ�����Եȱ�����
            ActiveSheet.Pictures.Insert(picPath).Select
            Set picShapeRange = Selection.ShapeRange
            
            '�ȱ�����
            picShapeRange.LockAspectRatio = msoTrue
            
            '��ȡͼƬ���
            picWidth = picShapeRange.Width
            picHeight = picShapeRange.Height
            
            '����ͼƬ�Ŀ�ߣ���ͼƬ���з��ã���˻���Ҫ����ͼƬ�ı߾�
            If picWidth >= picHeight Then
                picShapeRange.Width = picRange.Width - 1
                picShapeRange.Left = picShapeRange.Left + 1
                picShapeRange.Top = picRange.Top + (picRange.Height - picShapeRange.Height) / 2
            Else
                picShapeRange.Height = picRange.Height - 1
                picShapeRange.Top = picShapeRange.Top + 1
                picShapeRange.Left = picRange.Left + (picRange.Width - picShapeRange.Width) / 2
            End If
            
            '����ͼƬ����Ϊ����С��λ���浥Ԫ����䣬 xlMoveAndSize:��С��λ���浥Ԫ�����,xlMove:��С�̶���λ���浥Ԫ�����,xlFreeFloating:��С��λ�ù̶�
            Selection.Placement = xlMoveAndSize
            
            '������
            'Set pic = ActiveSheet.Shapes.AddPicture(picPath, False, True, picRange.Left, picRange.Top, -1, -1)
            'pic.Height = picRange.Height
            'pic.Width = picRange.Width  '(picRange.Width - pic.Width) / 2 + picRange.Left
            
            '��������������ȫ��䵥Ԫ��
            'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (picRange.Left + 1), (picRange.Top + 1), (picRange.Width - 1), (picRange.Height - 1)).Fill.UserPicture picPath

        End If
            
    Next i

    ActiveSheet.Shapes.SelectAll
    
    '���þ��ζ����ޱ߿�
    'Selection.ShapeRange.Line.Visible = msoFalse
    
    'Application.ScreenUpdating = True
    
    Range("A1").Select

End Sub
