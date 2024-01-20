Attribute VB_Name = "货号1"
Sub InsertPicture货号1()
        'copy right 2020 by billy
    '定义变量
    Dim cellcolumn, Piccolumn As String
    
    Dim picDir, picPath As String
    
    Dim i, MaxRowCount As Integer
    
    Dim picColWidth, picRowHeight As Integer
    
    Dim picWidth, picHeight As Integer
    
    Dim SrcRange, picRange As Range
    
    Dim picShapeRange As ShapeRange
    
    
    '容错处理
    On Error Resume Next
    
    '关闭屏幕更新，提升速度
    'Application.ScreenUpdating = False
    
    '设置款号所在列
    cellcolumn = InputBox("请输入款号所在列的名称（图片的文件名在哪一列）。", "款号列名称", "A")

    '设置插入图片所在第几列
    Piccolumn = InputBox("请输入图片插入后所在列的名称（图片插入后放在哪一列）。", "图片列名称", "F")

    '图片存放的文件夹路径,如：E:\FX_Image\
    picDir = InputBox("请输入图片文件存放的文件夹路径。", "图片路径", "D:\User\Desktop\新建文件夹 (2)\货号")

    '输入有误是，则退出
    If cellcolumn = "" Or Piccolumn = "" Or picDir = "" Then Exit Sub
    
    '若图片路径文件夹最后没有斜杠\，则加上
    If Right(picDir, 1) <> "\" Then picDir = picDir & "\"
    
    '图片单元格的宽高
    picColWidth = 10
    picRowHeight = 60

    '获取数据区域的最后一行行号
    MaxRowCount = Cells(Rows.Count, cellcolumn).End(xlUp).Row

    '设置列宽
    Columns(Piccolumn).ColumnWidth = picColWidth
        'MsgBox "列宽" & picColWidth
    
    '设置行高
    Rows("2:" & MaxRowCount).RowHeight = picRowHeight
    
    '数字2是设置开始填充图片的行号是第2行
    For i = 2 To MaxRowCount
        
        '图片文件名所在的单元格对象
        Set SrcRange = Cells(i, cellcolumn)
        
        '读取图片文件，优先读取jpg格式，若没有，则读取jpeg格式，若仍然没有，最后在读取png格式
        picPath = picDir & SrcRange & ".jpg"
        '检查文件是否存在
        If Dir(picPath) = "" Then
            '获取jpeg格式图片
            picPath = picDir & SrcRange & ".jpeg"
            '获取png格式图片
            If Dir(picPath) = "" Then picPath = picDir & SrcRange & ".png"
        End If
        
        If Dir(picPath) <> "" Then
        
            '获取放置图片的单元格对象
            Set picRange = Cells(i, Piccolumn)
            
            '选中单元格
            picRange.Select
            

            
            '方法一，可以等比缩放
            ActiveSheet.Pictures.Insert(picPath).Select
            Set picShapeRange = Selection.ShapeRange
            
            '等比缩放
            picShapeRange.LockAspectRatio = msoTrue
            
            '获取图片宽高
            picWidth = picShapeRange.Width
            picHeight = picShapeRange.Height
            
            '设置图片的宽高，将图片居中放置，因此还需要计算图片的边距
            If picWidth >= picHeight Then
                picShapeRange.Width = picRange.Width - 1
                picShapeRange.Left = picShapeRange.Left + 1
                picShapeRange.Top = picRange.Top + (picRange.Height - picShapeRange.Height) / 2
            Else
                picShapeRange.Height = picRange.Height - 1
                picShapeRange.Top = picShapeRange.Top + 1
                picShapeRange.Left = picRange.Left + (picRange.Width - picShapeRange.Width) / 2
            End If
            
            '设置图片属性为：大小和位置随单元格而变， xlMoveAndSize:大小和位置随单元格而变,xlMove:大小固定，位置随单元格而变,xlFreeFloating:大小和位置固定
            Selection.Placement = xlMoveAndSize
            
            '方法二
            'Set pic = ActiveSheet.Shapes.AddPicture(picPath, False, True, picRange.Left, picRange.Top, -1, -1)
            'pic.Height = picRange.Height
            'pic.Width = picRange.Width  '(picRange.Width - pic.Width) / 2 + picRange.Left
            
            '方法三，可以完全填充单元格
            'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (picRange.Left + 1), (picRange.Top + 1), (picRange.Width - 1), (picRange.Height - 1)).Fill.UserPicture picPath

        End If
            
    Next i

    ActiveSheet.Shapes.SelectAll
    
    '设置矩形对象无边框
    'Selection.ShapeRange.Line.Visible = msoFalse
    
    'Application.ScreenUpdating = True
    
    Range("A1").Select

End Sub
