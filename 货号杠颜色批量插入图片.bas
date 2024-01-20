Attribute VB_Name = "货号杠颜色自1"
Sub InsertPictures货号杠颜色自1()
    '定义变量
    
    '货号列、颜色列、图片插入列、图片路径、图片名称
    Dim No_name, Color_name, Piccolumn, Pic_dir, Product_name As String

    Dim picPath As String
    
    '单元格宽度、高度
    Dim Pic_ColWidth, Pic_RowHeight As Integer
    
    '图片的宽度、高度
    Dim picWidth, picHeight As Integer
    
    '开始列数i、货号最后一行行号
    Dim i, MaxRowCount As Integer
    
    '单元格对象货号、颜色
    Dim No_Range, Color_Range As Range
    
    Dim prevNoRange As Variant
    Dim picShapeRange As ShapeRange
    

    '容错处理
    On Error Resume Next

    '关闭屏幕更新，提升速度
    'Application.ScreenUpdating = False
    
    '设置货号所在列
    No_name = InputBox("请输入货号所在列的名称", "货号列名称:", "A")
    '如果为空则退出
    If No_name = "" Then Exit Sub

    '设置颜色所在列
    Color_name = InputBox("请输入颜色所在列的名称", "颜色列名称:", "C")
    '如果为空则退出
    If Color_name = "" Then Exit Sub
    
    
    '设置图片插入列
    Piccolumn = InputBox("请输入图片插入所在列的名称。", "图片插入列名称", "E")
    '如果为空则退出
    If Piccolumn = "" Then Exit Sub
    
    '设置图片文件夹路径
    Pic_dir = InputBox("请输入图片文件存放的文件夹路径。", "图片文件夹路径", "D:\User\Desktop\新建文件夹 (2)\货号-颜色")
    '如果为空则退出
    If Pic_dir = "" Then Exit Sub

    '若图片文件夹路径最后没有斜杠\，则加上
    If Right(Pic_dir, 1) <> "\" Then Pic_dir = Pic_dir & "\"
    
    '设置单元格的宽高
    Pic_ColWidth = 10
    Pic_RowHeight = 60
    
    '获取货号区域的最后一行行号
    MaxRowCount = Cells(Rows.Count, No_name).End(xlUp).Row
    'MsgBox MaxRowCount
    
    '设置列宽
    Columns(Piccolumn).ColumnWidth = Pic_ColWidth
    'MsgBox "列宽" & Pic_ColWidth
    
    '设置行高
    Rows("2:" & MaxRowCount).RowHeight = Pic_RowHeight
    'MsgBox "行高" & Pic_RowHeight
    
    '数字2是设置开始填充图片的行号是第2行
    For i = 2 To MaxRowCount
                
        ' 初始化 prevNoRange 为第一个 No_Range 的值（如果存在）
        If Not IsEmpty(Cells(i, No_name).Value) Then
            prevNoRange = Cells(i, No_name).Value
        End If
        
        ' 获取 No_Range 和 Color_Range 的值
        Set No_Range = Cells(i, No_name)
        Set Color_Range = Cells(i, Color_name)
                
        ' 检查 No_Range 是否为空，并使用上一个值
        If No_Range.Value = "" Then
            No_Range.Value = prevNoRange
        Else
            prevNoRange = No_Range.Value
        End If
        'MsgBox prevNoRange & Color_Range
        
        '读取图片文件，优先读取jpg格式，若没有，则读取jpeg格式，若仍然没有，最后在读取png格式
        picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".jpg"
        'MsgBox picPath

        If Dir(picPath) = "" Then
            '获取jpeg格式图片
            picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".jpeg"
            '获取png格式图片
            If Dir(picPath) = "" Then picPath = Pic_dir & prevNoRange & "-" & Color_Range & ".png"
        End If
        
        '检查文件是否存在
        If Dir(picPath) <> "" Then
            'MsgBox "文件存在" & picPath
            
            '获取放置图片的单元格对象
            Set No_Range = Cells(i, Piccolumn)
            
            '选中单元格
            No_Range.Select

            '方法一，可以等比缩放
            ActiveSheet.Pictures.Insert(picPath).Select
            Set picShapeRange = Selection.ShapeRange
            
            '获取图片宽高
            picWidth = picShapeRange.Width
            picHeight = picShapeRange.Height
            'MsgBox "宽" & picWidth & "高" & picHeight
            
            '设置图片的宽高，将图片居中放置，因此还需要计算图片的边距
            If picWidth >= picHeight Then
                picShapeRange.Width = No_Range.Width - 2
                picShapeRange.Left = picShapeRange.Left + 1
                picShapeRange.Top = No_Range.Top + (No_Range.Height - picShapeRange.Height) / 2
            Else
                picShapeRange.Height = No_Range.Height - 2
                picShapeRange.Top = picShapeRange.Top + 1
                picShapeRange.Left = No_Range.Left + (No_Range.Width - picShapeRange.Width) / 2
            End If
            
            '方法二
            'Set pic = ActiveSheet.Shapes.AddPicture(picPath, False, True, No_Range.Left, No_Range.Top, -2, -2)
            'pic.Height = No_Range.Height
            'pic.Width = No_Range.Width  '(No_Range.Width - pic.Width) / 2 + No_Range.Left
            
            '方法三，可以完全填充单元格
            'ActiveSheet.Shapes.AddShape(msoShapeRectangle, (No_Range.Left + 1), (No_Range.Top + 1), (No_Range.Width - 1), (No_Range.Height - 1)).Fill.UserPicture picPath
            
        End If
           
    Next i
    ActiveSheet.Shapes.SelectAll
    
    '设置矩形对象无边框
    Selection.ShapeRange.Line.Visible = msoFalse
    
    Application.ScreenUpdating = True

End Sub
