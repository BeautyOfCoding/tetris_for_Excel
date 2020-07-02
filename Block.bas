Option Explicit

Dim MySheet As Worksheet
Dim iCenterRow As Integer   '方块中心行
Dim iCenterCol As Integer   '方块中心列
Dim ColorArr()              '7种颜色
Dim ShapeArr()              '7种方块
Dim iColorIndex As Integer  '颜色索引
Dim MyBlock(4, 2) As Integer    '每个方框的坐标数组，会随着方块的移动而变化
Dim bIsObjectEnd As Boolean     '本个方块是否下降到最低点
Dim iScore As Integer       '分数

'移动对象 By@yaxi_liu
Public Sub MoveObject(ByVal dir As Integer)
    Call MoveBlock(iCenterRow, iCenterCol, MyBlock, ColorArr(iColorIndex), dir)
End Sub
'旋转对象 By@yaxi_liu
Public Sub RotateObject()
    Call RotateBlock(iCenterRow, iCenterCol, MyBlock, ColorArr(iColorIndex))
End Sub

Sub Start()
    Call Init
    
'    iCenterRow = 5
'    iCenterCol = 6
'    iColorIndex = 4
'    Dim i As Integer
'    For i = 0 To 3
'        MyBlock(i, 0) = ShapeArr(iColorIndex)(i)(0)
'        MyBlock(i, 1) = ShapeArr(iColorIndex)(i)(1)
'    Next
'    Call DrawBlock(iCenterRow, iCenterCol, MyBlock, ColorArr(iColorIndex))
    
    While (True)
        Call GetBlock
        bIsObjectEnd = False    '本方块对象是否结束

        While (bIsObjectEnd = False)
            Call delay(0.5)
            Call MoveBlock(iCenterRow, iCenterCol, MyBlock, ColorArr(iColorIndex), 0)
            MySheet.Range("L21").Select
            With MySheet.Range("B1:K20")
                .Borders(xlEdgeBottom).Weight = xlMedium
                .Borders(xlEdgeRight).Weight = xlMedium
                .Borders(xlEdgeLeft).Weight = xlMedium
            End With
        Wend
        Call DeleteFullRow
    Wend
End Sub

Private Sub DeleteFullRow()
    Dim i As Integer, j As Integer
    For i = 1 To 20
        For j = 2 To 11
            If MySheet.Cells(i, j).Interior.ColorIndex < 0 Then
                Exit For
            ElseIf j = 11 Then
                MySheet.Range(Cells(1, 2), Cells(i - 1, j)).Cut Destination:=MySheet.Range(Cells(2, 2), Cells(i, j))       'Range("B2:K18")
                iScore = iScore + 10
            End If
        Next j
    Next i
    MySheet.Range("N1").Value = "分数"
    MySheet.Range("O1").Value = iScore
End Sub

Private Sub EndGame()
    
End Sub

Private Sub Init()
    Set MySheet = Sheets("Sheet1")
    ColorArr = Array(3, 4, 5, 6, 7, 8, 9)
    ShapeArr = Array(Array(Array(0, 0), Array(0, 1), Array(0, -1), Array(0, 2)), _
                 Array(Array(0, 0), Array(0, 1), Array(0, -1), Array(-1, -1)), _
                 Array(Array(0, 0), Array(0, 1), Array(0, -1), Array(-1, 1)), _
                 Array(Array(0, 0), Array(-1, 1), Array(-1, 0), Array(0, 1)), _
                 Array(Array(0, 0), Array(0, -1), Array(-1, 0), Array(-1, 1)), _
                 Array(Array(0, 0), Array(0, 1), Array(-1, 0), Array(-1, -1)), _
                 Array(Array(0, 0), Array(0, 1), Array(0, -1), Array(-1, 0)))
                 
    With MySheet.Range("B1:K20")
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    
    '设定长宽比例
    MySheet.Columns("A:L").ColumnWidth = 2
    MySheet.Rows("1:30").RowHeight = 13.5
    
    iScore = 0
    MySheet.Range("N1").Value = "分数"
    MySheet.Range("O1").Value = iScore
End Sub

'随机生成新的方块函数 By@yaxi_liu
Private Sub GetBlock()
    Randomize (Timer)
    Dim i As Integer
    iColorIndex = Int(7 * Rnd)
    iCenterRow = 2
    iCenterCol = 6
    
    For i = 0 To 3
        MyBlock(i, 0) = ShapeArr(iColorIndex)(i)(0)
        MyBlock(i, 1) = ShapeArr(iColorIndex)(i)(1)
    Next
    Call DrawBlock(iCenterRow, iCenterCol, MyBlock, ColorArr(iColorIndex))
End Sub
'绘制方块 By@yaxi_liu
Private Sub DrawBlock(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer, ByVal icolor As Integer)
    Dim Row As Integer, Col As Integer
    Dim i As Integer
    For i = 0 To 3
        Row = center_row + block(i, 0)
        Col = center_col + block(i, 1)
        MySheet.Cells(Row, Col).Interior.ColorIndex = icolor  '颜色索引
        MySheet.Cells(Row, Col).Borders.LineStyle = xlContinuous    '周围加外框线
    Next
End Sub

'擦除方块 By@yaxi_liu
Private Sub EraseBlock(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer)
    Dim Row As Integer, Col As Integer
    Dim i As Integer
    For i = 0 To 3
        Row = center_row + block(i, 0)
        Col = center_col + block(i, 1)
        MySheet.Cells(Row, Col).Interior.Pattern = xlNone
        MySheet.Cells(Row, Col).Borders.LineStyle = xlNone
    Next
End Sub
'移动方块 By@yaxi_liu
Private Sub MoveBlock(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer, ByVal icolor As Integer, ByVal direction As Integer)
    Dim Row As Integer, Col As Integer
    Dim i As Integer
    Dim old_row As Integer, old_col As Integer  '保存最早的中心坐标
    old_row = center_row
    old_col = center_col
    
    '首先擦除掉原来位置的
    Call EraseBlock(center_row, center_col, block)
    
    '-1 代表向左，1 代表向右，0 代表乡下
    Select Case direction
        Case Is = -1
            center_col = center_col - 1
        Case Is = 1
            center_col = center_col + 1
        Case Is = 0
            center_row = center_row + 1
    End Select
    
    '再绘制
    If CanMoveRotate(center_row, center_col, block) Then
        Call DrawBlock(center_row, center_col, block, icolor)
        '保存中心坐标
        iCenterRow = center_row
        iCenterCol = center_col
    Else
        Call DrawBlock(old_row, old_col, block, icolor)
        '保存中心坐标
        iCenterRow = old_row
        iCenterCol = old_col
        If direction = 0 Then
            bIsObjectEnd = True
        End If
    End If
    
    '保存方块坐标
    For i = 0 To 3
        MyBlock(i, 0) = block(i, 0)
        MyBlock(i, 1) = block(i, 1)
    Next
    
End Sub

Private Function CanMove(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer, ByVal icolor As Integer, ByVal direction As Integer)
    Dim Row As Integer, Col As Integer
    Dim i As Integer
    Dim old_row As Integer, old_col As Integer  '保存最早的中心坐标
    
    CanMove = True
    '首先擦除掉原来位置的，防止干扰
    Call EraseBlock(center_row, center_col, block)
    old_row = center_row
    old_col = center_col
    
    '-1 代表向左，1 代表向右，0 代表乡下
    Select Case direction
        Case Is = -1
            center_col = center_col - 1
        Case Is = 1
            center_col = center_col + 1
        Case Is = 0
            center_row = center_row + 1
    End Select
    
    For i = 0 To 3
        Row = center_row + block(i, 0)
        Col = center_col + block(i, 1)
        If Row > 20 Or Row < 0 Or Col > 11 Or Col < 2 Then      '越界
            CanMove = False
        End If
        If MySheet.Cells(Row, Col).Interior.Pattern <> xlNone Then  '只要有一个颜色，则为阻挡
            CanMove = False
        End If
    Next
    
    '恢复原来的图画
    Call DrawBlock(old_row, old_col, block, icolor)
End Function
'旋转方块函数 By@yaxi_liu
Private Sub RotateBlock(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer, ByVal icolor As Integer)
    Dim i As Integer
    '先擦除原来的
    Call EraseBlock(center_row, center_col, block)
    Dim tempArr(4, 2) As Integer
    '保存数组
    For i = 0 To 3
        tempArr(i, 0) = block(i, 0)
        tempArr(i, 1) = block(i, 1)
    Next
    '旋转后的坐标重新赋值
    For i = 0 To 3
        block(i, 0) = -tempArr(i, 1)
        block(i, 1) = tempArr(i, 0)
    Next i
    
    '重新绘制新的方块
    If CanMoveRotate(center_row, center_col, block) Then
        Call DrawBlock(center_row, center_col, block, icolor)
        '保存方块坐标
        For i = 0 To 3
            MyBlock(i, 0) = block(i, 0)
            MyBlock(i, 1) = block(i, 1)
        Next
    Else
        Call DrawBlock(center_row, center_col, tempArr, icolor)
        '保存方块坐标
        For i = 0 To 3
            MyBlock(i, 0) = tempArr(i, 0)
            MyBlock(i, 1) = tempArr(i, 1)
        Next
    End If
    
    '保存中心坐标
    iCenterRow = center_row
    iCenterCol = center_col
    
End Sub

'是否能够移动或者旋转函数，By@yaxi_liu
Private Function CanMoveRotate(ByVal center_row As Integer, ByVal center_col As Integer, ByRef block() As Integer) As Boolean
    '本函数形参均为变换后的坐标
    
    '首先判断是否越界
    Dim Row As Integer, Col As Integer
    Dim i As Integer
    CanMoveRotate = True
    For i = 0 To 3
        Row = center_row + block(i, 0)
        Col = center_col + block(i, 1)
        If Row > 20 Or Row < 0 Or Col > 11 Or Col < 2 Then      '越界
            CanMoveRotate = False
        End If
        If MySheet.Cells(Row, Col).Interior.Pattern <> xlNone Then  '只要有一个颜色，则为阻挡
            CanMoveRotate = False
        End If
    Next
End Function

'延时函数 By@yaxi_liu
Private Sub delay(T As Single)
    Dim T1 As Single
    T1 = Timer
    Do
        DoEvents
    Loop While Timer - T1 < T
End Sub
