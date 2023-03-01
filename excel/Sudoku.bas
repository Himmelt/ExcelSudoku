Attribute VB_Name = "Sudoku"
Sub Plot_Layout()
Attribute Plot_Layout.VB_Description = "绘制基础数独布局"
Attribute Plot_Layout.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Plot_Layout 宏
' 绘制基础数独布局
'
    ActiveSheet.Unprotect
    Cells.Select
    Selection.Clear
    
    Columns("A:Z").Select
    Selection.ColumnWidth = 6.83
    Columns("G:R").Select
    Selection.ColumnWidth = 1.83
    Rows("1:100").Select
    Selection.RowHeight = 14.3
    
    Range("K2:N5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    For i = 7 To 10
        With Selection.Borders(i)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next
    
    Range("G7:R18").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    For i = 7 To 10
        With Selection.Borders(i)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    Next
        
    Range("G9:R9,G12:R12,G15:R15").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("I7:I18,L7:L18,O7:O18").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range("K2:N5,G7:Q18").Locked = False
    Range("A1").Locked = False
    Range("A1").Select
    
    ActiveSheet.Protect
End Sub

Sub SetMatrix()
'
' 设置不可填区域
'

'
Dim num As Integer

Dim secs(3, 3) As Range
Dim section As Range
Dim row(3) As Integer
Dim lin(3) As Integer

ActiveSheet.Unprotect

Set secs(0, 0) = Range("G7:I9")
Set secs(0, 1) = Range("J7:L9")
Set secs(0, 2) = Range("M7:O9")
Set secs(0, 3) = Range("P7:R9")
Set secs(1, 0) = Range("G10:I12")
Set secs(1, 1) = Range("J10:L12")
Set secs(1, 2) = Range("M10:O12")
Set secs(1, 3) = Range("P10:R12")
Set secs(2, 0) = Range("G13:I15")
Set secs(2, 1) = Range("J13:L15")
Set secs(2, 2) = Range("M13:O15")
Set secs(2, 3) = Range("P13:R15")
Set secs(3, 0) = Range("G16:I18")
Set secs(3, 1) = Range("J16:L18")
Set secs(3, 2) = Range("M16:O18")
Set secs(3, 3) = Range("P16:R18")

    Range("G7:R18").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    For i = 0 To 3
        For j = 0 To 3
            If Cells(i + 2, j + 11) = "x" Then
                If section Is Nothing Then
                    Set section = secs(i, j)
                Else
                    Set section = Union(section, secs(i, j))
                End If
                row(i) = row(i) + 1
                lin(j) = lin(j) + 1
            End If
        Next
    Next
    
    For i = 0 To 3
        If row(i) <> 1 Or lin(i) <> 1 Then
            Range("A1").Select
            Range("A1").Locked = False
            ActiveSheet.Protect
            MsgBox ("无效的异形数独布局 !")
            Exit Sub
        End If
    Next
    
    section.Locked = True
    section.Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    Range("A1").Select
    ActiveSheet.Protect
End Sub


Public Sub RandMatrix()
    
Dim l(3) As Integer
Dim t As Integer
Dim i As Integer

ActiveSheet.Unprotect
' Range("K2:N5").ClearContents

Plot_Layout

i = 0
l(0) = -1
l(1) = -1
l(2) = -1
l(3) = -1

    Do While i <> 4
        Randomize
        xxx = Now
        bbb = Second(xxx)
        t = Int(Rnd(bbb) * 4)
        If Not ListContains(t, l) Then
            l(i) = t
            i = i + 1
        End If
    Loop
    
    Range("K2").Offset(0, l(0)) = "x"
    Range("K3").Offset(0, l(1)) = "x"
    Range("K4").Offset(0, l(2)) = "x"
    Range("K5").Offset(0, l(3)) = "x"
    
    SetMatrix
    
    Range("A1").Select
    ActiveSheet.Protect
End Sub


Private Function ListContains(num As Integer, arra() As Integer) As Boolean
   Dim length As Integer
   length = UBound(arra) - LBound(arra)
   For i = 0 To length
        If arra(i) = num Then
            ListContains = True
            Exit Function
        End If
    Next
    ListContains = False
End Function
