Attribute VB_Name = "Module1"
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
    
    Range("A1").Select
    
    ActiveSheet.Protect
End Sub

Sub SetMatrix()
'
' 设置不可填区域
'

'
Dim num As Integer

Dim secs(4, 4) As Range
Dim section As Range
Dim row(4) As Integer
Dim lin(4) As Integer

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
            MsgBox ("invalid")
            Exit Sub
        End If
    Next
    
    section.Locked = True
    section.Select

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    
    Range("A1").Select
    ActiveSheet.Protect
End Sub

