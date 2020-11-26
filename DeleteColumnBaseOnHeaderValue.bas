Attribute VB_Name = "Module1"
Sub FindColumnByHeaderValue()
' Updateby jeanwang2dev
    MsgBox "VBA: Finding Columns by Header Value"
    
    ' Define which sheet you want to work on
    Dim sheetName As String
    sheetName = "Sheet1"
    
    ' Get the Range
    Dim rangeAdd As String
    rangeAdd = "A1:" + Split(Columns(Range("A1").End(xlToRight).Column).Address(, False), ":")(1) + "1"
    
    ' Declare an array to store the header value of all the columns we want to delete
    Dim sheetHeaderKeyword As String
    sheetHeaderKeyword = "Sheet4"
    Dim valueArray() As String
    Dim size As Integer
    With ThisWorkbook.Worksheets(sheetHeaderKeyword)
        size = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    MsgBox "Size: " + CStr(size)
    ReDim valueArray(size)
    Dim ii As Integer
    ii = 1
    With ThisWorkbook.Worksheets(sheetHeaderKeyword)
        ' Select cell A1, *first line of data*.
        Sheets(sheetHeaderKeyword).Activate
        .Range("A1").Select
        ' Set Do loop to stop when an empty cell is reached.
        Do Until IsEmpty(ActiveCell)
           ' Insert your code here.
           valueArray(ii) = ActiveCell.Value
           ii = ii + 1
           ' Step down 1 row from present location.
           ActiveCell.Offset(1, 0).Select
        Loop
    End With
    
    With ThisWorkbook.Worksheets(sheetName)
        Sheets(sheetName).Activate
        For i = 1 To UBound(valueArray)
            'MsgBox "Array value " + CStr(i) + " : " + valueArray(i)
            Dim xRg As Range
            Dim xRgUni As Range
            Dim xFirstAddress As String
            Dim xStr As String
            On Error Resume Next
            xStr = valueArray(i)
            MsgBox "Delete Column " + xStr
            Set xRg = Range(rangeAdd).Find(xStr, , xlValues, xlWhole, , , True)
            If Not xRg Is Nothing Then
                xFirstAddress = xRg.Address
                Do
                    Set xRg = Range(rangeAdd).FindNext(xRg)
                    If xRgUni Is Nothing Then
                        Set xRgUni = xRg
                    Else
                        Set xRgUni = Application.Union(xRgUni, xRg)
                    End If
                Loop While (Not xRg Is Nothing) And (xRg.Address <> xFirstAddress)
            End If
            xRgUni.EntireColumn.Select
        Next i
        xRgUni.EntireColumn.Delete
    End With
    
End Sub


Sub GetLastRow()
    Dim lastrow As Long
    With ThisWorkbook.Worksheets("Sheet2")
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        MsgBox "Last Row: " + CStr(lastrow)
    End With
End Sub

Sub GetLastColumn()
    Dim lastColumn As Long
    With ThisWorkbook.Worksheets("Sheet2")
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        MsgBox "Last Row: " + CStr(lastrow)
    End With
End Sub

Sub GetLastColumnLetter()
    Dim lastCol$
    lastCol = Split(Columns(Range("A1").End(xlToRight).Column).Address(, False), ":")(1)
    MsgBox lastCol
End Sub

Sub Test2()

   Dim arrayStr(3) As String
   Dim ii As Integer
   ii = 1
   
   With ThisWorkbook.Worksheets("Sheet5")
        ' Select cell A2, *first line of data*.
        Range("A1").Select
        ' Set Do loop to stop when an empty cell is reached.
        Do Until IsEmpty(ActiveCell)
           ' Insert your code here.
           ' MsgBox "hello" + ActiveCell.Value
           arrayStr(ii) = ActiveCell.Value
           ii = ii + 1
           ' Step down 1 row from present location.
           ActiveCell.Offset(1, 0).Select
        Loop
   End With
   
   For i = 1 To UBound(arrayStr)
        MsgBox "arrayStr Value: " + arrayStr(i)
   Next i
End Sub
