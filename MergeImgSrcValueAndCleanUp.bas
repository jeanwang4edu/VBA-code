Attribute VB_Name = "Module2"
Sub MergeImgSrcValueAndCleanUp()
'For MRIG site to clean shopify imported data

    'Define which sheet you want to work on
    Dim sheetName As String
    sheetName = "Sheet1"

    'Declare the column where ImgSrc is at
    Dim colImgSrc, colImgSrcOne As String
    colImgSrc = "G"
    colImgSrcOne = "G1"

    'Get all the Range with the same handle
    Dim lastrow As Long
    Dim xRg As Range, yRg As Range
    'change Sheet1 to suit your need
    With ThisWorkbook.Worksheets(sheetName)
    
        'Get total column number
        lastCol = .Cells(1, Columns.Count).End(xlToLeft).Column
        lastrow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Application.ScreenUpdating = False
        For Each xRg In .Range("A1:A" & lastrow)
            'Put in the handle value
            Dim handleValue As String
            handleValue = "19-day-feast-pages-for-kids-issue-6-baha-splendor-naw-ruz"
            If LCase(xRg.Text) = handleValue Then
                If yRg Is Nothing Then
                    Set yRg = .Range("A" & xRg.Row).Resize(, lastCol)
                Else
                    Set yRg = Union(yRg, .Range("A" & xRg.Row).Resize(, lastCol))
                End If
            End If
        Next xRg
        Application.ScreenUpdating = True
    End With

    'Get the result Range and Select
    If Not yRg Is Nothing Then yRg.Select
    
    Dim myArea As Range, b As Range
    Set myArea = yRg
    
    'Merge the ImgSrc column for the Range And trim the last ","
    Dim newValue As String
    For Each b In myArea.Rows
        'change D1 to suit your need
        newValue = newValue + b.Range(colImgSrcOne).Value + ", "
    Next
    'MsgBox "new Value: " + newValue
    Dim myLen As Integer
    myLen = Len(newValue)
    Dim newValue2 As String
    newValue2 = Left(newValue, myLen - 2)
    'MsgBox "new Value2: " + newValue2
    Dim MyRange As String
    MyRange = colImgSrc + CStr(myArea.Row)
    'MsgBox "here1: " + MyRange
    
    'Replace the first row's ImgSrc value with newValue2
    With ThisWorkbook.Worksheets(sheetName)
        .Range(MyRange).Value = newValue2
        'MsgBox "hey!: " + .Range(MyRange)
    End With
    
    'Delete all the other rows
    Dim MyArray() As Integer
    Dim size As Integer
    size = myArea.Rows.Count
    ReDim MyArray(size)
    
    Dim i As Integer
    For Each b In myArea.Rows
        i = i + 1
        If i <> 1 Then
           'MsgBox "This is not the first row in the area!"
           'MsgBox "hey: " + CStr(b.Row)
           MyArray(i) = b.Row
        End If
    Next
    
    With ThisWorkbook.Worksheets(sheetName)
        Dim j As Integer
        j = MyArray(2)
        For i = 2 To UBound(MyArray)
           MsgBox "Deleting... " + CStr(MyArray(i))
           'MsgBox "I want to delete: " + .Range(MyRange).Value
           .Rows(j).Delete
        Next i
    End With
    
End Sub

