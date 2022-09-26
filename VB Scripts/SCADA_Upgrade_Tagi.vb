Option Explicit

Global x, y, z As Integer
Global wsResult As Worksheet
Global wbOld As Workbook
Global wbNew As Workbook
Global wsOld As Worksheet
Global wsNew As Worksheet
Global redColor As Long
Global yelColor As Long
Global bluColor As Long
'--------------------------------------------------------------------------------------------------------------------
Sub Compare()
    
    'application setup
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'variables for worksheet
    Dim data_v6_3, data_v7_5 As Variant
    
    '------------------------------------------------------------------------------------
    'Import old report data
    data_v6_3 = Application.GetOpenFilename(Title:="Browse OLD report from WinCC v6.3", _
    FileFilter:="Excel Files (*.xlsx),*xlsx*")
    
    If data_v6_3 = 0 Then
        MsgBox "No old report file selected."
        Exit Sub
    Else
        ActiveSheet.Cells(9, 2) = data_v6_3
    End If
    
    '------------------------------------------------------------------------------------
    'Import new report data
    data_v7_5 = Application.GetOpenFilename(Title:="Browse NEW report from WinCC v7.5", _
    FileFilter:="Excel Files (*.xlsx),*xlsx*")
    
    If data_v7_5 = False Then
        MsgBox "No new report file selected."
        Exit Sub
    Else
        ActiveSheet.Cells(10, 2) = data_v7_5
    End If
    
    '------------------------------------------------------------------------------------
    'check if files are imported
    If (data_v6_3 = 0) Or (data_v7_5 = 0) Then
        MsgBox "Files are not imported"
        Exit Sub
    End If
    
    '------------------------------------------------------------------------------------
    'variables for data compare
    Dim i, j, k, l As Integer
    Dim tagName As String
    Dim cellValue As String
    Dim lastRow, nextRow As Integer
    Set wsResult = ActiveSheet
    Set wbOld = Workbooks.Open(data_v6_3)
    Set wbNew = Workbooks.Open(data_v7_5)
    Set wsOld = wbOld.Worksheets(1)
    Set wsNew = wbNew.Worksheets(1)
        
    'get all rows
    lastRow = wsOld.Cells(Rows.Count, 1).End(xlUp).Row
    
    'define column/row for NOK tag log
    x = 1  'column
    y = 13   'row
    z = 1   'start log num # (do not change)
    redColor = RGB(255, 0, 0)   'red color
    yelColor = RGB(255, 255, 0) 'yellow color
    bluColor = RGB(102, 178, 255)   'blue color
    
    'old data
    For i = 30 To 100 'lastRow
        cellValue = wsOld.Cells(i, 1).Value
        
        'If cell is not empty and does not contain forbiden words
        If (IsEmpty(wsOld.Cells(i, 1)) = False) Then
            
            'get next tag position (to find the num of parameters)
            For j = (i + 1) To lastRow
                'check if next tag is the same as primary
                Dim sameVar: sameVar = StrComp((wsOld.Cells(j, 1).Value), cellValue, 1)
                Dim stringTest As Integer: stringTest = StrComp(wsOld.Cells(j, 1), "Tag Management", 1)
                
                If (IsEmpty(wsOld.Cells(j, 1)) = False) And sameVar <> 0 And stringTest <> 0 Then
                    nextRow = wsOld.Cells(j, 1).Row
                Exit For
                End If
            Next j
            
            'define string for storing parameters
            Dim oldTag() As String
            ReDim oldTag(1 To 9) As String
            Dim paramTest() As Integer
            ReDim paramTest(1 To 7) As Integer
    
            'save row position and tag name
            oldTag(1) = wsOld.Cells(i, 1).Row
            oldTag(2) = wsOld.Cells(i, 1).Value
                
            'Check for wanted parameters
            For k = i To j - 1
                'define parameters
                paramTest(1) = StrComp((wsOld.Cells(k, 2).Value), "Data Type", 1)
                paramTest(2) = StrComp((wsOld.Cells(k, 2).Value), "Group", 1)
                paramTest(3) = StrComp((wsOld.Cells(k, 2).Value), "Parameters", 1)
                paramTest(4) = StrComp((wsOld.Cells(k, 2).Value), "Connection", 1)
                paramTest(5) = StrComp((wsOld.Cells(k, 2).Value), "Channel", 1)
                paramTest(6) = StrComp((wsOld.Cells(k, 2).Value), "Channel unit", 1)
                paramTest(7) = StrComp((wsOld.Cells(k, 2).Value), "Unit", 1)

                'store parameters value
                For l = 1 To 7
                    If paramTest(l) = 0 Then
                        oldTag(l + 2) = wsOld.Cells(k, 3).Value
                    End If
                Next l
            Next k
            
            'function for compare created aray of old tag with new data base
            CheckNewParameters oldTag:=oldTag 'define the old tag array
                 
            'jump to next tag
            i = i + ((j - 1) - i)
        End If
    Next i
    
    'comment*
    CheckNumOfMergedTags
    
    'close the opened workbooks
    wbOld.Close
    wbNew.Close
    
End Sub
'--------------------------------------------------------------------------------------------------------------------
Function CheckNewParameters(ByRef oldTag() As String) As String
'variables
Dim i, j, k, l As Integer
Dim lastRow, nextRow As Integer
Dim isNotPresent As Boolean
Dim cellValue As String

'get all rows
lastRow = wsNew.Cells(Rows.Count, 1).End(xlUp).Row
'set merker for compare check
isNotPresent = True

'new data
For i = 1 To 100 'lastRow
    cellValue = wsNew.Cells(i, 1).Value
    Dim compName As Integer: compName = StrComp(cellValue, oldTag(2), 1)
    
    'if new tag is the same as old
    If compName = 0 Then
        'reset merker for compare check
        isNotPresent = False
        'get next tag position (to find the num of parameters)
        For j = (i + 1) To lastRow
            'check if next tag is the same as primary
            Dim sameVar: sameVar = StrComp((wsNew.Cells(j, 1).Value), cellValue, 1)
            Dim stringTest As Integer: stringTest = StrComp(wsNew.Cells(j, 1), "Tag Management", 1)
            
            If (IsEmpty(wsNew.Cells(j, 1)) = False) And sameVar <> 0 And stringTest <> 0 Then
                nextRow = wsNew.Cells(j, 1).Row
            Exit For
            End If
        Next j
        
        'define string for storing parameters
        Dim newTag() As String
        ReDim newTag(1 To 9) As String
        Dim paramTest() As Integer
        ReDim paramTest(1 To 7) As Integer

        'store row # and name
        newTag(1) = wsNew.Cells(i, 1).Row
        newTag(2) = wsNew.Cells(i, 1).Value

        'Check for wanted parameters
        For k = i To j - 1
            'define parameters
            paramTest(1) = StrComp((wsNew.Cells(k, 2).Value), "Data Type", 1)
            paramTest(2) = StrComp((wsNew.Cells(k, 2).Value), "Group", 1)
            paramTest(3) = StrComp((wsNew.Cells(k, 2).Value), "Parameters", 1)
            paramTest(4) = StrComp((wsNew.Cells(k, 2).Value), "Connection", 1)
            paramTest(5) = StrComp((wsNew.Cells(k, 2).Value), "Channel", 1)
            paramTest(6) = StrComp((wsNew.Cells(k, 2).Value), "Channel unit", 1)
            paramTest(7) = StrComp((wsNew.Cells(k, 2).Value), "Unit", 1)

            'store parameters value
            For l = 1 To 7
                If paramTest(l) = 0 Then
                    newTag(l + 2) = wsNew.Cells(k, 3).Value
                End If
            Next l
        Next k
        
        'debug
        'dim l
        'for l = 0 To 17
        '    Select Case l
        '        Case 0
        '            wsResult.Cells(y, x) = z
        '        Case Else
        '            wsResult.Cells(y, x + l) = newTag(l)
        'next l
        
        'function for compare created aray of old tag with new data base
        CheckBothParameters oldTag:=oldTag, newTag:=newTag 'define the old tag array
                
        'jump to next tag
        'Exit For
    End If
Next i

'if new tag does not exist
If isNotPresent Then
    'print out tags with error
    Dim m As Integer
    
    wsResult.Cells(y, x) = z    'old tag
    wsResult.Cells(y + 1, x) = z 'new tag
    'style
    wsResult.Cells(y, x).Interior.Color = yelColor
    wsResult.Cells(y + 1, x).Interior.Color = yelColor
    For m = 1 To 9
    
        wsResult.Cells(y, x + m) = oldTag(m)    'old tag
        wsResult.Cells(y + 1, x + m) = "?"  'new tag
        'style
        wsResult.Cells(y, x + m).Interior.Color = yelColor
        wsResult.Cells(y + 1, x + m).Interior.Color = yelColor
    Next m
    
    y = y + 2
    z = z + 1
End If
End Function
'--------------------------------------------------------------------------------------------------------------------
Function CheckBothParameters(ByRef oldTag() As String, ByRef newTag() As String) As String
    'variables
    Dim nok As Boolean
    Dim i, compParam As Integer
    Dim nokParameters() As Boolean
    ReDim nokParameters(1 To 9) As Boolean
    
    'set merker for compare check
    nok = False
    
    'check both arrays
    For i = 3 To 9
        compParam = StrComp(oldTag(i), newTag(i), 1)
        If compParam <> 0 Then
            nokParameters(i) = True
            nok = True
        End If
    Next i

    'if they do not match
    If nok Then
        'print out tags with error
        Dim j As Integer
        
        wsResult.Cells(y, x) = z    'old tag
        wsResult.Cells(y + 1, x) = z 'new tag
        For j = 1 To 9
            wsResult.Cells(y, x + j) = oldTag(j)   'old tag
            wsResult.Cells(y + 1, x + j) = newTag(j) 'new tag
            'style
            If nokParameters(j) = True Then
                wsResult.Cells(y, x + j).Interior.Color = redColor
                wsResult.Cells(y + 1, x + j).Interior.Color = redColor
            End If
        Next j
        
        y = y + 2
        z = z + 1
    End If
    
End Function
'--------------------------------------------------------------------------------------------------------------------
Function CheckNumOfMergedTags()

Dim i, j, k, l, m, n, lastRow, nextRow As Integer
Dim isNotPresent As Boolean
Dim cellValue, testString1, testString2, testString3 As String

'get all rows
lastRow = wsNew.Cells(Rows.Count, 1).End(xlUp).Row

For i = 1 To lastRow
    cellValue = wsNew.Cells(i, 1).Value
    Dim stringTest As Integer: stringTest = StrComp(cellValue, "Tag Management", 1)
    'set merker for compare check
    isNotPresent = True
    
    'If cell is not empty and does not contain forbiden words
    If (IsEmpty(wsNew.Cells(i, 1)) = False) And stringTest <> 0 Then
        For j = 1 To lastRow
            Dim sameVar: sameVar = StrComp((wsOld.Cells(j, 1).Value), cellValue, 1)

            'check if next tag is the same as primary
            If sameVar = 0 Then
               isNotPresent = False
            Exit For
            End If
        Next j

        'if tags not match
        If isNotPresent Then
            'get next tag position (to find the num of parameters)
            For k = (i + 1) To lastRow
                'check if next tag is the same as primary
                Dim nextVar: nextVar = StrComp((wsNew.Cells(k, 1).Value), cellValue, 1)
                stringTest = StrComp(wsNew.Cells(k, 1), "Tag Management", 1)
                
                If (IsEmpty(wsNew.Cells(k, 1)) = False) And nextVar <> 0 And stringTest <> 0 Then
                    nextRow = wsNew.Cells(k, 1).Row
                Exit For
                End If
            Next k

            'print values
            wsResult.Cells(y, x) = z    'old tag
            wsResult.Cells(y + 1, x) = z 'new tag
            wsResult.Cells(y, x + 1) = "?"  'old tag
            wsResult.Cells(y + 1, x + 1) = wsNew.Cells(i, 1).Row 'new tag
            wsResult.Cells(y, x + 2) = "?"  'old tag
            wsResult.Cells(y + 1, x + 2) = wsNew.Cells(i, 1).Value 'new tag

            'definiraj string za shranjevanje parametrov
            Dim paramTest() As Integer
            ReDim paramTest(1 To 7) As Integer

            'Check for wanted parameters
            For l = i To k - 1
                'define parameters
                paramTest(1) = StrComp((wsNew.Cells(l, 2).Value), "Data Type", 1)
                paramTest(2) = StrComp((wsNew.Cells(l, 2).Value), "Group", 1)
                paramTest(3) = StrComp((wsNew.Cells(l, 2).Value), "Parameters", 1)
                paramTest(4) = StrComp((wsNew.Cells(l, 2).Value), "Connection", 1)
                paramTest(5) = StrComp((wsNew.Cells(l, 2).Value), "Channel", 1)
                paramTest(6) = StrComp((wsNew.Cells(l, 2).Value), "Channel unit", 1)
                paramTest(7) = StrComp((wsNew.Cells(l, 2).Value), "Unit", 1)

                For m = 1 To 7
                    'store parameters value
                    If paramTest(m) = 0 Then
                        wsResult.Cells(y, x + 2 + m) = "?"
                        wsResult.Cells(y + 1, x + 2 + m) = wsNew.Cells(l, 3).Value
                    End If
                Next m
            Next l
               
            'style
            wsResult.Cells(y, x).Interior.Color = bluColor
            wsResult.Cells(y + 1, x).Interior.Color = bluColor
            For n = 1 To 9
                'style
                wsResult.Cells(y, x + n).Interior.Color = bluColor
                wsResult.Cells(y + 1, x + n).Interior.Color = bluColor
            Next n
               
            y = y + 2
            z = z + 1
        End If
    End If
Next i
End Function
'--------------------------------------------------------------------------------------------------------------------
Sub Clear()
Dim i, j As Integer
Dim lastRow As Integer

'application setup
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

If lastRow <= 13 Then
    MsgBox "Table Empty"
End If

'ActiveSheet.Cells(9, 2).Delete
'ActiveSheet.Cells(10, 2).Delete

Do While lastRow >= 13  'preskakuje vrstice
    For i = 13 To lastRow 'define the data start row
        ActiveSheet.Cells(i, 1).EntireRow.Delete
    Next i
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
Loop

End Sub
