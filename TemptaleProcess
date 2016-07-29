Option Explicit

Sub organizetemptaledata()
    'This macro is buit to organize Point Base TempTale Data
    
    Dim num_col, min_row, max_row, i As Integer
    Dim min_adrs, lmax_adrs, msg1, msg2, date_entry, gap_entry As String
    
    min_row = Application.WorksheetFunction.CountA(Columns("A"))
    max_row = Application.WorksheetFunction.CountA(Columns("A"))
    
    'look through the top row and identifying minimum and maximum number of rows
    If Range("A1").Value = "Point" Then
        'Count number of not empty columns in the database
        num_col = Application.WorksheetFunction.CountA(Rows("1"))
        'loop through each column to identify the min and max row
        For i = 1 To (num_col)
            If min_row > Application.WorksheetFunction.CountA(Columns(i)) Then
                min_row = Application.WorksheetFunction.CountA(Columns(i))
            ElseIf max_row < Application.WorksheetFunction.CountA(Columns(i)) Then
                max_row = Application.WorksheetFunction.CountA(Columns(i))
            End If
        Next i
        
        'Deleting the hanging rows
        If Not (min_row = max_row) Then
            min_adrs = Range("A1").Offset(min_row).Address
            lmax_adrs = Range("A1").Offset(max_row - 1, num_col - 1).Address
            Range(min_adrs, lmax_adrs).Select
            Selection.Delete
        End If
        'Removing the Min, Max, Std Dev.... columns
        If Range("B1").Value = "Min" And Range("F1").Value = "Std Dev" Then
            Range("B1", Range("F1").Offset(min_row - 1)).Select
            Selection.Delete
        End If
        
            'Parsing column header to only leave serial number
    'only works if the serial number is the last part of the string
    For i = 0 To (num_col - 1)
        Range("A1").Offset(, i).Value = Right(Range("A1").Offset(, i).Value, 10)
    Next i
    
    'Arrange Cells Sizes
    Cells.Select
    Selection.Columns.AutoFit
    
    'Make a copy of sheet 1
    Application.CutCopyMode = False
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(1)
    Sheets("Sheet1").Name = "Point Base Data"
    Sheets("Sheet1 (2)").Name = "Date Base Data"
    Sheets("Date Base Data").Select
    Range("A1").Value = "Date/Time"
    
    'Take Start time and gap form the use and autofill the rest of the cell
    '>>> fix verification of the 1 second deviation
    '>>> Fix the date Format it does not appear properly on the
    msg1 = "Enter Start Date and Time " & vbCrLf & "(mm/dd/yyyy hh:mm:ss)"
    date_entry = InputBox(msg1, "Sensor Start Time", Format(Now(), "mm/dd/yyyy hh:mm:ss"))
    date_entry = Format(date_entry, "mm/dd/yyyy hh:mm:ss")
    msg2 = "Enter Start Date and Time " & vbCrLf & date_entry & vbCrLf & " + gap interval " & vbCrLf & "(mm/dd/yyyy hh:mm:ss)"
    gap_entry = InputBox(msg2, "Next Data Point Time", date_entry)
    gap_entry = Format(gap_entry, "mm/dd/yyyy hh:mm:ss")
    MsgBox "Your start date and time is " & date_entry
    MsgBox "Your interval " & gap_entry
    Range("A1").Offset(1) = Format(date_entry, "mm/dd/yyyy hh:mm:ss")
    Range("A1").Offset(2) = Format(gap_entry, "mm/dd/yyyy hh:mm:ss")
    Range(Range("A1").Offset(1), Range("A1").Offset(2)).Select
    Selection.AutoFill Destination:=Range(Range("A1").Offset(1), Range("A1").Offset(min_row - 1))
    
    'Arrange Cells Sizes
    Cells.Select
    Selection.Columns.AutoFit
    
    Else
        MsgBox "Download Point Base Data "
    End If

    Call MappingDataAnalysis
    
End Sub

Sub MakeFormulaSheet(num_rows As Integer, cal As Boolean)
    Dim input_rows As String
    If Not cal Then
        input_rows = "2:" + CStr(num_rows - 2)
    Else
        input_rows = "2:" + CStr(num_rows - 12)
    End If
    'Make a copy of sheet 1
    Application.CutCopyMode = False
    ActiveSheet.Copy After:=Sheets(1)
    ActiveSheet.Name = "Formula"
    ActiveWindow.DisplayFormulas = True
    Rows(input_rows).Select
    Selection.EntireRow.Hidden = True
    
    'Arrange Cells Sizes
    Cells.Select
    Selection.Columns.AutoFit
    
End Sub
