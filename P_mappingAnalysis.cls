Option Explicit

Sub MappingDataAnalysis()
    If IsEmpty(Range("A1")) Then
        Range("A1").Value = "Date/Time"
    End If
    
    ' To go through cells in the table
    Dim ccellcount, rcellcount, rbelowdata, rcount As Integer
    Dim drows, dcols, rbcount As Integer
    'To store the end of the table cell value
    Dim b As String
    
    ' reseting variables at zero
    ccellcount = 0
    rcellcount = 0
    
    'Checking the top row of the table until
    'we find an empty cell
    Do While Not IsEmpty(Range("A1").Offset(rcellcount, ccellcount))
        Range("A1").Offset(rcellcount, ccellcount).Select
        'MsgBox ">> " + CStr(ccellcount) + " Has Value"
        ccellcount = ccellcount + 1
        
    Loop
    dcols = ccellcount
    'Reset variable to the last functioning non empty position
    ccellcount = ccellcount - 1
    'MsgBox "Number of Columns " & dcols
    
    ' Checking last column until find empty cell
    Do While Not IsEmpty(Range("A1").Offset(rcellcount, ccellcount))
        Range("A1").Offset(rcellcount, ccellcount).Select
        'MsgBox ">> " + CStr(rcellcount) + " Has Value"
        rcellcount = rcellcount + 1
        
    Loop
    drows = rcellcount
    'Reset variable to the last functioning non empty position
    rcellcount = rcellcount - 1
    'MsgBox "Number of Rows " & drows
    
    
    'Record address of bottom far right cell
    b = ActiveCell.Address(False, False)
    'Select entire table
    Range("A1", b).Select
    
    'Setting the analysis area below the datat
    rbelowdata = drows + 3
    rbcount = 0
    'Msgbox Range("A1", b).count
    
    'Moving outside probe to the end of the table
    Dim omit_probe As Boolean
    omit_probe = outsideprobe()
    
    'Use range code under  comment"select entire table" to make a table
    '>>>>Make table named: Mapping_Data
    
    
    'Naming and calculating the average of first column
    'Assign variables for average row
    Dim avgrow, avgbcol, avgecol As Integer
    Dim r_avg_up, r_avg_down, f_avgaddress, l_avgaddress As String
    avgrow = rbelowdata
    'location relative data
    r_avg_up = "-" + CStr(avgrow - 1)
    r_avg_down = "-" + CStr(avgrow - drows + 1)
    
    Range("A1").Offset(avgrow).Select
    ActiveCell.FormulaR1C1 = "Average"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=AVERAGE(R[" + r_avg_up + "]C:R[" + r_avg_down + "]C)"
    f_avgaddress = ActiveCell.Offset(, 1).Address(False, False)
    l_avgaddress = ActiveCell.Offset(, ccellcount).Address(False, False)
    rbcount = rbcount + 1
    
    'Naming and calculating the Maximum of first column
    'Assign variables for Maximum row
    Dim maxrow, maxbcol, maxecol As Integer
    Dim r_max_up, r_max_down, f_maxaddress, l_maxaddress As String
    maxrow = rbelowdata + 1
    'location relative data
    r_max_up = "-" + CStr(maxrow - 1)
    r_max_down = "-" + CStr(maxrow - drows + 1)
    
    Range("A1").Offset(maxrow).Select
    ActiveCell.FormulaR1C1 = "Maximum"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MAX(R[" + r_max_up + "]C:R[" + r_max_down + "]C)"
    f_maxaddress = ActiveCell.Offset(, 1).Address(False, False)
    l_maxaddress = ActiveCell.Offset(, ccellcount).Address(False, False)
    rbcount = rbcount + 1

    'Naming and calculating the Minimum of first column
    'Assign variables for Minimum row
    Dim minrow, minbcol, minecol As Integer
    Dim r_min_up, r_min_down, f_minaddress, l_minaddress As String
    minrow = rbelowdata + 2
    'location relative data
    r_min_up = "-" + CStr(minrow - 1)
    r_min_down = "-" + CStr(minrow - drows + 1)
    
    Range("A1").Offset(minrow).Select
    ActiveCell.FormulaR1C1 = "Minimum"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MIN(R[" + r_min_up + "]C:R[" + r_min_down + "]C)"
    f_minaddress = ActiveCell.Offset(, 1).Address(False, False)
    l_minaddress = ActiveCell.Offset(, ccellcount).Address(False, False)
    rbcount = rbcount + 1

    'MKT
    'Naming and calculating the MKT of first column
    'Assign variables for MKT row
    Dim MKTrow, MKTbcol, MKTecol As Integer
    Dim r_MKT_up, r_MKT_down, f_MKTaddress, l_MKTaddress As String
    MKTrow = rbelowdata + 3
    'location relative data
    r_MKT_up = "-" + CStr(MKTrow - 1)
    r_MKT_down = "-" + CStr(MKTrow - drows + 1)
    
    Range("A1").Offset(MKTrow).Select
    ActiveCell.FormulaR1C1 = "MKT"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MeanKT(R[" + r_MKT_up + "]C:R[" + r_MKT_down + "]C)"
    f_MKTaddress = ActiveCell.Offset(, 1).Address(False, False)
    l_MKTaddress = ActiveCell.Offset(, ccellcount).Address(False, False)
    rbcount = rbcount + 1
    
    'AutoFilling the remaining column calculation
    Range(f_avgaddress, f_minaddress).Select
    Dim lastcell_address, lrcell As String
    If rbcount = 4 Then
        Range(f_avgaddress, f_MKTaddress).Select
        lrcell = MKTrow
    Else
        lrcell = minrow
    End If
    
    lastcell_address = Range("A1").Offset(lrcell, ccellcount).Address(False, False)
    
    'Use selection instead of activecell
    Selection.AutoFill Destination:=Range(f_avgaddress, lastcell_address)
        
        
        
'    'moving cell to the the end of the data set
'
'    Dim mo_mode, max_value, min_value As Integer
'    Dim rcell, avgrange As Range
'    Set avgrange = Range(f_avgaddress, l_avgaddress)
    
'    mo_mode = Roundmode(avgrange)
'    'define the min and max of the input range
'    max_value = Application.WorksheetFunction.Max(avgrange)
'    min_value = Application.WorksheetFunction.Min(avgrange)
'
'    For Each rcell In avgrange
'        If (mo_mode > rcell.Value) And (mo_mode - rcell.Value > 10) Then
'            min_value = 1
'            Application.CutCopyMode = False ' don't want an existing operation to interfere
'            Columns(ccellcount + 2).Insert XlDirection.xlToRight
'            Columns(ccellcount + 2).Value = Columns(rcell.Column).Value ' this would be one greater if to the right of F
'            Columns(rcell.Column).Delete
'        End If
'    Next rcell
    
    If omit_probe Then
        l_avgaddress = Range("A1").Offset(avgrow, ccellcount - 1).Address(False, False)
        l_minaddress = Range("A1").Offset(minrow, ccellcount - 1).Address(False, False)
        l_maxaddress = Range("A1").Offset(maxrow, ccellcount - 1).Address(False, False)
        l_MKTaddress = Range("A1").Offset(MKTrow, ccellcount - 1).Address(False, False)
     End If
        
  
    'Calculate Overall Max, Min Avg and Max/Min Mean & MKT
    Dim rlocal, ro_mean, ro_max, ro_min, rmax_mean, rmin_mean, ro_mkt, rmax_mkt, rmin_mkt As Integer
    Dim adrs_rmax_mean, adrs_rmin_mean, adrs_rmax_mkt, adrs_ro_min, adrs_ro_max, adrs_ro_mean As String
    'Significant figues
    Dim s_figures As Integer
    s_figures = CInt(InputBox("Enter number of significant figures"))
    
    
    rlocal = rbelowdata + rbcount + 2
    
    'Overall Mean
    ro_mean = rlocal
    Range("A1").Offset(ro_mean).Select
    ActiveCell.FormulaR1C1 = "Overall Mean"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=AVERAGE(" + f_avgaddress + ": " + l_avgaddress + ")"
    adrs_ro_mean = ActiveCell.Offset(, 1).Address
    
    'Overall Max
    ro_max = rlocal + 1
    Range("A1").Offset(ro_max).Select
    ActiveCell.FormulaR1C1 = "Overall Max"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MAX(" + f_maxaddress + ": " + l_maxaddress + ")"
    adrs_ro_max = ActiveCell.Offset(, 1).Address
    Call ReturnSensorID(Range(f_maxaddress, l_maxaddress), adrs_ro_max, s_figures)
    
    'Overall Min
    ro_min = rlocal + 2
    Range("A1").Offset(ro_min).Select
    ActiveCell.FormulaR1C1 = "Overall Min"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MIN(" + f_minaddress + ": " + l_minaddress + ")"
    adrs_ro_min = ActiveCell.Offset(, 1).Address
    Call ReturnSensorID(Range(f_minaddress, l_minaddress), adrs_ro_min, s_figures)
    
    'Max Average
    rmax_mean = rlocal + 3
    Range("A1").Offset(rmax_mean).Select
    ActiveCell.FormulaR1C1 = "Max Average"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MAX(" + f_avgaddress + ": " + l_avgaddress + ")"
    adrs_rmax_mean = ActiveCell.Offset(, 1).Address
    Call ReturnSensorID(Range(f_avgaddress, l_avgaddress), adrs_rmax_mean, s_figures)
    
    'Min Average
    rmin_mean = rlocal + 4
    Range("A1").Offset(rmin_mean).Select
    ActiveCell.FormulaR1C1 = "Min Average"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MIN(" + f_avgaddress + ": " + l_avgaddress + ")"
    adrs_rmin_mean = ActiveCell.Offset(, 1).Address
    Call ReturnSensorID(Range(f_avgaddress, l_avgaddress), adrs_rmin_mean, s_figures)
    
    'Max MKT'
    rmax_mkt = rlocal + 5
    Range("A1").Offset(rmax_mkt).Select
    ActiveCell.FormulaR1C1 = "Max MKT"
    'Give gray color to the cell with labels
    Call Graycolorformat(ActiveCell)
    ActiveCell.Offset(, 1).Formula = "=MAX(" + f_MKTaddress + ": " + l_MKTaddress + ")"
    adrs_rmax_mkt = ActiveCell.Offset(, 1).Address
    Call ReturnSensorID(Range(f_MKTaddress, l_MKTaddress), adrs_rmax_mkt, s_figures)
    
    'Arrange Cells Sizes
    Cells.Select
    Selection.Columns.AutoFit
    
    Dim f_rows As Integer
    f_rows = rcellcount
    Call MakeFormulaSheet(f_rows, False)
    
End Sub

Function Roundmode(rng As Range) As Double
    Dim mi_mode, mi_count, li_count As Integer
    Dim cell, cellu As Range
    li_count = 0
    'looping through all the cells and comparing round values
    For Each cell In rng
        mi_count = 0
        For Each cellu In rng
            If Round(cell.Value) = Round(cellu.Value) Then
                mi_count = mi_count + 1
            End If
        Next cellu
            'Compare the counted match and associate the largest one as the new mode
        If mi_count > li_count Then
            li_count = mi_count
            mi_mode = Round(cell.Value)
        End If
    Next cell
    Roundmode = mi_mode
End Function

Sub ReturnSensorID(r_range As Range, v_match As Variant, s_figures As Integer)
    Dim cell As Range
    Dim Id_count As Integer
    Dim sn As Variant
    
    Id_count = 0
    'Get number of significant figures to use
    's_figures = CInt(InputBox("Enter number of significant figures"))
    'THis macro should be able to return the SN at the top of the column
    For Each cell In r_range
        If Round(Range(v_match).Value, s_figures) = Round(cell.Value, s_figures) Then
            Id_count = Id_count + 1
            sn = CStr(Range("A1").Offset(, cell.Column - 1).Value)
            Range(v_match).Offset(, Id_count).Value = sn
            Call Graycolorformat(Range(v_match).Offset(, Id_count))
        End If
    Next cell
End Sub

Function outsideprobe() As Boolean
    Dim i, num_columns As Integer
    Dim cell As Range
    Dim msg1, input_SN As String
    num_columns = Application.WorksheetFunction.CountA(Rows("1"))
    msg1 = "Enter Outside probe Serial #"
    input_SN = InputBox(msg1, "Probe To Omit", "N/A")
    outsideprobe = False
    'loop through column header to find match
    For i = 1 To (num_columns)
        If Range("A1").Offset(, i).Value = input_SN Then
            Range("A1").Offset(, i).Select
            MsgBox "Found it " & input_SN
            Application.CutCopyMode = False ' don't want an existing operation to interfere
            'Moving the match to the end of the table
            Columns(num_columns + 1).Insert XlDirection.xlToRight
            Columns(num_columns + 1).Value = Columns(Range("A1").Offset(, i).Column).Value ' this would be one greater if to the right
            Columns(Range("A1").Offset(, i).Column).Delete
            'Acknowledging a move was made
            'This should be a numerical value in the future
            outsideprobe = True
        End If
    Next i
End Function

Sub Graycolorformat(user_input As Range)
    'Puts a gray background color on the cells placed in the input
    Dim cell As Range
    'loop through cell input and apply the theme color
    For Each cell In user_input
        cell.Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.149998474074526
            .PatternTintAndShade = 0
        End With
    Next cell
End Sub

