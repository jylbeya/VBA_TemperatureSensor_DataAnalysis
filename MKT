Option Explicit

Function MeanKT(Temperatures) As Double
    'This Code will Calculate MKT from a range of Celsuis
    'Inputs and Declaring variables
    Dim cell As Range
    Dim Ckelvin, HRT, eHRT, Ans As Double
    Dim Tkelvin, HR, Rgas, Hactiv, n, denom As Double
    Static count
    count = 0
    
    ' Defining Constants and Variable for calculation
    Tkelvin = 273.15
    Rgas = 0.008314472
    Hactiv = 83.14472
    HR = -Hactiv / Rgas
    denom = 0
    n = Temperatures.count
    
    
    'Looping through Every cell in a Range
    For Each cell In Temperatures
        Ckelvin = cell.value + Tkelvin
        HRT = HR / Ckelvin
        eHRT = Exp(HRT)
        denom = denom + eHRT
    Next cell
    
    'Grouping and Final Analysis
    'This functions follows the following Formula
    '(-H/R)/ln((Sum(e^(H/RT1)+..+e^(-H/RTn)))/n)
    denom = denom / n
    denom = Application.WorksheetFunction.Ln(denom)
    MeanKT = HR / denom
    MeanKT = MeanKT - Tkelvin
    
    'By JYKB :)
End Function

Function FindColumnValue(Tofind, List)
    Dim cell As Range
    Dim count As Double
    count = 0
    For Each cell In List
            If cell.value = Tofind Then
                FindColumnValue = cell.Address
                count = count + 1
            End If
    Next cell
    FindColumnValue = count
End Function

Function PrintBox(value)
    MsgBox value
End Function



