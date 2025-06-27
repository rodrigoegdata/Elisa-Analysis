Attribute VB_Name = "Module11"
Sub GenReport()
Attribute GenReport.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' Generates a report for every sheet.
'
Dim x As Integer

For x = 1 To Worksheets.Count

Worksheets(x).Select
Call Report

Next x

End Sub
Sub Report()
Attribute Report.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Generates a report for the active sheet.
'
Call maincalc_1
Call minusblank_2
Call standardcurv_3
Call ClearValuesOutsideRange_4
Call RegressionPrediction_5
Call ClearNegativeValues_6
Call ConcentCalc_7
Call resultCalc_8
Call dfCalc_9
Call formatcells_10

End Sub
Sub maincalc_1()
Attribute maincalc_1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Calculates the average of the blank wells excluding outliers based on Tukey's fences method.
'
    Range("A26").Value = "Raw data"
    Range("Q25").Value = "Blank selection and averaging"
    Range("Q26").Value = "Q1"
    Range("R26").Formula = "=QUARTILE.INC(R[1]C[-15]:R[8]C[-15],1)"
    Range("Q27").Value = "Q3"
    Range("R27").Formula = "=QUARTILE.INC(RC[-15]:R[7]C[-15],3)"
    Range("Q28").Value = "IQR"
    Range("R28").Select
    Selection.Formula = "=R[-1]C-R[-2]C"
    Range("Q29").Value = "lowBd"
    Range("R29").Formula = "=R[-3]C-(1.5*R[-1]C)"
    Range("Q30").Value = "upBd"
    Range("R30").Formula = "=R[-3]C+(1.5*R[-2]C)"
    Range("S26").Value = "Averagelbub"
    Range("T26").Formula = _
        "=AVERAGEIFS(R[1]C[-17]:R[8]C[-17],R[1]C[-17]:R[8]C[-17],"">=""&R[3]C[-2],R[1]C[-17]:R[8]C[-17],""<=""&R[4]C[-2])"
    
End Sub
Sub minusblank_2()
Attribute minusblank_2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Subtracts the average blank from all values in the plate.
'

'
    Range("A37").Value = "Blank subtraction"
    Range("B26:B34").Copy Range("B37:B45")
    Range("C26:N26").Copy Range("C37:N37")
    Range("C38").Formula = "=R[-11]C-R26C20"
    Range("C38").Select
    Selection.AutoFill Destination:=Range("C38:C45"), Type:=xlFillDefault
    Range("C38:C45").Select
    Selection.AutoFill Destination:=Range("C38:N45"), Type:=xlFillDefault
    Range("C38:N45").Select

End Sub
Sub standardcurv_3()
Attribute standardcurv_3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Creates a table extracting the absorbance of the standard curve vs concentration. Calculates the Slope and the Intercept and plots the data in a scatter chart.
'

'
    Range("Q36").Value = "Standard curve"
    Range("Q37").Value = "Absorbance"
    Range("R37").Value = "Concentration (ng/mL)"
    Range("Q38").Select
    Selection.FormulaR1C1 = "=RC[-13]"
    Selection.AutoFill Destination:=Range("Q38:Q45"), Type:=xlFillDefault
    Range("R38").Value = "100"
    Range("R39").Formula = "=R[-1]C/2"
    Range("R39").Select
    Selection.AutoFill Destination:=Range("R39:R45"), Type:=xlFillDefault
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter, Left:=Range("T33").Left, Top:=Range("T33").Top).Select
    ActiveChart.SetSourceData Source:=Range("$Q$37:$R$45")
        ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Trendlines.Add Type:=xlLinear, Forward _
        :=0, Backward:=0, DisplayEquation:=0, DisplayRSquared:=0, Name:= _
        "Linear (Conc)"
    ActiveChart.FullSeriesCollection(1).Trendlines(1).Select
    Selection.DisplayEquation = True
    Selection.DisplayRSquared = True
    Application.CommandBars("Format Object").Visible = False
    
    Range("Q47").Value = "Slope"
    Range("R47").Value = "Intercept"
    Range("Q48").Formula2R1C1 = "=LINEST(R[-10]C[1]:R[-3]C[1],R[-10]C:R[-3]C)"
End Sub

Sub ClearValuesOutsideRange_4()
'
' Removes all values with absorbance higher or lower than the standard curve.
'
    Dim cell As Range
    Dim lowerBound As Double
    Dim upperBound As Double
    
    ' Get bounds from cells
    lowerBound = Range("D45").Value
    upperBound = Range("D38").Value

    ' Loop through the target range
    For Each cell In Range("C38:N45")
        If IsNumeric(cell.Value) Then
            If cell.Value < lowerBound Or cell.Value > upperBound Then
                cell.ClearContents
            End If
        End If
    Next cell
End Sub

Sub RegressionPrediction_5()
'
' Interpolates the absorbance values of each sample in the standard curve to calculate the concentration in ng/mL.
'

'
    Range("A48").Value = "Equation application"
    Range("B26:B34").Copy Range("B48:B56")
    Range("C26:N26").Copy Range("C48:N48")
    Range("C49").Select
    Selection.Formula = "=$Q$48*C38+$R$48"
    Selection.AutoFill Destination:=Range("C49:C56"), Type:=xlFillDefault
    Range("C49:C56").Select
    Selection.AutoFill Destination:=Range("C49:N56"), Type:=xlFillDefault
     
End Sub

Sub ClearNegativeValues_6()
'
' Removes all negative values after sample interpolation.
'
    Dim cell As Range
            
' Loop through the target range
    For Each cell In Range("C49:N56")
        If IsNumeric(cell.Value) Then
            If cell.Value <= 0 Then
                cell.ClearContents
            End If
        End If
    Next cell
End Sub
Sub ConcentCalc_7()
'
' Multiples all values by the dilution factor of the serial dilution and removes values that are <= 0.
'

'
    Range("A59").Value = "Adjust to serial dilution"
    Range("B26:N26").Copy Range("B59:N59")
    Range("B60").Value = "1"
    Range("B61").Value = "2"
    Range("B62").Value = "4"
    Range("B63").Value = "8"
    Range("B64").Value = "16"
    Range("B65").Value = "32"
    Range("B66").Value = "64"
    Range("B67").Value = "128"
    Range("B60:B67").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(153, 204, 255)
    End With
    Range("C60:N67").Formula = "=R[-11]C*RC2"
    Range("C60:C67").Select
    Selection.AutoFill Destination:=Range("C60:N67"), Type:=xlFillDefault
        
    Dim cell As Range
            
' Loop through the target range
    For Each cell In Range("C60:N67")
        If IsNumeric(cell.Value) Then
            If cell.Value <= 0 Then
                cell.ClearContents
            End If
        End If
    Next cell
       
End Sub
Sub resultCalc_8()
'
' Averages the most consistent values for each sample excluding outliers using a derived value based on Tukey's fences method.
'

'
    Dim i As Integer
    For i = 2 To 12
        Cells(70, 2 + i).Value = i
    Next i
    
    Range("D70:N70").Interior.Color = RGB(153, 204, 255)
    Range("A71").Value = "Dilution selection and averaging"
    Range("C71").Value = "Q1"
    Range("D71").Select
    Selection.Formula = "=QUARTILE.INC(D60:D67,1)"
    Selection.AutoFill Destination:=Range("D71:N71"), Type:=xlFillDefault
    Range("C72").Value = "Q3"
    Range("D72").Select
    Selection.Formula = "=QUARTILE.INC(D60:D67,3)"
    Selection.AutoFill Destination:=Range("D72:N72"), Type:=xlFillDefault
    Range("C73").Value = "IQR"
    Range("D73").Select
    Selection.Formula = "=D72-D71"
    Selection.AutoFill Destination:=Range("D73:N73"), Type:=xlFillDefault
    Range("C74").Value = "lowBd"
    Range("D74").Select
    Selection.Formula = "=D71-0.5*D73"
    Selection.AutoFill Destination:=Range("D74:N74"), Type:=xlFillDefault
    Range("C75").Value = "upBd"
    Range("D75").Select
    Selection.Formula = "=D72+0.5*D73"
    Selection.AutoFill Destination:=Range("D75:N75"), Type:=xlFillDefault
    Range("C76").Value = "Averagelbub"
    Range("D76").Formula = "=AVERAGEIFS(D60:D67,D60:D67,"">=""&D74,D60:D67,""<=""&D75)"
    Range("D76").Select
    Selection.AutoFill Destination:=Range("D76:N76")
    

End Sub

Sub dfCalc_9()
'
' Multiplies the average sample value by the initial standard or sample dilution to obtain the sample concentration in ng/mL and divides all values by a 1000 to obtain the samples concentration in µg/mL.
'

'
    Range("A79").Value = "Adjust to sample dilution"
    Range("A81").Value = "Final concentration"
    Range("C26:N26").Copy Range("C79:N79")
    Range("B79").Value = "75"
    Range("B80").Value = "ng/mL"
    Range("B81").Value = "µg/mL"
    Range("B79:B81").HorizontalAlignment = xlCenter
    Range("D80").Select
    Selection.FormulaR1C1 = "=R[-4]C*10000"
    Range("E80").Select
    Selection.FormulaR1C1 = "=R[-4]C*R79C2"
    Selection.AutoFill Destination:=Range("E80:N80"), Type:=xlFillDefault
    Range("E80:N80").Select
    Range("D81").Select
    Selection.Formula = "=R[-1]C/1000"
    Selection.AutoFill Destination:=Range("D81:N81"), Type:=xlFillDefault

End Sub
Sub formatcells_10()
'
' Formats all cells with values.
'
Dim iRange As Range
Dim iCells As Range

Set iRange = ActiveSheet.UsedRange

For Each iCells In iRange
    If Not IsEmpty(iCells) Then
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    End If
    
Next iCells

Range("A:A, Q:S").EntireColumn.AutoFit
End Sub


