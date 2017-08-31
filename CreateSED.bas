Attribute VB_Name = "Module1"
'
' Visual Basic macro designed to create Spectral Energy Distribution (SED) plots for Young Stellar Objects.
' The macro CreateSED() sources a user data sheet, prompts for the selection of star IDs, and creates SEDs for each selected object.
'
' The set of variables below contains all of the relevant values used for the calculations.
' Should anything need to be adapted or updated, reasigning these values should be sufficient and easy.
'
' Author: J. McKernan
' Date: 10.13.16
' Version: 1.7
' Updates since 1.6: Added dynamic source sheet thanks to a Andrew A. Trendlines updated.
'                    Data column references updated to account for notes columns.
' Updates since 1.5: Updated error output
' Updates since 1.4: Added trendlines and class distinction.
'
Sub CreateSED()
Attribute CreateSED.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ' Reference data
    Dim sources() As String
    Dim baseSources() As String
    Dim lambdaUm() As Variant
    Dim columns() As String
    Dim catalogueNumberColumn As String
    sources = Split("iphas-r, iphas-i, iphas-ha, 2mass-J, 2mass-H, 2mass-K, irac-1, irac-2, irac-3, irac-4, mips24, wise-1, wise-2, wise-3, wise-4", ", ")
    baseSources = Split("IPHAS, 2MASS, IRAC, MIPS, WISE", ", ")
    lambdaUm = Array(0.624, 0.656, 0.774, 1.235, 1.662, 2.159, 3.6, 4.5, 5.8, 8, 24, 3.4, 4.6, 12, 22)
    ' !!!!Change columns here!!!!
    columns = Split("O, V, AC, AK, AS, BA, BI, BQ, BY, CG, CO, CW, DE, DM, DU", ", ")
    catalogueNumberColumn = "B"
    

    ' Graph formatting variables
    Dim defaultTargetBaseName As String
    Dim hasGridlines As Boolean
    Dim maxYScale As Double
    Dim minYScale As Double
    Dim maxXScale As Double
    Dim minXScale As Double
    Dim horizontalCrossPoint As Double
    defaultTargetBaseName = "Plot"
    hasGridlines = False
    addTrendline = True
    displayTrendlineEquationOnGraph = True ' Only applicable if addTrendline is True
    maxYScale = -10
    minYScale = -14
    maxXScale = 2.3
    minXScale = -0.8
    horizontalCrossPoint = minYScale
    verticalCrossPoint = minXScale
    
    
    ' Variable declaration
    Dim dataRowId As Integer
    Dim excelRowId As Integer
    Dim idRange As Range
    Dim dataSheetName As String
    Dim dataSheet As Worksheet
    Dim workingSheet As Worksheet
    Dim targetBaseName As String
    Dim address As String
    Dim strKMipsEquation As String
    Dim strKWiseEquation As String
    Dim trendlineSlope As Double
    Dim trendlineXRange As String
    Dim trendlineYRange As String
    Dim result As String
    Dim errors As String
    Dim errorReply As VbMsgBoxResult
    targetBaseName = ""
    
    ' Misc. variables
    Dim i As Integer
    Dim j As Integer
    Dim start As Integer
    Dim length As Integer
    Dim count As Integer
    i = 0
    j = 0

    
    ' Get data sheet name
sheetIn:
    dataSheetName = Application.InputBox("Enter the name of your data sheet (case sensitive):", "Data Sheet Name", "Data")
    
    If SheetExists(dataSheetName) Then
        Set dataSheet = Sheets(dataSheetName)
        dataSheet.Activate



rangeIn:
        Set idRange = Application.InputBox("Select the ID numbers (first column of data) of the stars you would like to plot:", Type:=8)

        ' Error handling for range input
        If idRange.Cells.count < 1 Then
            MsgBox ("No cells were selected. Please try again.")
            GoTo rangeIn
        Else
            For Each cell In Range(idRange.address)
                If IsEmpty(cell) Then
                    MsgBox ("One or more selected cells is empty. Please try again")
                    GoTo rangeIn
                End If
            Next
        End If
        
        
        
        i = 0
        count = Range(idRange.address).count
        For Each cell In Range(idRange.address)
            dataRowId = cell
            excelRowId = dataSheet.Range(cell.address).Row
            Call AddNewSheet(defaultTargetBaseName, targetBaseName, dataRowId) ' targetBaseName passed by reference and updated; reset to "Plot" each time
            Set workingSheet = Sheets(targetBaseName) ' Contingency for loss of ActiveSheet
            j = 0
            
            Application.ScreenUpdating = False    ' Prevents screen refreshing
            
            ' Column headers
            workingSheet.Range("A2:E2").HorizontalAlignment = xlCenter
            workingSheet.Range("A2").Value = "Band"
            workingSheet.Range("B2").Value = ChrW(955) & "(" & ChrW(181) & "m)"
            workingSheet.Range("C2").Value = ChrW(955) & "(cm)"
            workingSheet.Range("D2").Value = "log" & ChrW(955)
            workingSheet.Range("E2").Value = "log" & ChrW(955) & "f" & ChrW(955)

            ' Data sheet printing
            For Each src In sources
                workingSheet.Range("A" & CStr(j + 3)).Value = src
                workingSheet.Range("B" & CStr(j + 3)).Value = lambdaUm(j)
                workingSheet.Range("C" & CStr(j + 3)).Value = "=B" & CStr(j + 3) & "*10^-4"
                workingSheet.Range("D" & CStr(j + 3)).Value = "=log(B" & CStr(j + 3) & ")"
                address = columns(j) & CStr(excelRowId)
                workingSheet.Range("E" & CStr(j + 3)).Value = "=IF('" & dataSheetName & "'!" & address & "<>"""", '" & dataSheetName & "'!" & address & ", NA()) "
                j = j + 1
            Next
            
            ' Add Chart
            Range("D2:E17").Select
            Charts.Add
            With ActiveChart
                .ChartType = xlXYScatter
                .Location Where:=xlLocationAsObject, name:=targetBaseName
            End With
            
            workingSheet.Activate
            
            ActiveSheet.ChartObjects("Chart 1").Activate
            ActiveChart.SetSourceData Source:=Range("$D$2:$E$17") ' Possibly redundant
            ActiveSheet.Shapes("Chart 1").Left = Range("H7").Left
            ActiveSheet.Shapes("Chart 1").Top = Range("H7").Top
            ActiveChart.Axes(xlValue).HasMajorGridlines = hasGridlines
            ActiveChart.Axes(xlCategory).HasMajorGridlines = hasGridlines
            ActiveChart.Axes(xlValue).HasMinorGridlines = False
            ActiveChart.Axes(xlCategory).HasMinorGridlines = False
            ActiveChart.Axes(xlValue).TickLabelPosition = xlTickLabelPositionLow
            ActiveChart.Axes(xlCategory).TickLabelPosition = xlTickLabelPositionLow
            ActiveChart.Axes(xlValue).CrossesAt = horizontalCrossPoint
            ActiveChart.Axes(xlCategory).CrossesAt = verticalCrossPoint
            ActiveChart.Axes(xlValue).MaximumScale = maxYScale
            ActiveChart.Axes(xlValue).MinimumScale = minYScale
            ActiveChart.Axes(xlCategory).MaximumScale = maxXScale
            ActiveChart.Axes(xlCategory).MinimumScale = minXScale
            ActiveChart.Axes(xlValue).ReversePlotOrder = False
            ActiveChart.Axes(xlCategory).ReversePlotOrder = False
            ActiveChart.ChartTitle.Select
            Selection.Caption = "='" & dataSheetName & "'!$" & catalogueNumberColumn & "$" & CStr(excelRowId)
        
            ActiveChart.SeriesCollection(1).Delete
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(1).name = "=""" & baseSources(0) & """"
            ActiveChart.SeriesCollection(1).XValues = "='" & targetBaseName & "'!$D$3:$D$5"
            ActiveChart.SeriesCollection(1).Values = "='" & targetBaseName & "'!$E$3:$E$5"
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(2).name = "=""" & baseSources(1) & """"
            ActiveChart.SeriesCollection(2).XValues = "='" & targetBaseName & "'!$D$6:$D$8"
            ActiveChart.SeriesCollection(2).Values = "='" & targetBaseName & "'!$E$6:$E$8"
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(3).name = "=""" & baseSources(2) & """"
            ActiveChart.SeriesCollection(3).XValues = "='" & targetBaseName & "'!$D$9:$D$12"
            ActiveChart.SeriesCollection(3).Values = "='" & targetBaseName & "'!$E$9:$E$12"
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(4).name = "=""" & baseSources(3) & """"
            ActiveChart.SeriesCollection(4).XValues = "='" & targetBaseName & "'!$D$13"
            ActiveChart.SeriesCollection(4).Values = "='" & targetBaseName & "'!$E$13"
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(5).name = "=""" & baseSources(4) & """"
            ActiveChart.SeriesCollection(5).XValues = "='" & targetBaseName & "'!$D$14:$D$17"
            ActiveChart.SeriesCollection(5).Values = "='" & targetBaseName & "'!$E$14:$E$17"
            ' Trendline series K-MIPS
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(6).name = vbNullString
            ActiveChart.SeriesCollection(6).XValues = "='" & targetBaseName & "'!$D$8:$D$13"
            ActiveChart.SeriesCollection(6).Values = "='" & targetBaseName & "'!$E$8:$E$13"
            ActiveChart.SeriesCollection(6).MarkerStyle = xlMarkerStyleNone
            ActiveChart.Legend.LegendEntries(6).Delete ' Delete legend entry for combined trend series
            ' Trendline series K-WISE
            ActiveChart.SeriesCollection.NewSeries
            ActiveChart.SeriesCollection(7).name = vbNullString
            ActiveChart.SeriesCollection(7).XValues = "=('" & targetBaseName & "'!$D$8,'" & targetBaseName & "'!$D$14,'" & targetBaseName & "'!$D$15,'" & targetBaseName & "'!$D$16,'" & targetBaseName & "'!$D$17)"
            ActiveChart.SeriesCollection(7).Values = "=('" & targetBaseName & "'!$E$8,'" & targetBaseName & "'!$E$14,'" & targetBaseName & "'!$E$15,'" & targetBaseName & "'!$E$16,'" & targetBaseName & "'!$E$17)"
            ActiveChart.SeriesCollection(7).MarkerStyle = xlMarkerStyleNone
            ActiveChart.Legend.LegendEntries(6).Delete ' Delete legend entry for combined trend series
            ActiveChart.HasLegend = True
        
            ActiveSheet.Shapes("Chart 1").ScaleWidth 1.0906248906, msoFalse, msoScaleFromBottomRight
            ActiveSheet.Shapes("Chart 1").ScaleHeight 1.3767359288, msoFalse, msoScaleFromBottomRight
            ActiveSheet.ChartObjects("Chart 1").Activate
            ActiveSheet.Shapes("Chart 1").ScaleWidth 1.3352436869, msoFalse, msoScaleFromTopLeft
            ActiveSheet.Shapes("Chart 1").ScaleHeight 1.0025220684, msoFalse, msoScaleFromBottomRight
            
            If addTrendline Then
                ' Trendline for K through MIPS data
                ActiveChart.SeriesCollection(6).Trendlines.Add
                ActiveChart.SeriesCollection(6).Trendlines(1).name = "Linear Trend (K-MIPS)"
                ActiveChart.SeriesCollection(6).Trendlines(1).Type = xlLinear
                ActiveChart.SeriesCollection(6).Trendlines(1).DisplayEquation = True ' Show trendline in order to grab the value
                'ActiveChart.SeriesCollection(6).Trendlines(1).DisplayEquation.AutoText = True
                
                ' Trendline for K through WISE data
                ' conditional if all of irac is gone ?
                ' errors = errors & "No IRAC data was present on """ & workingSheet.name & """." & vbNewLine & "WISE data has been used to calculate this trendline." & vbNewLine & vbNewLine
                ' ActiveSheet.Range("C21").Font.Bold = True
                ' Range("C21") = "No IRAC data present."
                ActiveChart.SeriesCollection(7).Trendlines.Add
                ActiveChart.SeriesCollection(7).Trendlines(1).name = "Linear Trend (K-WISE)"
                ActiveChart.SeriesCollection(7).Trendlines(1).Type = xlLinear
                ActiveChart.SeriesCollection(7).Trendlines(1).Border.LineStyle = xlDash
                ActiveChart.SeriesCollection(7).Trendlines(1).DisplayEquation = True ' Show trendline in order to grab the value
                'ActiveChart.SeriesCollection(7).Trendlines(1).DisplayEquation.AutoText = True
                
                Application.ScreenUpdating = True     ' Could not grab data label with screen unupdated
                
                ' Store trendline equations, if present
                ActiveChart.SeriesCollection(6).Trendlines(1).Select
                strKMipsEquation = ActiveChart.SeriesCollection(6).Trendlines(1).DataLabel.Formula
                ActiveChart.SeriesCollection(7).Trendlines(1).Select
                strKWiseEquation = ActiveChart.SeriesCollection(7).Trendlines(1).DataLabel.Formula
                
                Application.ScreenUpdating = False
                
                ' Slope/Class labels
                ActiveSheet.Range("B20:D23").Font.Bold = True
                ActiveSheet.Range("B20:D23").HorizontalAlignment = xlRight
                Range("B20") = "K-MIPS - "
                Range("C20") = "Slope:"
                Range("C21") = "Class:"
                Range("B22") = "K-WISE - "
                Range("C22") = "Slope:"
                Range("C23") = "Class:"
                
                ' Slope extraction and class calculation
                ' K-MIPS
                If Range("E8:E13").count - CountError(Range("E8:E13")) > 1 Then
                    start = InStr(strKMipsEquation, "=") + 1
                    length = InStr(strKMipsEquation, "x") - start
                    Range("D20") = Mid(strKMipsEquation, start, length)
                    trendlineSlope = CDbl(Mid(strKMipsEquation, start, length))
                    
                    Select Case trendlineSlope
                        Case Is < -1.6
                            result = "III"
                        Case Is < -0.3
                            result = "II"
                        Case Is < 0.3
                            result = "Flat"
                        Case Is > 0.3
                            result = "I"
                        Case Else
                            result = "Slope outside class ranges"
                    End Select
                    Range("D21") = result ' K-MIPS class
                End If
                
                ' K-WISE
                If (Range("E8").count + Range("E14:E17").count) - (CountError(Range("E8")) + CountError(Range("E14:E17"))) > 1 Then
                    start = InStr(strKWiseEquation, "=") + 1
                    length = InStr(strKWiseEquation, "x") - start
                    Range("D22") = Mid(strKWiseEquation, start, length)
                    trendlineSlope = CDbl(Mid(strKWiseEquation, start, length))
                    
                    Select Case trendlineSlope
                        Case Is < -1.6
                            result = "III"
                        Case Is < -0.3
                            result = "II"
                        Case Is < 0.3
                            result = "Flat"
                        Case Is > 0.3
                            result = "I"
                        Case Else
                            result = "Slope outside class ranges"
                    End Select
                    Range("D23") = result ' K-WISE class
                End If
                
                ' Recheck if trendline on graph should be hidden
                ActiveChart.SeriesCollection(6).Trendlines(1).DisplayEquation = displayTrendlineEquationOnGraph
                ActiveChart.SeriesCollection(7).Trendlines(1).DisplayEquation = displayTrendlineEquationOnGraph
                
            End If
            
            Application.ScreenUpdating = True     ' Update screen
            
            ' Status update
            Application.StatusBar = (i + 1) & " of " & count & " plots completed."
            
            i = i + 1
        Next
        
        
        
        ' Completion
        Application.ScreenUpdating = True     ' Update screen
        If errors <> "" Then
            errors = errors & vbNewLine & vbNewLine & " -- End of Error List -- "
            errorReply = MsgBox(errors, vbOKCancel, "Error List")
        End If
        MsgBox ("Complete.")
        Application.StatusBar = False
        
        
    
    ' Error handling for sheet name
    Else
        If dataSheetName = "False" Then
            Exit Sub
        Else
            MsgBox ("Error: The sheet " & dataSheetName & " does not exist.")
            GoTo sheetIn
        End If
    End If
    

End Sub

Function SheetExists(sheetName As String, Optional wb As Workbook) As Boolean

    Dim sheet As Worksheet

     If wb Is Nothing Then Set wb = ActiveWorkbook
     On Error Resume Next
     Set sheet = wb.Sheets(sheetName)
     On Error GoTo 0
     SheetExists = Not sheet Is Nothing
     
 End Function
 
Sub AddNewSheet(ByVal default As String, ByRef name As String, number As Integer)
       
    name = default ' Reset to "Plot"
       
    If number <> 0 Then
        name = name & CStr(number)
    End If
    
    If SheetExists(name) Then
        name = name & "(Duplicate)"
    End If
    
    Application.ScreenUpdating = False    ' Prevents screen refreshing
    Worksheets.Add
    ActiveSheet.Move After:=Sheets(ActiveWorkbook.Sheets.count) ' Move sheet to last
    ActiveSheet.name = name
    Application.ScreenUpdating = True     ' Enables screen refreshing
    
End Sub

Function CountError(MyRg As Range) As Integer
     '
    Dim Rg   As Range
    Dim C As Variant
    Set Rg = MyRg
    CountError = 0
    For Each C In Rg
        If (IsError(C.Value)) Then CountError = CountError + 1
    Next C
     '
End Function
