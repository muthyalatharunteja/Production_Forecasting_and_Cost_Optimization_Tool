Attribute VB_Name = "Module1"
Sub GenerateProductionForecast()
    Dim wsInputs As Worksheet, wsCalendar As Worksheet, wsForecast As Worksheet, wsCost As Worksheet
    Dim batchSize As Integer, stepsPerOrder As Integer, weekInterval As Integer, maxClientsPerMonth As Integer
    Dim costPerStep As Double, handlingCost As Double, bonusProfit As Double

    ' Read Inputs
    Set wsInputs = ThisWorkbook.Sheets("Inputs") 'set used _when assigning an object & 'This workbook current working vba workbook
    batchSize = wsInputs.Range("B2").Value
    stepsPerOrder = wsInputs.Range("B3").Value
    weekInterval = wsInputs.Range("B4").Value
    maxClientsPerMonth = wsInputs.Range("B5").Value
    costPerStep = wsInputs.Range("B6").Value
    handlingCost = wsInputs.Range("B7").Value
    bonusProfit = wsInputs.Range("B8").Value

    ' Remove old sheets
    On Error Resume Next 'allows code to continue if an error occurs
    Application.DisplayAlerts = False  'error alert message off
    Sheets("Calendar").Delete
    Sheets("Forecast").Delete
    Sheets("Cost Summary").Delete
    Sheets("Dashboard").Delete
    Application.DisplayAlerts = True 'error alert message on
    On Error GoTo 0 'reset as before

    ' Create new sheets
    Set wsCalendar = Sheets.Add(After:=wsInputs): wsCalendar.Name = "Calendar" 'ws calendar wll be as sheet1 andthen it renamed to calendar
    Set wsForecast = Sheets.Add(After:=wsCalendar): wsForecast.Name = "Forecast"
    Set wsCost = Sheets.Add(After:=wsForecast): wsCost.Name = "Cost Summary"

    ' Headers
    'assign calendar sheet
    wsCalendar.Range("A1:E1").Value = Array("Client", "Batch", "Cutting Date", "Assembly Date", "Finishing Date") 'array only 1D row, beacuse _any data might come so we fix it as_ variant
    wsForecast.Range("A1:C1").Value = Array("Month", "Steps Needed", "Clients Completed")
    wsCost.Range("A1:F1").Value = Array("Month", "Production Cost", "Handling Cost", "Cumulative Clients", "Bonus Triggered", "Total Cost")

    ' Initialize i.e start with these
    Dim startDate As Date: startDate = DateSerial(2025, 7, 1)
    Dim clientsScheduled As Integer: clientsScheduled = 0
    Dim clientID As Integer: clientID = 1
    Dim cumulativeClients As Integer: cumulativeClients = 0 ' adding up overtime after it getting processsed
    Dim bonusGiven As Boolean: bonusGiven = False  'to avoid duplicates i.e non repeating

    Dim monthList() As String, stepsList() As Integer, clientsList() As Integer
    'monthList() Ð holds month name
    'stepsList() Ð holds number of steps scheduled for each month
    'clientsList() Ð holds how many clients finished all steps in that month
    Dim monthCount As Integer: monthCount = 0

    ' 1.Start scheduling in table
    Dim batchNum As Integer: batchNum = 1
    Do While clientsScheduled < 18
        Dim i As Integer
        For i = 1 To batchSize
            If clientsScheduled >= 18 Then Exit For

            Dim stepDates(1 To 3) As Date
            stepDates(1) = DateAdd("ww", (batchNum - 1) * stepsPerOrder * weekInterval, startDate) '(" ",0*3*2 =0,01-07-2025) start date
            stepDates(2) = DateAdd("ww", weekInterval, stepDates(1)) '(2 weeks+ from starting date(1))
            stepDates(3) = DateAdd("ww", weekInterval, stepDates(2)) '(2 weeks+ from starting date(2))

            Dim j As Integer
            For j = 1 To 3 'j=1:Cutting, J=2:Assembly;j=3:Finishing
                Dim monthKey As String
                monthKey = Format(stepDates(j), "mmm-yyyy")
                Call UpdateMonthData(monthList, stepsList, clientsList, monthCount, monthKey, IIf(j = 3, 1, 0))
                'if j=3rd step i.e finishing  then 1 if itÕs the last step)_> count client as "completed"0 otherwise
                '=======1st step end
                
            Next j '__it repeats for all 3 steps(i.e update and subupdate)

           '====4th step begins
           ' Write calendar with actual dates
            wsCalendar.Cells(clientID + 1, 1).Value = "Client " & clientID
            wsCalendar.Cells(clientID + 1, 2).Value = batchNum
            wsCalendar.Cells(clientID + 1, 3).Value = Format(stepDates(1), "dd.mm.yyyy") ' Cutting
            wsCalendar.Cells(clientID + 1, 4).Value = Format(stepDates(2), "dd.mm.yyyy") ' Assembly
            wsCalendar.Cells(clientID + 1, 5).Value = Format(stepDates(3), "dd.mm.yyyy") ' Finishing

            clientsScheduled = clientsScheduled + 1
            clientID = clientID + 1
        Next i '(( ==end for one client and for next cleinrt it goes to step 1))
        batchNum = batchNum + 1
    Loop

    ' Output Forecast and Cost Summary
    Dim row As Integer: row = 2
    Dim totalClients As Integer: totalClients = 0

    Dim k As Integer
    For k = 0 To monthCount - 1 '(4_months, jul-oct, as per calendar)
        totalClients = totalClients + clientsList(k)  'clients(k) who recieved 3 step in same month
        Dim bonus As Double: bonus = 0
        If Not bonusGiven And totalClients >= 10 Then
            bonus = bonusProfit
            bonusGiven = True
        End If

        wsForecast.Cells(row, 1).Value = monthList(k)
        wsForecast.Cells(row, 2).Value = stepsList(k)
        wsForecast.Cells(row, 3).Value = totalClients

        Dim prodCost As Double, handleCost As Double
        prodCost = stepsList(k) * costPerStep 'no of steps in a month*cost per step(2000)
        handleCost = clientsList(k) * handlingCost '(1500 handling cost)

        wsCost.Cells(row, 1).Value = monthList(k)
        wsCost.Cells(row, 2).Value = prodCost
        wsCost.Cells(row, 3).Value = handleCost
        wsCost.Cells(row, 4).Value = totalClients
        wsCost.Cells(row, 5).Value = IIf(bonus > 0, "YES", "")
        wsCost.Cells(row, 6).Value = prodCost + handleCost - bonus '(bonus is not extra cost)

        row = row + 1
    Next k

    ' Dashboard Chart
    Dim chartSheet As Worksheet
    Set chartSheet = Sheets.Add(After:=wsCost): chartSheet.Name = "Dashboard"
    Dim chartObj As ChartObject
    Set chartObj = chartSheet.ChartObjects.Add(Left:=100, Width:=500, Top:=50, Height:=300)
    chartObj.Chart.ChartType = xlColumnClustered
    chartObj.Chart.SetSourceData Source:=wsForecast.Range("A1:B" & wsForecast.Cells(wsForecast.Rows.count, 1).End(xlUp).row)
    chartObj.Chart.HasTitle = True
    chartObj.Chart.ChartTitle.Text = "Monthly Production Steps"

    MsgBox "Forecast and cost analysis generated successfully!", vbInformation
End Sub
'======2nd step (to check/Refer, whether month step or clients exists)
Sub UpdateMonthData(ByRef months() As String, ByRef steps() As Integer, ByRef clients() As Integer, ByRef count As Integer, month As String, Optional clientInc As Integer = 0)
    Dim i As Integer
    For i = 0 To count - 1 'arrays are 0 by default i.e when count=3 (it counts from 0,1,2)
        If months(i) = month Then 'if current month = current month
            steps(i) = steps(i) + 1 'Increases the number of steps scheduled in this month by 1.
            clients(i) = clients(i) + clientInc 'clientInc is either 1 or 0:
            '1 if the current step is step 3 (Finishing) -->Client is completed
            '0 for steps 1 and 2 ? Not completed yet
            Exit Sub
        End If ' step_2ends_ if only 3 steps finished
        
        '=====3rd step begins
    Next i
    ReDim Preserve months(0 To count)
    ReDim Preserve steps(0 To count)
    ReDim Preserve clients(0 To count)
    months(count) = month
    steps(count) = 1
    clients(count) = clientInc
    count = count + 1
End Sub '==  3rd step ends __because it preserves/store new value and create a new array for upcoming data


