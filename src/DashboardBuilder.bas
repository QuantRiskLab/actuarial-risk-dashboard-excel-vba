
Attribute VB_Name = "DashboardBuilder"
Option Explicit

' Actuarial Risk Dashboard
' Automated Excel Dashboard Builder

Public Sub BuildActuarialDashboard()

    Dim ws As Worksheet
    Dim wsPricing As Worksheet
    Dim wsStress As Worksheet
    Dim wsSolvency As Worksheet

    Dim chartObject As ChartObject
    Dim chartItem As ChartObject
    Dim dashboardButton As Button

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set wsPricing = ThisWorkbook.Worksheets("Pricing")
    Set wsStress = ThisWorkbook.Worksheets("Stress_Testing")
    Set wsSolvency = ThisWorkbook.Worksheets("Solvency_II")

    If wsPricing.Range("I3").Value <= 0 Or _
       wsStress.Range("A2").Value <> "Baseline" Or _
       wsSolvency.Range("B13").Value <= 0 Then

        MsgBox "Run portfolio generation, claims simulation, " & _
               "pricing, stress testing and capital analysis first.", _
               vbExclamation, "Dashboard Builder"
        Exit Sub

    End If

    ' Reset dashboard
    For Each chartItem In ws.ChartObjects
        chartItem.Delete
    Next chartItem

    For Each dashboardButton In ws.Buttons
        dashboardButton.Delete
    Next dashboardButton

    ws.Cells.Clear

    ' Dashboard title
    With ws.Range("B2:L3")
        .Merge
        .Value = "ACTUARIAL RISK DASHBOARD"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Key performance indicators
    ws.Range("B5").Value = "Total Premium"
    ws.Range("D5").Value = "Total Claims"
    ws.Range("F5").Value = "Loss Ratio"
    ws.Range("H5").Value = "Capital Requirement"
    ws.Range("J5").Value = "Coverage Ratio"

    ws.Range("B6").Value = wsPricing.Range("I3").Value
    ws.Range("D6").Value = wsPricing.Range("I5").Value
    ws.Range("F6").Value = wsPricing.Range("I7").Value
    ws.Range("H6").Value = wsSolvency.Range("B13").Value
    ws.Range("J6").Value = wsSolvency.Range("B14").Value

    With ws.Range("B5:K5")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    With ws.Range("B6:K6")
        .Font.Size = 14
        .Font.Bold = True
    End With

    ws.Range("B6").NumberFormat = "#,##0.00"
    ws.Range("D6").NumberFormat = "#,##0.00"
    ws.Range("F6").NumberFormat = "0.00%"
    ws.Range("H6").NumberFormat = "#,##0.00"
    ws.Range("J6").NumberFormat = "0.00%"

    ' Stress scenario chart
    Set chartObject = ws.ChartObjects.Add( _
        Left:=ws.Range("B9").Left, _
        Top:=ws.Range("B9").Top, _
        Width:=550, _
        Height:=300 _
    )

    With chartObject.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=wsStress.Range("A1:A6,D1:D6")
        .HasTitle = True
        .ChartTitle.Text = "Claims Under Stress Scenarios"
        .HasLegend = False
    End With

    ' Capital components chart
    Set chartObject = ws.ChartObjects.Add( _
        Left:=ws.Range("B27").Left, _
        Top:=ws.Range("B27").Top, _
        Width:=550, _
        Height:=300 _
    )

    With chartObject.Chart
        .ChartType = xlBarClustered
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Risk Capital"
        .SeriesCollection(1).XValues = _
            wsSolvency.Range("A7:A9")
        .SeriesCollection(1).Values = _
            wsSolvency.Range("B7:B9")
        .HasTitle = True
        .ChartTitle.Text = "Insurance Risk Capital Components"
        .HasLegend = False
    End With

    ' Refresh button
    Set dashboardButton = ws.Buttons.Add( _
        ws.Range("J9").Left, _
        ws.Range("J9").Top, _
        150, _
        35 _
    )

    With dashboardButton
        .Caption = "Refresh Dashboard"
        .OnAction = "BuildActuarialDashboard"
    End With

    ' Formatting
    ws.Columns("A:L").ColumnWidth = 14
    ws.Rows("2:3").RowHeight = 25

    ws.Range("B46").Value = _
        "Synthetic data | Educational model | Not regulatory SCR"

    ws.Range("B46").Font.Italic = True
    ws.Range("B46").Font.Color = RGB(100, 100, 100)

    ws.Activate

    MsgBox "Actuarial dashboard created successfully.", _
           vbInformation, "Dashboard Builder"

End Sub
