
Attribute VB_Name = "SolvencyCapital"
Option Explicit

' Actuarial Risk Dashboard
' Illustrative Solvency Capital Model
'
' This is a simplified educational model.
' It does not implement the regulatory Solvency II SCR formula.

Public Sub CalculateSolvencyCapital()

    Dim wsPricing As Worksheet
    Dim wsStress As Worksheet
    Dim wsSolvency As Worksheet

    Dim totalPremium As Double
    Dim expectedClaims As Double
    Dim availableCapital As Double

    Dim premiumRiskCapital As Double
    Dim reserveRiskCapital As Double
    Dim catastropheRiskCapital As Double
    Dim underwritingCapital As Double
    Dim operationalCapital As Double
    Dim totalCapitalRequirement As Double

    Dim premiumRiskFactor As Double
    Dim reserveRiskFactor As Double
    Dim operationalRiskFactor As Double
    Dim correlation As Double

    Dim worstStressLoss As Double
    Dim stressLoss As Double
    Dim baselineClaims As Double

    Dim solvencyRatio As Double
    Dim i As Long
    Dim lastRow As Long

    Set wsPricing = ThisWorkbook.Worksheets("Pricing")
    Set wsStress = ThisWorkbook.Worksheets("Stress_Testing")
    Set wsSolvency = ThisWorkbook.Worksheets("Solvency_II")

    totalPremium = wsPricing.Range("I3").Value
    expectedClaims = wsPricing.Range("I4").Value

    If totalPremium <= 0 Or expectedClaims < 0 Then
        MsgBox "Run pricing analysis first.", _
               vbExclamation, "Solvency Capital"
        Exit Sub
    End If

    If wsStress.Range("A2").Value <> "Baseline" Then
        MsgBox "Run stress testing first.", _
               vbExclamation, "Solvency Capital"
        Exit Sub
    End If

    ' Illustrative assumptions
    premiumRiskFactor = 0.15
    reserveRiskFactor = 0.10
    operationalRiskFactor = 0.03
    correlation = 0.25

    ' Hypothetical capital available to the insurer
    availableCapital = totalPremium * 0.6

    ' Simplified risk capital components
    premiumRiskCapital = totalPremium * premiumRiskFactor
    reserveRiskCapital = expectedClaims * reserveRiskFactor

    ' Catastrophe capital from additional stressed losses
    baselineClaims = wsStress.Range("D2").Value
    worstStressLoss = 0#

    lastRow = wsStress.Cells( _
        wsStress.Rows.Count, 1).End(xlUp).Row

    For i = 3 To lastRow

        stressLoss = wsStress.Cells(i, 4).Value - baselineClaims

        If stressLoss > worstStressLoss Then
            worstStressLoss = stressLoss
        End If

    Next i

    catastropheRiskCapital = worstStressLoss

    ' Illustrative correlated aggregation
    underwritingCapital = Sqr( _
        premiumRiskCapital ^ 2 + _
        reserveRiskCapital ^ 2 + _
        2# * correlation * premiumRiskCapital * reserveRiskCapital _
    ) + catastropheRiskCapital

    operationalCapital = totalPremium * operationalRiskFactor

    totalCapitalRequirement = _
        underwritingCapital + operationalCapital

    If totalCapitalRequirement > 0 Then
        solvencyRatio = availableCapital / totalCapitalRequirement
    Else
        solvencyRatio = 0#
    End If

    ' Write results
    wsSolvency.Cells.Clear

    wsSolvency.Range("A1").Value = _
        "Illustrative Solvency Capital Analysis"

    wsSolvency.Range("A3").Value = "Total Premium"
    wsSolvency.Range("B3").Value = totalPremium

    wsSolvency.Range("A4").Value = "Expected Claims"
    wsSolvency.Range("B4").Value = expectedClaims

    wsSolvency.Range("A5").Value = "Available Capital (Assumed)"
    wsSolvency.Range("B5").Value = availableCapital

    wsSolvency.Range("A7").Value = "Premium Risk Capital"
    wsSolvency.Range("B7").Value = premiumRiskCapital

    wsSolvency.Range("A8").Value = "Reserve Risk Capital"
    wsSolvency.Range("B8").Value = reserveRiskCapital

    wsSolvency.Range("A9").Value = "Catastrophe Risk Capital"
    wsSolvency.Range("B9").Value = catastropheRiskCapital

    wsSolvency.Range("A10").Value = "Underwriting Risk Capital"
    wsSolvency.Range("B10").Value = underwritingCapital

    wsSolvency.Range("A11").Value = "Operational Risk Capital"
    wsSolvency.Range("B11").Value = operationalCapital

    wsSolvency.Range("A13").Value = _
        "Total Illustrative Capital Requirement"
    wsSolvency.Range("B13").Value = totalCapitalRequirement

    wsSolvency.Range("A14").Value = "Illustrative Coverage Ratio"
    wsSolvency.Range("B14").Value = solvencyRatio

    wsSolvency.Range("A16").Value = _
        "Educational model - not a regulatory Solvency II SCR."

    ' Formatting
    wsSolvency.Range("A1:B1").Font.Bold = True
    wsSolvency.Range("A1:B1").Interior.Color = RGB(31, 78, 121)
    wsSolvency.Range("A1:B1").Font.Color = RGB(255, 255, 255)

    wsSolvency.Range("A13:B14").Font.Bold = True

    wsSolvency.Range("B3:B13").NumberFormat = "#,##0.00"
    wsSolvency.Range("B14").NumberFormat = "0.00%"

    wsSolvency.Columns("A:B").AutoFit

    MsgBox "Illustrative capital analysis completed.", _
           vbInformation, "Solvency Capital"

End Sub
