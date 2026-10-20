
Attribute VB_Name = "StressTesting"
Option Explicit

' Actuarial Risk Dashboard
' Insurance Stress Testing

Public Sub RunStressTesting()

    Dim wsPricing As Worksheet
    Dim wsStress As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim scenarioIndex As Long

    Dim totalPremium As Double
    Dim totalBaseClaims As Double
    Dim stressedClaims As Double
    Dim stressedLossRatio As Double
    Dim underwritingResult As Double

    Dim inflationFactor As Double
    Dim catastropheLoss As Double

    Dim scenarioNames As Variant
    Dim inflationFactors As Variant
    Dim catastropheLosses As Variant

    Set wsPricing = ThisWorkbook.Worksheets("Pricing")
    Set wsStress = ThisWorkbook.Worksheets("Stress_Testing")

    lastRow = wsPricing.Cells( _
        wsPricing.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Run pricing analysis first.", _
               vbExclamation, "Stress Testing"
        Exit Sub
    End If

    totalPremium = 0#
    totalBaseClaims = 0#

    For i = 2 To lastRow
        totalPremium = totalPremium + wsPricing.Cells(i, 3).Value
        totalBaseClaims = totalBaseClaims + wsPricing.Cells(i, 5).Value
    Next i

    If totalPremium <= 0 Then
        MsgBox "Total premium must be positive.", _
               vbExclamation, "Stress Testing"
        Exit Sub
    End If

    scenarioNames = Array( _
        "Baseline", _
        "Moderate Inflation", _
        "Severe Inflation", _
        "Catastrophe Event", _
        "Combined Stress" _
    )

    inflationFactors = Array( _
        1#, _
        1.1, _
        1.25, _
        1#, _
        1.25 _
    )

    catastropheLosses = Array( _
        0#, _
        0#, _
        0#, _
        500000#, _
        1000000# _
    )

    wsStress.Cells.Clear

    wsStress.Range("A1:F1").Value = Array( _
        "Scenario", _
        "Inflation_Factor", _
        "Catastrophe_Loss", _
        "Total_Claims", _
        "Loss_Ratio", _
        "Underwriting_Result" _
    )

    For scenarioIndex = 0 To UBound(scenarioNames)

        inflationFactor = inflationFactors(scenarioIndex)
        catastropheLoss = catastropheLosses(scenarioIndex)

        stressedClaims = _
            totalBaseClaims * inflationFactor + catastropheLoss

        stressedLossRatio = stressedClaims / totalPremium
        underwritingResult = totalPremium - stressedClaims

        wsStress.Cells(scenarioIndex + 2, 1).Value = _
            scenarioNames(scenarioIndex)

        wsStress.Cells(scenarioIndex + 2, 2).Value = _
            inflationFactor

        wsStress.Cells(scenarioIndex + 2, 3).Value = _
            catastropheLoss

        wsStress.Cells(scenarioIndex + 2, 4).Value = _
            stressedClaims

        wsStress.Cells(scenarioIndex + 2, 5).Value = _
            stressedLossRatio

        wsStress.Cells(scenarioIndex + 2, 6).Value = _
            underwritingResult

    Next scenarioIndex

    ' Portfolio summary
    wsStress.Range("H1").Value = "Stress Testing Summary"

    wsStress.Range("H3").Value = "Total Premium"
    wsStress.Range("I3").Value = totalPremium

    wsStress.Range("H4").Value = "Baseline Claims"
    wsStress.Range("I4").Value = totalBaseClaims

    wsStress.Range("H5").Value = "Baseline Loss Ratio"
    wsStress.Range("I5").Value = totalBaseClaims / totalPremium

    ' Formatting
    wsStress.Range("A1:F1").Font.Bold = True
    wsStress.Range("A1:F1").Interior.Color = RGB(31, 78, 121)
    wsStress.Range("A1:F1").Font.Color = RGB(255, 255, 255)

    wsStress.Range("H1:I1").Font.Bold = True
    wsStress.Range("H1:I1").Interior.Color = RGB(31, 78, 121)
    wsStress.Range("H1:I1").Font.Color = RGB(255, 255, 255)

    wsStress.Range("B2:B6").NumberFormat = "0.00"
    wsStress.Range("C2:D6").NumberFormat = "#,##0.00"
    wsStress.Range("E2:E6").NumberFormat = "0.00%"
    wsStress.Range("F2:F6").NumberFormat = "#,##0.00"

    wsStress.Range("I3:I4").NumberFormat = "#,##0.00"
    wsStress.Range("I5").NumberFormat = "0.00%"

    wsStress.Columns("A:I").AutoFit

    MsgBox "Insurance stress testing completed.", _
           vbInformation, "Stress Testing"

End Sub
