
Attribute VB_Name = "PricingAnalysis"
Option Explicit

' Actuarial Risk Dashboard
' Insurance Pricing and Loss Ratio Analysis

Public Sub AnalyzeInsurancePricing()

    Dim wsPortfolio As Worksheet
    Dim wsClaims As Worksheet
    Dim wsPricing As Worksheet

    Dim lastRow As Long
    Dim claimsLastRow As Long
    Dim i As Long

    Dim totalPremium As Double
    Dim totalExpectedLoss As Double
    Dim totalActualClaims As Double

    Dim premium As Double
    Dim expectedLoss As Double
    Dim actualClaims As Double

    Dim expectedLossRatio As Double
    Dim actualLossRatio As Double
    Dim underwritingResult As Double

    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsClaims = ThisWorkbook.Worksheets("Claims")
    Set wsPricing = ThisWorkbook.Worksheets("Pricing")

    lastRow = wsPortfolio.Cells( _
        wsPortfolio.Rows.Count, 1).End(xlUp).Row

    claimsLastRow = wsClaims.Cells( _
        wsClaims.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Or claimsLastRow <> lastRow Then
        MsgBox "Generate the portfolio and simulate claims first.", _
               vbExclamation, "Pricing Analysis"
        Exit Sub
    End If

    wsPricing.Cells.Clear

    wsPricing.Range("A1:F1").Value = Array( _
        "Policy_ID", _
        "Policy_Type", _
        "Annual_Premium", _
        "Expected_Loss", _
        "Actual_Claims", _
        "Actual_Loss_Ratio" _
    )

    totalPremium = 0#
    totalExpectedLoss = 0#
    totalActualClaims = 0#

    For i = 2 To lastRow

        If wsPortfolio.Cells(i, 1).Value <> _
           wsClaims.Cells(i, 1).Value Then

            MsgBox "Policy IDs do not match. Simulate claims again.", _
                   vbExclamation, "Pricing Analysis"
            Exit Sub

        End If

        premium = wsPortfolio.Cells(i, 8).Value
        expectedLoss = wsPortfolio.Cells(i, 7).Value
        actualClaims = wsClaims.Cells(i, 6).Value

        wsPricing.Cells(i, 1).Value = _
            wsPortfolio.Cells(i, 1).Value

        wsPricing.Cells(i, 2).Value = _
            wsPortfolio.Cells(i, 2).Value

        wsPricing.Cells(i, 3).Value = premium
        wsPricing.Cells(i, 4).Value = expectedLoss
        wsPricing.Cells(i, 5).Value = actualClaims

        If premium > 0 Then
            wsPricing.Cells(i, 6).Value = actualClaims / premium
        Else
            wsPricing.Cells(i, 6).Value = 0
        End If

        totalPremium = totalPremium + premium
        totalExpectedLoss = totalExpectedLoss + expectedLoss
        totalActualClaims = totalActualClaims + actualClaims

    Next i

    If totalPremium <= 0 Then
        MsgBox "Total premium must be positive.", _
               vbExclamation, "Pricing Analysis"
        Exit Sub
    End If

    expectedLossRatio = totalExpectedLoss / totalPremium
    actualLossRatio = totalActualClaims / totalPremium
    underwritingResult = totalPremium - totalActualClaims

    ' Portfolio summary
    wsPricing.Range("H1").Value = "Portfolio Pricing Summary"

    wsPricing.Range("H3").Value = "Total Premium"
    wsPricing.Range("I3").Value = totalPremium

    wsPricing.Range("H4").Value = "Expected Loss"
    wsPricing.Range("I4").Value = totalExpectedLoss

    wsPricing.Range("H5").Value = "Actual Claims"
    wsPricing.Range("I5").Value = totalActualClaims

    wsPricing.Range("H6").Value = "Expected Loss Ratio"
    wsPricing.Range("I6").Value = expectedLossRatio

    wsPricing.Range("H7").Value = "Actual Loss Ratio"
    wsPricing.Range("I7").Value = actualLossRatio

    wsPricing.Range("H8").Value = "Underwriting Result"
    wsPricing.Range("I8").Value = underwritingResult

    ' Formatting
    wsPricing.Range("A1:F1").Font.Bold = True
    wsPricing.Range("A1:F1").Interior.Color = RGB(31, 78, 121)
    wsPricing.Range("A1:F1").Font.Color = RGB(255, 255, 255)

    wsPricing.Range("H1:I1").Font.Bold = True
    wsPricing.Range("H1:I1").Interior.Color = RGB(31, 78, 121)
    wsPricing.Range("H1:I1").Font.Color = RGB(255, 255, 255)

    wsPricing.Range("C2:E" & lastRow).NumberFormat = "#,##0.00"
    wsPricing.Range("F2:F" & lastRow).NumberFormat = "0.00%"

    wsPricing.Range("I3:I5").NumberFormat = "#,##0.00"
    wsPricing.Range("I6:I7").NumberFormat = "0.00%"
    wsPricing.Range("I8").NumberFormat = "#,##0.00"

    wsPricing.Columns("A:I").AutoFit

    MsgBox "Pricing and loss ratio analysis completed.", _
           vbInformation, "Pricing Analysis"

End Sub
