
Attribute VB_Name = "PortfolioGenerator"
Option Explicit

' Actuarial Risk Dashboard
' Synthetic Insurance Portfolio Generator

Public Sub GenerateInsurancePortfolio()

    Dim ws As Worksheet
    Dim i As Long
    Dim policyCount As Long
    Dim annualPremium As Double
    Dim expectedFrequency As Double
    Dim expectedSeverity As Double
    Dim riskFactor As Double
    Dim policyType As String

    Set ws = ThisWorkbook.Worksheets("Portfolio")

    policyCount = 1000

    ws.Cells.Clear

    ws.Range("A1:H1").Value = Array( _
        "Policy_ID", _
        "Policy_Type", _
        "Age", _
        "Risk_Factor", _
        "Expected_Frequency", _
        "Expected_Severity", _
        "Expected_Loss", _
        "Annual_Premium" _
    )

    Randomize

    For i = 2 To policyCount + 1

        Select Case Int(Rnd() * 3)

            Case 0
                policyType = "Motor"
                expectedFrequency = 0.12
                expectedSeverity = 2500

            Case 1
                policyType = "Property"
                expectedFrequency = 0.06
                expectedSeverity = 6000

            Case Else
                policyType = "Liability"
                expectedFrequency = 0.04
                expectedSeverity = 10000

        End Select

        riskFactor = 0.7 + Rnd() * 0.8

        expectedFrequency = expectedFrequency * riskFactor

        annualPremium = expectedFrequency * expectedSeverity * 1.25

        ws.Cells(i, 1).Value = i - 1
        ws.Cells(i, 2).Value = policyType
        ws.Cells(i, 3).Value = Int(18 + Rnd() * 63)
        ws.Cells(i, 4).Value = Round(riskFactor, 4)
        ws.Cells(i, 5).Value = Round(expectedFrequency, 4)
        ws.Cells(i, 6).Value = expectedSeverity
        ws.Cells(i, 7).Value = Round( _
            expectedFrequency * expectedSeverity, 2)
        ws.Cells(i, 8).Value = Round(annualPremium, 2)

    Next i

    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)

    ws.Range("D2:E1001").NumberFormat = "0.0000"
    ws.Range("F2:H1001").NumberFormat = "#,##0.00"
    ws.Columns("A:H").AutoFit

    MsgBox "1,000 synthetic insurance policies generated.", _
           vbInformation, "Portfolio Generator"

End Sub
