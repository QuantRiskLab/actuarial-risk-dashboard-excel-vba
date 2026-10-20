
Attribute VB_Name = "ClaimsSimulator"
Option Explicit

' Actuarial Risk Dashboard
' Insurance Claims Simulation

' Generate a Poisson-distributed claim count
Private Function SimulatePoisson(ByVal lambda As Double) As Long

    Dim threshold As Double
    Dim product As Double
    Dim claimCount As Long

    If lambda <= 0 Then
        SimulatePoisson = 0
        Exit Function
    End If

    threshold = Exp(-lambda)
    product = 1#
    claimCount = 0

    Do
        claimCount = claimCount + 1
        product = product * Rnd()
    Loop While product > threshold

    SimulatePoisson = claimCount - 1

End Function

' Generate a Gamma-distributed claim severity
' Shape parameter is fixed at 2
Private Function SimulateGammaSeverity( _
    ByVal expectedSeverity As Double) As Double

    Dim u1 As Double
    Dim u2 As Double

    u1 = 1# - Rnd()
    u2 = 1# - Rnd()

    SimulateGammaSeverity = _
        -(expectedSeverity / 2#) * Log(u1 * u2)

End Function

Public Sub SimulateInsuranceClaims()

    Dim wsPortfolio As Worksheet
    Dim wsClaims As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim claimCount As Long
    Dim j As Long

    Dim expectedFrequency As Double
    Dim expectedSeverity As Double
    Dim totalClaimAmount As Double
    Dim annualPremium As Double

    Set wsPortfolio = ThisWorkbook.Worksheets("Portfolio")
    Set wsClaims = ThisWorkbook.Worksheets("Claims")

    lastRow = wsPortfolio.Cells( _
        wsPortfolio.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Generate the insurance portfolio first.", _
               vbExclamation, "Claims Simulation"
        Exit Sub
    End If

    wsClaims.Cells.Clear

    wsClaims.Range("A1:G1").Value = Array( _
        "Policy_ID", _
        "Policy_Type", _
        "Expected_Frequency", _
        "Expected_Severity", _
        "Claim_Count", _
        "Total_Claims", _
        "Loss_Ratio" _
    )

    Randomize

    For i = 2 To lastRow

        expectedFrequency = wsPortfolio.Cells(i, 5).Value
        expectedSeverity = wsPortfolio.Cells(i, 6).Value
        annualPremium = wsPortfolio.Cells(i, 8).Value

        claimCount = SimulatePoisson(expectedFrequency)
        totalClaimAmount = 0#

        For j = 1 To claimCount
            totalClaimAmount = totalClaimAmount + _
                SimulateGammaSeverity(expectedSeverity)
        Next j

        wsClaims.Cells(i, 1).Value = _
            wsPortfolio.Cells(i, 1).Value

        wsClaims.Cells(i, 2).Value = _
            wsPortfolio.Cells(i, 2).Value

        wsClaims.Cells(i, 3).Value = expectedFrequency
        wsClaims.Cells(i, 4).Value = expectedSeverity
        wsClaims.Cells(i, 5).Value = claimCount
        wsClaims.Cells(i, 6).Value = Round(totalClaimAmount, 2)

        If annualPremium > 0 Then
            wsClaims.Cells(i, 7).Value = _
                totalClaimAmount / annualPremium
        Else
            wsClaims.Cells(i, 7).Value = 0
        End If

    Next i

    wsClaims.Range("A1:G1").Font.Bold = True
    wsClaims.Range("A1:G1").Interior.Color = RGB(31, 78, 121)
    wsClaims.Range("A1:G1").Font.Color = RGB(255, 255, 255)

    wsClaims.Range("C2:D" & lastRow).NumberFormat = "0.00"
    wsClaims.Range("F2:F" & lastRow).NumberFormat = "#,##0.00"
    wsClaims.Range("G2:G" & lastRow).NumberFormat = "0.00%"
    wsClaims.Columns("A:G").AutoFit

    MsgBox "Insurance claims simulation completed.", _
           vbInformation, "Claims Simulation"

End Sub
