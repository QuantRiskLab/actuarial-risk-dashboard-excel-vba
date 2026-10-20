
Attribute VB_Name = "RunAllModels"
Option Explicit

' Actuarial Risk Dashboard
' Master Automation Controller

Public Sub RunCompleteActuarialAnalysis()

    Dim startTime As Double
    Dim elapsedSeconds As Double

    On Error GoTo ErrorHandler

    startTime = Timer

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Step 1: Generate synthetic insurance portfolio
    Call GenerateInsurancePortfolio

    ' Step 2: Simulate insurance claims
    Call SimulateInsuranceClaims

    ' Step 3: Calculate pricing and loss ratios
    Call AnalyzeInsurancePricing

    ' Step 4: Run insurance stress scenarios
    Call RunStressTesting

    ' Step 5: Calculate illustrative solvency capital
    Call CalculateSolvencyCapital

    ' Step 6: Build the actuarial dashboard
    Call BuildActuarialDashboard

    elapsedSeconds = Timer - startTime

    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Complete actuarial analysis finished." & vbCrLf & _
           "Execution time: " & _
           Format(elapsedSeconds, "0.00") & " seconds.", _
           vbInformation, "Actuarial Risk Dashboard"

    Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "Analysis failed: " & Err.Description, _
           vbCritical, "Actuarial Risk Dashboard"

End Sub
