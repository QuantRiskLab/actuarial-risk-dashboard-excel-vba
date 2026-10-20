# Actuarial Risk Dashboard – Excel VBA

An Excel-based actuarial risk analysis and insurance portfolio dashboard developed using Visual Basic for Applications (VBA).

The project demonstrates an automated workflow for synthetic insurance portfolio generation, claims simulation, premium analysis, stress testing, and illustrative solvency capital calculations.

## Key Features

- **Insurance Portfolio Generation:** Generate 1,000 synthetic insurance policies across Motor, Property, and Liability insurance.
- **Claims Simulation:** Simulate claim frequencies using a Poisson distribution and claim severities using a Gamma distribution.
- **Pricing Analysis:** Compare expected losses, premium income, simulated claims, and underwriting performance.
- **Stress Testing:** Evaluate insurance portfolio performance under inflation, catastrophe, and combined stress scenarios.
- **Solvency Capital Analysis:** Calculate illustrative risk capital requirements and a capital coverage ratio.
- **Automated Excel Dashboard:** Display key performance indicators and charts using VBA.

## Technologies

- Microsoft Excel
- Visual Basic for Applications (VBA)
- Actuarial risk modeling
- Monte Carlo simulation
- Insurance pricing and portfolio analytics
- Dashboard automation

## Project Structure

```text
actuarial-risk-dashboard-excel-vba/
├── Actuarial_Risk_Dashboard.xlsm
├── src/
│   ├── PortfolioGenerator.bas
│   ├── ClaimsSimulator.bas
│   ├── PricingAnalysis.bas
│   ├── StressTesting.bas
│   ├── SolvencyCapital.bas
│   ├── DashboardBuilder.bas
│   └── RunAllModels.bas
└── README.md
```

## How to Run

1. Download `Actuarial_Risk_Dashboard.xlsm`.
2. Open the workbook in Microsoft Excel Desktop.
3. Enable macros after reviewing the VBA source code.
4. Open the VBA editor using `Alt + F11`.
5. Run `RunCompleteActuarialAnalysis` from the `RunAllModels` module.
6. Navigate to the `Dashboard` worksheet to explore the results.

## Dashboard Outputs

The Excel dashboard provides:

- Total premium income
- Total simulated claims
- Portfolio loss ratio
- Underwriting performance analysis
- Illustrative capital requirement
- Capital coverage ratio
- Stress scenario visualizations

## Methodological Note

All insurance policies and claims are synthetically generated. The model is intended for educational and portfolio demonstration purposes.

The solvency capital calculations are simplified illustrations and do not constitute a regulatory Solvency Capital Requirement (SCR) calculation under Solvency II.

## Skills Demonstrated

- Actuarial modeling and insurance analytics
- VBA programming and workflow automation
- Quantitative risk assessment
- Stochastic claims simulation
- Financial reporting and data visualization

## Author

QuantRiskLab

*Independent actuarial analytics portfolio project.*
