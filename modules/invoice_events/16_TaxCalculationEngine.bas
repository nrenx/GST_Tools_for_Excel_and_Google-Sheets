Option Explicit
' ===============================================================================
' MODULE: 17_TaxCalculationEngine
' DESCRIPTION: Core tax calculation engine for GST, including formula setup,
'              multi-item calculations, and tax computation logic.
' ===============================================================================

' ████████████████████████████████████████████████████████████████████████████████
' 🧮 TAX CALCULATION ENGINE
' ████████████████████████████████████████████████████████████████████████████████

Public Sub SetupTaxCalculationFormulas(ws As Worksheet)
    ' Set up formulas for automatic tax calculations in the item table with enhanced structure - UPDATED FOR TWO-ROW HEADER
    On Error Resume Next

    With ws
        ' For row 19 (first item row), set up formulas - ENHANCED STRUCTURE A-O - UPDATED FOR TWO-ROW HEADER
        ' Column G (Amount) = Quantity * Rate
        .Range("G19").Formula = "=IF(AND(D19<>"""",F19<>""""),D19*F19,"""")"

        ' Column H (Taxable Value) = Amount (same as Amount for simplicity)
        .Range("H19").Formula = "=IF(G19<>"""",G19,"""")"

        ' Column I (CGST Rate) - VLOOKUP formula to get tax rate from HSN data (half of total rate for intrastate)
        .Range("I19").Formula = "=IF(N7=""Intrastate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE)/2,"""")"

        ' Column J (CGST Amount) = Taxable Value * CGST Rate / 100
        .Range("J19").Formula = "=IF(AND(H19<>"""",I19<>""""),H19*I19/100,"""")"

        ' Column K (SGST Rate) - VLOOKUP formula to get tax rate from HSN data (half of total rate for intrastate)
        .Range("K19").Formula = "=IF(N7=""Intrastate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE)/2,"""")"

        ' Column L (SGST Amount) = Taxable Value * SGST Rate / 100
        .Range("L19").Formula = "=IF(AND(H19<>"""",K19<>""""),H19*K19/100,"""")"

        ' Column M (IGST Rate) - VLOOKUP formula to get tax rate from HSN data (only for interstate)
        .Range("M19").Formula = "=IF(N7=""Interstate"",VLOOKUP(C19, warehouse!A:E, 5, FALSE),"""")"

        ' Column N (IGST Amount) = Taxable Value * IGST Rate / 100 (only for interstate)
        .Range("N19").Formula = "=IF(AND(H19<>"""",M19<>""""),H19*M19/100,"""")"

        ' Column O (Total Amount) = Taxable Value + Tax Amounts (IGST for interstate, CGST+SGST for intrastate)
        .Range("O19").Formula = "=IF(N7=""Interstate"",H19+N19,IF(N7=""Intrastate"",H19+J19+L19,H19))"

        ' Format the formula cells - ENHANCED STRUCTURE
        .Range("G19:O19").NumberFormat = "0.00"
        .Range("I19,K19,M19").NumberFormat = "0.00"
    End With

    On Error GoTo 0
End Sub

Public Sub UpdateMultiItemTaxCalculations(ws As Worksheet)
    ' Update tax calculations to sum all item rows with enhanced structure - UPDATED FOR TWO-ROW HEADER
    On Error Resume Next

    With ws
        ' Row 25: Total Quantity - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        .Range("D25").Formula = "=SUM(D19:D24)"
        .Range("D25").NumberFormat = "#,##0.00"

        ' Row 25: Sub Total calculations
        .Range("G25").Formula = "=SUM(G19:G24)"  ' Amount column
        .Range("H25").Formula = "=SUM(H19:H24)"  ' Taxable Value column
        .Range("G25:H25").NumberFormat = "#,##0.00"

        ' Row 25: Tax amounts - ENHANCED STRUCTURE - UPDATED FOR CORRECT COLUMN ORDER (CGST, SGST, IGST)
        .Range("I25").Formula = "=SUM(I19:I24)"  ' CGST Rate (average)
        .Range("J25").Formula = "=SUM(J19:J24)"  ' CGST Amount
        .Range("K25").Formula = "=SUM(K19:K24)"  ' SGST Rate (average)
        .Range("L25").Formula = "=SUM(L19:L24)"  ' SGST Amount
        .Range("M25").Formula = "=SUM(M19:M24)"  ' IGST Rate (average)
        .Range("N25").Formula = "=SUM(N19:N24)"  ' IGST Amount
        .Range("O25").Formula = "=SUM(O19:O24)"  ' Total Amount
        .Range("I25:O25").NumberFormat = "#,##0.00"

        ' Tax summary section (right side) - ENHANCED STRUCTURE - UPDATED FOR CORRECT COLUMN ORDER
        ' Row 26: Total Amount Before Tax
        .Range("O26").Formula = "=SUM(H19:H24)"

        ' Row 27: CGST
        .Range("O27").Formula = "=SUM(J19:J24)"

        ' Row 28: SGST
        .Range("O28").Formula = "=SUM(L19:L24)"

        ' Row 29: IGST
        .Range("O29").Formula = "=SUM(N19:N24)"

        ' Row 30: CESS (0 by default)
        .Range("O30").Value = 0

        ' Row 31: Total Tax
        .Range("O31").Formula = "=O27+O28+O29+O30"

        ' Row 32: Total Amount After Tax
        .Range("O32").Formula = "=O26+O31"

        ' Format all calculation cells
        .Range("O26:O32").NumberFormat = "#,##0.00"

        ' Update Amount in Words (A27 merged cell) - ENHANCED STRUCTURE - UPDATED FOR TWO-ROW HEADER
        .Range("A27").Formula = "=NumberToWords(O32)"
    End With

    On Error GoTo 0
End Sub
