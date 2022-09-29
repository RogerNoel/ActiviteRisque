Attribute VB_Name = "Module1"
Sub traitement()
        Application.ScreenUpdating = False
        Worksheets("Export Prisma").Name = "ExportPrisma"
    Worksheets("sheet1").Activate
    Range("A1").EntireRow.Insert
    Range("c1").Value = "num_entr"
    Range("u1").Value = "num_dossier"
    
    Dim Nbre_lignes As Long
        
    Nbre_lignes = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Nbre_lignes
        Range("u" & i).Activate
        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-18],ExportPrisma!R1C1:R4436C2,2,0)"

    Next i

Worksheets("ExportPrisma").Activate

    Nbre_lignes = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(1, 9).Value = "Description"

        For i = 2 To Nbre_lignes
                Range("j" & i).FormulaR1C1 = "=COUNTIF(Sheet1!C[11],ExportPrisma!RC[-8])"
                
                'Range("j" & i).Copy
 '   Range("j" & i).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
  '      :=False, Transpose:=False
  '  ActiveSheet.Paste
  '  Application.CutCopyMode = False
                
        Next i
        
For i = 2 To Nbre_lignes
        Range("I" & i).Select
                If Range("j" & i) <> 0 Then
                ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Codes!R1C1:R19C2,2,0)"
                Dim num_emp As Integer
                num_emp = Range("B" & i).Value
                If num_emp <> Range("B" & i + 1).Value Then
                        sheets.Add After:=sheets(sheets.Count)
                        ActiveSheet.Name = num_emp
                        Range("A1").Value = "Num. Dossier"
                        Range("B1").Value = "Société"
                        Range("C1").Value = "N° Trav"
                        Range("D1").Value = "Nom Trav."
                        Range("E1").Value = "Prén. Trav."
                        Range("F1").Value = "Code risque"
                        Range("G1").Value = "Description"
                        Range("A1:G1").font.Bold = True
                        Range("A1:G1").HorizontalAlignment = xlCenter
                        For Each iCells In Range("A1:G1")
                        iCells.BorderAround _
                        LineStyle:=xlContinuous, _
                        Weight:=xlThin
                        Next iCells
                                                
                Worksheets("ExportPrisma").Activate
                End If
        End If
Next i
    
    Dim num_feuille As Integer
    Dim sas As Integer
    num_feuille = 4
    Dim idem As Boolean
    idem = True
    
    Dim compteur As Integer
    compteur = 2
    Columns("d:d").Delete
    
    Do While Range("a" & compteur) <> ""
    If Range("i" & compteur) <> 0 Then
        If idem = False Then
            num_feuille = num_feuille + 1
        End If
        Range("B" & compteur & ":H" & compteur).Select
        Selection.Copy
        sheets(num_feuille).Select
        Rows("2:2").Select
        Selection.Insert Shift:=xlDown
        '---------------------------------
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.HorizontalAlignment = xlCenter
        '------------------------------------
        sas = Range("A2").Value
        sheets("ExportPrisma").Select
        ' ---------------------------------------------- la colonne I d'export remet tout à 0
        If sas <> Range("b" & compteur + 1).Value Then
            idem = False
        Else
            idem = True
        End If
        
        End If
        compteur = compteur + 1
    Loop
    
    Dim WS_Count As Integer
    Dim k As Integer

    WS_Count = ActiveWorkbook.Worksheets.Count

    For k = 4 To WS_Count
        sheets(k).Activate
                Columns("A:G").font.size = 10
                sheets(k).PageSetup.Orientation = xlLandscape
        Columns("A:G").AutoFit
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:="P:\Activité par rapport au risque\2022\PDF\" & "dossier " & Range("A2") & ".pdf"
    Next k
    Application.ScreenUpdating = True
    MsgBox "Traitement terminé."
    
End Sub
