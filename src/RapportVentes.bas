Attribute VB_Name = "RapportVentes"
Option Explicit

' =============================================================================
'  PROJET : AUTOMATISATION RAPPORT DE VENTES
'  Auteur  : Souheil
'  Description : Genere un dashboard KPI + graphiques depuis la feuille Donnees
'                et exporte le rapport en PDF
' =============================================================================

Private Const SHEET_DATA      As String = "Données"
Private Const SHEET_DASHBOARD As String = "Dashboard"
Private Const SHEET_PIVOT     As String = "Analyse"
Private Const COULEUR_BLEU    As Long = 2040607
Private Const COULEUR_VERT    As Long = 2016000
Private Const COULEUR_GRIS    As Long = 14277081


' =============================================================================
'  MACRO PRINCIPALE
' =============================================================================
Sub GenererDashboard()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo GestionErreur

    If Not SheetExiste(SHEET_DATA) Then
        MsgBox "La feuille '" & SHEET_DATA & "' est introuvable !", vbCritical
        GoTo Nettoyage
    End If

    SupprimerFeuille SHEET_DASHBOARD
    SupprimerFeuille SHEET_PIVOT
    CreerFeuilleAnalyse
    CreerDashboard

    ThisWorkbook.Sheets(SHEET_DASHBOARD).Activate
    MsgBox "Dashboard genere avec succes !" & vbNewLine & _
           "Utilisez 'ExporterPDF' pour exporter en PDF.", vbInformation, "Succes"
    GoTo Nettoyage

GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Erreur"

Nettoyage:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub


' =============================================================================
'  FEUILLE ANALYSE
' =============================================================================
Private Sub CreerFeuilleAnalyse()

    Dim wsData As Worksheet
    Dim wsAna  As Worksheet

    Set wsData = ThisWorkbook.Sheets(SHEET_DATA)
    Set wsAna  = ThisWorkbook.Sheets.Add(After:=wsData)
    wsAna.Name = SHEET_PIVOT

    Dim DerniereLigne As Long
    DerniereLigne = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' KPI Globaux
    wsAna.Range("A1").Value  = "KPI GLOBAUX"
    wsAna.Range("A2").Value  = "CA Total (EUR)"
    wsAna.Range("B2").Formula = "=SUMIF(Données!I:I,""Payé"",Données!H:H)"
    wsAna.Range("A3").Value  = "Nb Commandes"
    wsAna.Range("B3").Formula = "=COUNTA(Données!A2:A10000)-1"
    wsAna.Range("A4").Value  = "Nb Commandes Payees"
    wsAna.Range("B4").Formula = "=COUNTIF(Données!I:I,""Payé"")"
    wsAna.Range("A5").Value  = "Nb Commandes En attente"
    wsAna.Range("B5").Formula = "=COUNTIF(Données!I:I,""En attente"")"
    wsAna.Range("A6").Value  = "Nb Commandes Annulees"
    wsAna.Range("B6").Formula = "=COUNTIF(Données!I:I,""Annulé"")"
    wsAna.Range("A7").Value  = "Panier Moyen (EUR)"
    wsAna.Range("B7").Formula = "=IFERROR(B2/B4,0)"
    wsAna.Range("A8").Value  = "Taux Annulation (%)"
    wsAna.Range("B8").Formula = "=IFERROR(B6/B3,0)"

    ' CA par Region
    wsAna.Range("D1").Value = "CA PAR REGION"
    Dim regions As Variant
    regions = Array("Paris", "Lyon", "Marseille", "Bordeaux")
    Dim i As Integer
    For i = 0 To 3
        wsAna.Cells(2 + i, 4).Value = regions(i)
        wsAna.Cells(2 + i, 5).Formula = _
            "=SUMPRODUCT((Données!D$2:Données!D$" & DerniereLigne & "=""" & regions(i) & """)*" & _
            "(Données!I$2:Données!I$" & DerniereLigne & "=""Payé"")*" & _
            "(Données!H$2:Données!H$" & DerniereLigne & "))"
    Next i

    ' CA par Vendeur
    wsAna.Range("G1").Value = "CA PAR VENDEUR"
    Dim vendeurs As Variant
    vendeurs = Array("Alice Martin", "Bruno Dupont", "Carla Leroy", "David Moreau", "Emma Bernard")
    For i = 0 To 4
        wsAna.Cells(2 + i, 7).Value = vendeurs(i)
        wsAna.Cells(2 + i, 8).Formula = _
            "=SUMPRODUCT((Données!C$2:Données!C$" & DerniereLigne & "=""" & vendeurs(i) & """)*" & _
            "(Données!I$2:Données!I$" & DerniereLigne & "=""Payé"")*" & _
            "(Données!H$2:Données!H$" & DerniereLigne & "))"
    Next i

    ' CA par Produit
    wsAna.Range("J1").Value = "CA PAR PRODUIT"
    Dim produits As Variant
    produits = Array("Laptop", "Smartphone", "Tablette", "Montre", "Accessoires")
    For i = 0 To 4
        wsAna.Cells(2 + i, 10).Value = produits(i)
        wsAna.Cells(2 + i, 11).Formula = _
            "=SUMPRODUCT((Données!E$2:Données!E$" & DerniereLigne & "=""" & produits(i) & """)*" & _
            "(Données!I$2:Données!I$" & DerniereLigne & "=""Payé"")*" & _
            "(Données!H$2:Données!H$" & DerniereLigne & "))"
    Next i

    ' CA par Mois
    wsAna.Range("A11").Value = "CA PAR MOIS"
    Dim mois As Variant
    mois = Array("Janv", "Fevr", "Mars", "Avr", "Mai", "Juin", _
                 "Juil", "Aout", "Sept", "Oct", "Nov", "Dec")
    For i = 0 To 11
        wsAna.Cells(12 + i, 1).Value = mois(i)
        wsAna.Cells(12 + i, 2).Formula = _
            "=SUMPRODUCT((MONTH(Données!B$2:Données!B$" & DerniereLigne & ")=" & (i + 1) & ")*" & _
            "(Données!I$2:Données!I$" & DerniereLigne & "=""Payé"")*" & _
            "(Données!H$2:Données!H$" & DerniereLigne & "))"
    Next i

    wsAna.Visible = xlSheetHidden

End Sub


' =============================================================================
'  DASHBOARD VISUEL
' =============================================================================
Private Sub CreerDashboard()

    Dim wsDash As Worksheet
    Dim wsAna  As Worksheet

    Set wsAna  = ThisWorkbook.Sheets(SHEET_PIVOT)
    Set wsDash = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    wsDash.Name = SHEET_DASHBOARD

    wsDash.Cells.Interior.Color = RGB(245, 247, 252)
    wsDash.Cells.Font.Name = "Arial"

    ' En-tete
    Dim rngTitre As Range
    Set rngTitre = wsDash.Range("A1:P2")
    rngTitre.Merge
    rngTitre.Value = "RAPPORT DE VENTES 2024 - TABLEAU DE BORD"
    rngTitre.Interior.Color = COULEUR_BLEU
    rngTitre.Font.Color = RGB(255, 255, 255)
    rngTitre.Font.Bold = True
    rngTitre.Font.Size = 18
    rngTitre.RowHeight = 45
    rngTitre.HorizontalAlignment = xlCenter
    rngTitre.VerticalAlignment = xlCenter

    ' Sous-titre
    Dim rngDate As Range
    Set rngDate = wsDash.Range("A3:P3")
    rngDate.Merge
    rngDate.Value = "Genere le : " & Format(Now(), "DD/MM/YYYY HH:MM")
    rngDate.Font.Italic = True
    rngDate.Font.Size = 10
    rngDate.Font.Color = RGB(100, 100, 100)
    rngDate.HorizontalAlignment = xlCenter

    ' Ligne separatrice
    wsDash.Range("A4:P4").Interior.Color = COULEUR_BLEU
    wsDash.Rows(4).RowHeight = 3

    AfficherKPI wsDash, wsAna
    AjouterGraphiqueRegion  wsDash, wsAna
    AjouterGraphiqueMois    wsDash, wsAna
    AjouterGraphiqueVendeur wsDash, wsAna
    AjouterGraphiqueProduit wsDash, wsAna

    Dim col As Integer
    For col = 1 To 16
        wsDash.Columns(col).ColumnWidth = 10
    Next col
    wsDash.Rows(5).RowHeight = 10

    Dim btn As Object
    Set btn = wsDash.Buttons.Add(10, 5, 200, 25)
    btn.Caption = "Exporter en PDF"
    btn.OnAction = "ExporterPDF"
    btn.Font.Bold = True
    btn.Font.Size = 11

End Sub


' =============================================================================
'  KPI CARDS
' =============================================================================
Private Sub AfficherKPI(wsDash As Worksheet, wsAna As Worksheet)

    Dim labels   As Variant
    Dim valCells As Variant
    Dim colors   As Variant
    Dim formats  As Variant

    labels   = Array("CA Total (EUR)", "Commandes", "Payees", "En Attente", "Panier Moyen (EUR)", "Taux Annulation")
    valCells = Array("B2", "B3", "B4", "B5", "B7", "B8")
    colors   = Array(RGB(31, 56, 100), RGB(52, 152, 219), RGB(39, 174, 96), _
                     RGB(243, 156, 18), RGB(142, 68, 173), RGB(231, 76, 60))
    formats  = Array("#,##0", "0", "0", "0", "#,##0", "0.0%")

    Dim i As Integer
    For i = 0 To 5
        Dim c As Integer
        c = 1 + i * 2 + i

        Dim rngCard As Range
        Set rngCard = wsDash.Range(wsDash.Cells(6, c), wsDash.Cells(10, c + 1))
        rngCard.Interior.Color = colors(i)
        rngCard.HorizontalAlignment = xlCenter
        rngCard.VerticalAlignment = xlCenter

        Dim rngLabel As Range
        Set rngLabel = wsDash.Range(wsDash.Cells(6, c), wsDash.Cells(7, c + 1))
        rngLabel.Merge
        rngLabel.Value = labels(i)
        rngLabel.Font.Color = RGB(255, 255, 255)
        rngLabel.Font.Bold = True
        rngLabel.Font.Size = 10
        rngLabel.HorizontalAlignment = xlCenter
        rngLabel.VerticalAlignment = xlCenter

        Dim rngVal As Range
        Set rngVal = wsDash.Range(wsDash.Cells(8, c), wsDash.Cells(10, c + 1))
        rngVal.Merge
        rngVal.Value = wsAna.Range(valCells(i)).Value
        rngVal.NumberFormat = formats(i)
        rngVal.Font.Color = RGB(255, 255, 255)
        rngVal.Font.Bold = True
        rngVal.Font.Size = 16
        rngVal.HorizontalAlignment = xlCenter
        rngVal.VerticalAlignment = xlCenter
    Next i

End Sub


' =============================================================================
'  GRAPHIQUES
' =============================================================================
Private Sub AjouterGraphiqueRegion(wsDash As Worksheet, wsAna As Worksheet)
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(Left:=10, Top:=200, Width:=380, Height:=250)
    With cht.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsAna.Range("D2:E5")
        .HasTitle = True
        .ChartTitle.Text = "CA par Region (EUR)"
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Color = COULEUR_BLEU
        .PlotArea.Interior.Color = RGB(245, 247, 252)
        .ChartArea.Border.LineStyle = xlNone
        .SeriesCollection(1).Interior.Color = RGB(52, 152, 219)
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With
End Sub

Private Sub AjouterGraphiqueMois(wsDash As Worksheet, wsAna As Worksheet)
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(Left:=410, Top:=200, Width:=420, Height:=250)
    With cht.Chart
        .ChartType = xlLineMarkers
        .SetSourceData wsAna.Range("A12:B23")
        .HasTitle = True
        .ChartTitle.Text = "Evolution CA Mensuel (EUR)"
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Color = COULEUR_BLEU
        .PlotArea.Interior.Color = RGB(245, 247, 252)
        .ChartArea.Border.LineStyle = xlNone
        With .SeriesCollection(1)
            .Border.Color = RGB(39, 174, 96)
            .MarkerBackgroundColor = RGB(39, 174, 96)
            .MarkerForegroundColor = RGB(39, 174, 96)
            .MarkerSize = 6
        End With
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With
End Sub

Private Sub AjouterGraphiqueVendeur(wsDash As Worksheet, wsAna As Worksheet)
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(Left:=10, Top:=470, Width:=380, Height:=250)
    With cht.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsAna.Range("G2:H6")
        .HasTitle = True
        .ChartTitle.Text = "CA par Vendeur (EUR)"
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Color = COULEUR_BLEU
        .PlotArea.Interior.Color = RGB(245, 247, 252)
        .ChartArea.Border.LineStyle = xlNone
        .SeriesCollection(1).Interior.Color = RGB(142, 68, 173)
        .HasLegend = False
        .Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    End With
End Sub

Private Sub AjouterGraphiqueProduit(wsDash As Worksheet, wsAna As Worksheet)
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects.Add(Left:=410, Top:=470, Width:=420, Height:=250)
    With cht.Chart
        .ChartType = xlPie
        .SetSourceData wsAna.Range("J2:K6")
        .HasTitle = True
        .ChartTitle.Text = "Repartition CA par Produit"
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Color = COULEUR_BLEU
        .PlotArea.Interior.Color = RGB(245, 247, 252)
        .ChartArea.Border.LineStyle = xlNone
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .ApplyDataLabels
        .SeriesCollection(1).DataLabels.NumberFormat = "0.0%"
        .SeriesCollection(1).DataLabels.ShowPercentage = True
        .SeriesCollection(1).DataLabels.ShowValue = False
    End With
End Sub


' =============================================================================
'  EXPORT PDF
' =============================================================================
Sub ExporterPDF()

    Dim wsDash As Worksheet

    If Not SheetExiste(SHEET_DASHBOARD) Then
        MsgBox "Le Dashboard n'existe pas encore." & vbNewLine & _
               "Lancez d'abord 'GenererDashboard'.", vbExclamation
        Exit Sub
    End If

    Set wsDash = ThisWorkbook.Sheets(SHEET_DASHBOARD)

    Dim cheminPDF As String
    cheminPDF = ThisWorkbook.Path & "\Rapport_Ventes_" & Format(Now(), "YYYYMMDD_HHMM") & ".pdf"

    With wsDash.PageSetup
        .Orientation    = xlLandscape
        .PaperSize      = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin      = Application.InchesToPoints(0.5)
        .BottomMargin   = Application.InchesToPoints(0.5)
        .LeftMargin     = Application.InchesToPoints(0.4)
        .RightMargin    = Application.InchesToPoints(0.4)
        .CenterHeader   = "&""Arial,Bold""&14RAPPORT DE VENTES 2024"
        .CenterFooter   = "Page &P sur &N  -  Confidentiel  -  Genere le &D"
        .PrintGridlines = False
    End With

    On Error GoTo ErreurPDF
    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=cheminPDF, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "PDF exporte avec succes !" & vbNewLine & cheminPDF, vbInformation, "Export PDF"
    Exit Sub

ErreurPDF:
    MsgBox "Erreur lors de l'export PDF : " & Err.Description, vbCritical

End Sub


' =============================================================================
'  UTILITAIRES
' =============================================================================
Private Function SheetExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    SheetExiste = Not ws Is Nothing
End Function

Private Sub SupprimerFeuille(nom As String)
    If SheetExiste(nom) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(nom).Delete
        Application.DisplayAlerts = True
    End If
End Sub
