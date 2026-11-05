Attribute VB_Name = "modInleesFirap"
Option Explicit
'==============================================================================
' CODA-inleesbestand voor een financiele rapportage (firap)
'
' Knop / macro : MaakOfWerkInleesbestandBij
'   Leest het bronblad "1e/2e/3e firap <jaar>" (naam wordt afgelezen van de
'   werkmap), bouwt het tabblad "tbv Coda" en het tabblad "controle inlees".
'   Bestaan de tabbladen al, dan worden ze opnieuw opgebouwd (bijwerken).
'
' Niet gebruiken voor begrotingswijzigingen (coda-raadswijziging,
' coda-adm-wijziging-inlezen), de investerings- en financieringsstaat
' (inlees-coda-ienf-v2) of de taakveldbladen (taakvelden-ai-oa-inlees).
'
' Het bronblad wordt nooit gewijzigd.
' Omschrijvingen (kolom E) komen uit het tabblad "omschrijvingen"
' (sleutel > samenvatting). Onbekende regels krijgen een CONCEPT dat u daar
' moet nalopen; daarna opnieuw op de knop drukken.
'
' Het bestand bevat uitsluitend ASCII-tekens, zodat het zonder problemen als
' .bas kan worden geimporteerd (VBA-editor > Bestand > Bestand importeren).
'==============================================================================

Private Const SH_CODA As String = "tbv Coda"
Private Const SH_CONTROLE As String = "controle inlees"
Private Const SH_OMSCHR As String = "omschrijvingen"
Private Const SH_MELD As String = "meldingen inlees"
Private Const MAX_OMS As Long = 35
Private Const BALANS_EERSTE_CIJFER As String = "01"   'balansnummers (7 cijfers) beginnen met 0 of 1
Private Const HULPKOL As Long = 14                    'kolom N op controle inlees: deelkenmerk

'--- toestand van een run ------------------------------------------------------
Private v As Variant            'bronblad, waarden
Private brn As Worksheet
Private ny As Long              'aantal jaren
Private yr() As Long            'jaartallen uit rij 10
Private lastR As Long
Private wijz As String          'W16/26
Private wijzStreep As String    'W16-26
Private n As Long               'aantal regels (excl. saldo- en sluitregel)
Private rGrb() As Variant
Private rCat() As Variant
Private rVerb() As String
Private rDeel() As String
Private rRij() As Long
Private rBed() As Variant       '(regel, jaar)
Private rSleutel() As String
Private rInfo() As String
Private rConcept() As String
Private rOms() As String
Private rOmsSt() As String
Private mMeld As Collection     'fouten: stoppen
Private mInfo As Collection     'overgeslagen bronregels
Private mStil As Boolean
Public gResultaat As String
Private mSluit As String
Private mLaatste As Long
Private mAantalRegels As Long

'rijnummers in het bronblad
Private rKop1 As Long, rSal1 As Long, rKop2 As Long, rTot2 As Long
Private rKop3a As Long, rSal3a As Long, rKopRes As Long, rSalRes As Long
Private rKopVrz As Long, rSalVrz As Long

'==============================================================================
' Ingang voor de knop
'==============================================================================
Public Sub MaakOfWerkInleesbestandBij()
    mStil = False
    Voer_uit
End Sub

'Zonder meldingsvensters (voor tests / automatisch gebruik). Resultaat in gResultaat.
Public Sub InleesbestandStil()
    mStil = True
    Voer_uit
End Sub

Private Sub Voer_uit()
    Dim wbk As Workbook
    Dim wsCoda As Worksheet
    Dim antw As VbMsgBoxResult
    Dim oudCalc As XlCalculation
    On Error GoTo Fout
    Set wbk = ActiveWorkbook
    Set mMeld = New Collection
    Set mInfo = New Collection
    gResultaat = ""

    Set brn = ZoekBronblad(wbk)
    If brn Is Nothing Then Exit Sub

    'bestaand inleesbestand? dan eerst vragen
    Set wsCoda = SheetByName(wbk, SH_CODA)
    If Not mStil Then
        If Not wsCoda Is Nothing Then
            If Application.WorksheetFunction.CountA(wsCoda.Cells) > 0 Then
                antw = MsgBox("Het tabblad '" & SH_CODA & "' en '" & SH_CONTROLE & "' worden opnieuw opgebouwd uit '" & brn.Name & "'." & vbCrLf & vbCrLf & "Doorgaan?", vbYesNo + vbQuestion, "Inleesbestand bijwerken")
                If antw <> vbYes Then Exit Sub
            End If
        End If
    End If

    Application.ScreenUpdating = False
    oudCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    '--- stap 1: startcontrole --------------------------------------------------
    If Not Startcontrole() Then
        Meld_stop "Startcontrole bronblad"
        GoTo Einde
    End If

    '--- stap 2: regels afleiden ------------------------------------------------
    ReDim rGrb(1 To lastR + 5)
    ReDim rCat(1 To lastR + 5)
    ReDim rVerb(1 To lastR + 5)
    ReDim rDeel(1 To lastR + 5)
    ReDim rRij(1 To lastR + 5)
    ReDim rSleutel(1 To lastR + 5)
    ReDim rInfo(1 To lastR + 5)
    ReDim rConcept(1 To lastR + 5)
    ReDim rOms(1 To lastR + 5)
    ReDim rOmsSt(1 To lastR + 5)
    ReDim rBed(1 To lastR + 5, 1 To ny)
    n = 0
    VerwerkDeel "1", rKop1 + 1, rSal1 - 1
    VerwerkDeel "3a", rKop3a + 1, rSal3a - 1
    VerwerkDeel "3b1", rKopRes + 1, rSalRes - 1
    VerwerkDeel "3b2", rKopVrz + 1, rSalVrz - 1
    If mMeld.Count > 0 Then
        Meld_stop "Controle op categorie en N/B"
        GoTo Einde
    End If

    '--- omschrijvingen (kolom E) -------------------------------------------------
    BepaalOmschrijvingen wbk

    '--- stap 3 t/m 5: tbv Coda -------------------------------------------------
    Dim nConcept As Long
    nConcept = SchrijfCoda(wbk)

    '--- stap 6: controle inlees ------------------------------------------------
    SchrijfControle wbk, nConcept

    Application.Calculation = oudCalc
    Application.Calculate

    '--- harde poort: statusregel terugkoppelen ------------------------------------
    Dim wsK As Worksheet, st As String
    Set wsK = SheetByName(wbk, SH_CONTROLE)
    st = CStr(wsK.Range("B3").Value)
    VerwijderMeldingenBlad wbk
    Terugkoppeling st, nConcept
    wsK.Activate
    wsK.Range("A1").Select
    GoTo Einde

Fout:
    Dim fe As String
    fe = "Fout " & Err.Number & ": " & Err.Description
    gResultaat = fe
    If Not mStil Then MsgBox fe, vbCritical, "Inleesbestand"
Einde:
    On Error Resume Next
    Application.Calculation = oudCalc
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' Bronblad zoeken: naam wisselt per rapportage (1e / 2e / 3e firap <jaar>)
'==============================================================================
Private Function ZoekBronblad(wbk As Workbook) As Worksheet
    Dim ws As Worksheet, k As Long, gevonden As Worksheet
    For Each ws In wbk.Worksheets
        If LCase$(ws.Name) Like "[123]e firap ####" Then
            k = k + 1
            Set gevonden = ws
        End If
    Next ws
    If k = 0 Then
        Zeg "Geen bronblad gevonden. Verwacht: een tabblad met de naam '1e firap <jaar>', '2e firap <jaar>' of '3e firap <jaar>'.", vbExclamation
        Exit Function
    End If
    If k > 1 Then
        If LCase$(ActiveSheet.Name) Like "[123]e firap ####" Then
            Set ZoekBronblad = ActiveSheet
        Else
            Zeg "Er zijn " & k & " bronbladen. Activeer het juiste bronblad en start de macro opnieuw.", vbExclamation
        End If
        Exit Function
    End If
    Set ZoekBronblad = gevonden
End Function

'==============================================================================
' Stap 1: startcontrole
'==============================================================================
Private Function Startcontrole() As Boolean
    Dim r As Long, c As Long, k As Long, a As String, t As String
    Dim ur As Range
    Set ur = brn.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    If lastR < 11 Then lastR = 11
    v = brn.Range(brn.Cells(1, 1), brn.Cells(lastR, 19)).Value

    'wijzigingsnummer (W13/26)
    wijz = ""
    For r = 1 To 9
        For c = 1 To 19
            t = Txt(v(r, c))
            If IsWijzNr(t) Then wijz = t: Exit For
        Next c
        If wijz <> "" Then Exit For
    Next r
    If wijz = "" Then
        AddMeld "Wijzigingsnummer (bijv. W13/26) niet gevonden in rij 1 t/m 9 van '" & brn.Name & "'."
    Else
        wijzStreep = Replace(wijz, "/", "-")
    End If

    'rij 10: verwachte paren N/B + jaar
    If LCase$(Txt(v(10, 1))) <> "grootboek" Then AddMeld "Rij 10, kolom A: 'Grootboek' verwacht, gevonden: '" & Txt(v(10, 1)) & "'."
    ny = 0
    c = 5
    Do While c + 1 <= 19
        If UCase$(Txt(v(10, c))) = "N/B" And IsNum(v(10, c + 1)) Then
            ny = ny + 1
            ReDim Preserve yr(1 To ny)
            yr(ny) = CLng(v(10, c + 1))
            c = c + 2
        Else
            Exit Do
        End If
    Loop
    If ny = 0 Then AddMeld "Rij 10 laat vanaf kolom E niet de verwachte paren (N/B + jaartal) zien."

    'kopregels en saldoregels op hun tekst
    rKop1 = 0: rSal1 = 0: rKop2 = 0: rTot2 = 0: rKop3a = 0: rSal3a = 0
    rKopRes = 0: rSalRes = 0: rKopVrz = 0: rSalVrz = 0
    For r = 11 To lastR
        a = LCase$(Txt(v(r, 1)))
        t = LCase$(Txt(v(r, 3)))
        If rKop1 = 0 And a Like "1. mutaties in de raming*" Then rKop1 = r
        If rKop2 = 0 And a Like "2. mutaties in onvoorziene*" Then rKop2 = r
        If rTot2 = 0 And (a Like "totaal 2*" Or t Like "totaal 2*") Then rTot2 = r
        If rKop3a = 0 And a Like "3a. investeringen*" Then rKop3a = r
        If rKopRes = 0 And a = "reserves" Then rKopRes = r
        If rKopVrz = 0 And a = "voorzieningen" Then rKopVrz = r
        If rKop1 > 0 And rSal1 = 0 And r > rKop1 And t Like "salderingsrekening begrotingssaldo*" Then rSal1 = r
        If rKop3a > 0 And rSal3a = 0 And r > rKop3a And t Like "saldo*investeringen*" Then rSal3a = r
        If rKopRes > 0 And rSalRes = 0 And r > rKopRes And t Like "saldo*reserves*" Then rSalRes = r
        If rKopVrz > 0 And rSalVrz = 0 And r > rKopVrz And t Like "saldo*voorzieningen*" Then rSalVrz = r
    Next r
    If rKop1 = 0 Then AddMeld "Kopregel '1. Mutaties in de raming van baten en lasten' niet gevonden."
    If rSal1 = 0 Then AddMeld "Saldoregel 'Salderingsrekening begrotingssaldo' (deel 1) niet gevonden in kolom C."
    If rKop2 = 0 Then AddMeld "Kopregel '2. Mutaties in onvoorziene uitgaven' niet gevonden."
    If rTot2 = 0 Then AddMeld "Saldoregel 'Totaal 2, saldo budget voor onvoorziene uitgaven' niet gevonden."
    If rKop3a = 0 Then AddMeld "Kopregel '3a. Investeringen' niet gevonden."
    If rSal3a = 0 Then AddMeld "Saldoregel 'Saldo ... investeringen' (3a) niet gevonden in kolom C."
    If rKopRes = 0 Then AddMeld "Kopregel 'Reserves' niet gevonden in kolom A."
    If rSalRes = 0 Then AddMeld "Saldoregel 'Saldo ... reserves' (3b-1) niet gevonden in kolom C."
    If rKopVrz = 0 Then AddMeld "Kopregel 'Voorzieningen' niet gevonden in kolom A."
    If rSalVrz = 0 Then AddMeld "Saldoregel 'Saldo ... voorzieningen' (3b-2) niet gevonden in kolom C."
    If mMeld.Count = 0 Then
        If Not (rKop1 < rSal1 And rSal1 < rKop2 And rKop2 < rKop3a And rKop3a < rSal3a And rSal3a < rKopRes And rKopRes < rSalRes And rSalRes < rKopVrz And rKopVrz < rSalVrz) Then
            AddMeld "De kop- en saldoregels staan niet in de verwachte volgorde (deel 1, 2, 3a, reserves, voorzieningen)."
        End If
    End If

    'negatieve bedragen in de bedragkolommen
    If ny > 0 Then
        For r = 11 To lastR
            For k = 1 To ny
                If IsNum(v(r, 4 + 2 * k)) Then
                    If v(r, 4 + 2 * k) < 0 Then AddMeld "Bronrij " & r & ": negatief bedrag (" & v(r, 4 + 2 * k) & ") in kolom " & ColL(4 + 2 * k) & ". In de bron moeten alle bedragen positief zijn."
                End If
            Next k
        Next r
    End If

    Startcontrole = (mMeld.Count = 0)
End Function

'==============================================================================
' Stap 2: regels afleiden uit een deel van het bronblad
'==============================================================================
Private Sub VerwerkDeel(ByVal deel As String, ByVal r1 As Long, ByVal r2 As Long)
    Dim r As Long, k As Long, p As Long
    Dim a As String, c As String, q As String
    Dim curNr As String, curNaam As String
    Dim hasAmt As Boolean, hasTekst As Boolean
    Dim catTxt As String, verb As String
    Dim nb As String, bedr As Double, sg As Long, eersteTeken As String
    Dim deelDeel As String, aOut As Variant, cOut As String, sleutelDeel As String
    Dim ok As Boolean, fout As Boolean
    Dim tekstBasis As String

    For r = r1 To r2
        a = Txt(v(r, 1))
        c = Txt(v(r, 3))
        q = Txt(v(r, 17))
        hasAmt = False: hasTekst = False
        For k = 1 To ny
            If IsNum(v(r, 4 + 2 * k)) Then
                If v(r, 4 + 2 * k) <> 0 Then hasAmt = True
            ElseIf Txt(v(r, 4 + 2 * k)) <> "" Then
                hasTekst = True
            End If
        Next k

        If Not hasAmt Then
            'kopregel zonder bedrag
            Select Case deel
                Case "1"
                    If Len(a) = 6 And AllDigits(a) Then curNr = a: curNaam = c
                Case "3a"
                    If Len(a) = 7 And AllDigits(a) Then curNr = a: curNaam = c
                Case "3b1"
                    If IsKopNr(a, "R") Then curNr = a: curNaam = c
                Case "3b2"
                    If IsKopNr(a, "O") Then curNr = a: curNaam = c
            End Select
            If hasTekst Then mInfo.Add "Bronrij " & r & ": geen bedrag (tekst in bedragkolom, bijv. pm), vervallen."
            GoTo Volgende
        End If

        If LCase$(a) Like "totaal*" Then GoTo Volgende

        'categorie uit kolom A van de bedragregel zelf
        catTxt = a: verb = ""
        p = InStr(a, "-")
        If p > 0 Then
            catTxt = Trim$(Left$(a, p - 1))
            verb = Trim$(Mid$(a, p + 1))
        End If
        ok = (Len(catTxt) = 5 And AllDigits(catTxt))
        If ok Then ok = (Left$(catTxt, 1) = "4" Or Left$(catTxt, 1) = "8")
        If Not ok Then
            If (Len(catTxt) = 6 Or Len(catTxt) = 7) And AllDigits(catTxt) Then
                AddMeld "Bronrij " & r & ": op de bedragregel staat het grootboek- of investeringsnummer '" & a & "' in kolom A; de categorie staat (als die er al is) op de regel eronder (omschrijving: " & Left$(c, 50) & "). Zet de categorie op de bedragregel zelf."
            Else
                AddMeld "Bronrij " & r & ": categorie '" & a & "' is geen 5-cijferig nummer dat met 4 of 8 begint (omschrijving: " & Left$(c, 50) & "). Corrigeer in de bron."
            End If
            GoTo Volgende
        End If
        If curNr = "" Then
            AddMeld "Bronrij " & r & ": geen kopregel (grootboek-, investerings-, R- of O-nummer) boven deze bedragregel gevonden."
            GoTo Volgende
        End If

        'bedragen met teken uit de N/B-kolom links van het bedrag
        n = n + 1
        eersteTeken = ""
        fout = False
        For k = 1 To ny
            rBed(n, k) = Empty
            If IsNum(v(r, 4 + 2 * k)) Then
                If v(r, 4 + 2 * k) <> 0 Then
                    nb = UCase$(Txt(v(r, 3 + 2 * k)))
                    If nb = "N" Then
                        sg = 1
                    ElseIf nb = "B" Then
                        sg = -1
                    Else
                        AddMeld "Bronrij " & r & ": bedrag " & v(r, 4 + 2 * k) & " in " & yr(k) & " (kolom " & ColL(4 + 2 * k) & ") zonder N of B in kolom " & ColL(3 + 2 * k) & "."
                        fout = True
                        sg = 0
                    End If
                    If sg <> 0 Then
                        bedr = Int(Abs(v(r, 4 + 2 * k)) + 0.5)
                        rBed(n, k) = sg * bedr
                        If eersteTeken = "" Then eersteTeken = nb
                    End If
                End If
            End If
        Next k
        If fout Then n = n - 1: GoTo Volgende

        'kolommen A, C en deelkenmerk
        Select Case deel
            Case "1"
                aOut = CLng(curNr): cOut = verb
                sleutelDeel = curNr
            Case "3a"
                If InStr(BALANS_EERSTE_CIJFER, Left$(curNr, 1)) > 0 Then
                    aOut = curNr: cOut = ""         'balansnummer: nummer zelf, als tekst
                Else
                    aOut = "+": cOut = curNr        'investeringsnummer
                End If
                sleutelDeel = "+" & curNr
            Case Else
                aOut = "+": cOut = curNr
                sleutelDeel = "+" & curNr
        End Select
        rGrb(n) = aOut
        rCat(n) = CLng(catTxt)
        rVerb(n) = cOut
        rDeel(n) = deel
        rRij(n) = r
        rSleutel(n) = sleutelDeel & "|" & eersteTeken & "|" & NormTxt(c)
        tekstBasis = q
        If tekstBasis = "" Then tekstBasis = c
        If tekstBasis = "" Then tekstBasis = curNaam
        rInfo(n) = Left$(tekstBasis, 160)
        rConcept(n) = ConceptOms(tekstBasis)
Volgende:
    Next r
End Sub

'==============================================================================
' Omschrijvingen: opzoeken in tabblad 'omschrijvingen', anders CONCEPT
'==============================================================================
Private Sub BepaalOmschrijvingen(wbk As Workbook)
    Dim ws As Worksheet, i As Long, lr As Long, nr As Long
    Dim kol As Collection, nBest As Long
    Dim sl As String, ov As String, stat As String
    Dim data As Variant, idx As Long

    Set kol = New Collection
    Set ws = SheetByName(wbk, SH_OMSCHR)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
        ws.Name = SH_OMSCHR
    End If
    If Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
        ws.Range("A1:F1").Value = Array("Sleutel", "Grootboek / nummer", "Bronomschrijving (ter info)", "Omschrijving (max 27 tekens, zonder voorvoegsel)", "Lengte", "Status")
        ws.Range("A1:F1").Font.Bold = True
    End If
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then lr = 1
    If lr >= 2 Then
        data = ws.Range(ws.Cells(2, 1), ws.Cells(lr, 6)).Value
        On Error Resume Next
        For i = 1 To lr - 1
            If CStr(data(i, 1)) <> "" Then kol.Add i, CStr(data(i, 1))
        Next i
        On Error GoTo 0
    End If
    nr = lr
    nBest = lr - 1
    For i = 1 To n
        sl = rSleutel(i)
        ov = "": stat = ""
        idx = 0
        On Error Resume Next
        idx = kol.Item(sl)
        On Error GoTo 0
        If idx > nBest Then idx = -1          'nieuwe sleutel uit deze run
        If idx > 0 Then
            ov = Trim$(CStr(data(idx, 4)))
            stat = LCase$(Trim$(CStr(data(idx, 6))))
        End If
        If ov <> "" Then
            rOms(i) = ov
            If stat Like "concept*" Then rOmsSt(i) = "concept" Else rOmsSt(i) = "vast"
        Else
            rOms(i) = rConcept(i)
            rOmsSt(i) = "concept"
            If idx = 0 Then
                nr = nr + 1
                ws.Cells(nr, 1).Value = sl
                ws.Cells(nr, 2).Value = "'" & Left$(sl, InStr(sl, "|") - 1)
                ws.Cells(nr, 3).Value = rInfo(i)
                ws.Cells(nr, 4).Value = rConcept(i)
                ws.Cells(nr, 5).Formula = "=LEN(D" & nr & ")"
                ws.Cells(nr, 6).Value = "concept"
                ws.Range(ws.Cells(nr, 1), ws.Cells(nr, 6)).Interior.Color = RGB(255, 242, 204)
                On Error Resume Next
                kol.Add nr - 1, sl
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

'==============================================================================
' Stap 3 t/m 5: tabblad 'tbv Coda'
' Geeft het aantal concept-omschrijvingen terug.
'==============================================================================
Private Function SchrijfCoda(wbk As Workbook) As Long
    Dim ws As Worksheet, i As Long, k As Long, r As Long
    Dim nTot As Long, kolMax As Long, laatste As Long
    Dim uit() As Variant, som As Double, s1() As Double, s2() As Double
    Dim prefix As String

    Set ws = SheetByName(wbk, SH_CODA)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=brn)
        ws.Name = SH_CODA
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.AutoFilterMode = False
    ws.Cells.Clear
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False
    On Error GoTo 0

    kolMax = 5 + ny
    nTot = n + 2                 'regels + saldoregel 611004 + sluitrekening
    ReDim uit(1 To nTot, 1 To kolMax)
    ReDim s1(1 To ny)
    ReDim s2(1 To ny)
    prefix = wijzStreep & " "

    'alle regels van deel 1, saldoregel 611004, dan 3a, 3b-1, 3b-2 (zoals de bron loopt)
    Dim pos As Long
    Dim saldoPlaatsing As Boolean
    pos = 0
    For i = 1 To n
        If rDeel(i) <> "1" And Not saldoPlaatsing Then
            pos = pos + 1
            VulSaldoregel uit, pos, s1
            saldoPlaatsing = True
        End If
        pos = pos + 1
        uit(pos, 1) = rGrb(i)
        uit(pos, 2) = rCat(i)
        If rVerb(i) <> "" Then uit(pos, 3) = rVerb(i)
        uit(pos, 5) = prefix & rOms(i)
        For k = 1 To ny
            If Not IsEmpty(rBed(i, k)) Then uit(pos, 5 + k) = rBed(i, k)
            If rDeel(i) = "1" And Not IsEmpty(rBed(i, k)) Then s1(k) = s1(k) + rBed(i, k)
        Next k
    Next i
    If Not saldoPlaatsing Then
        pos = pos + 1
        VulSaldoregel uit, pos, s1
    End If
    'sluitrekening
    pos = pos + 1
    uit(pos, 1) = 2999999
    For k = 1 To ny
        som = 0
        For i = 1 To pos - 1
            If Not IsEmpty(uit(i, 5 + k)) Then som = som + uit(i, 5 + k)
        Next i
        uit(pos, 5 + k) = 0 - som
    Next k
    laatste = 4 + nTot

    'opmaak vooraf (tekstformaat voor C en D, nummers voor A en B)
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = vbBlack
    ws.Cells.VerticalAlignment = xlTop
    ws.Range(ws.Cells(5, 1), ws.Cells(laatste, 2)).NumberFormat = "0"
    ws.Range(ws.Cells(5, 3), ws.Cells(laatste, 4)).NumberFormat = "@"
    ws.Range(ws.Cells(5, 6), ws.Cells(laatste, kolMax)).NumberFormat = "#,##0;-#,##0;-"
    'tekst in kolom A (balansnummers met voorloopnullen) cel voor cel als tekst
    For i = 1 To nTot
        If VarType(uit(i, 1)) = vbString Then ws.Cells(4 + i, 1).NumberFormat = "@"
    Next i

    'kopregels
    ws.Cells(1, 1).Value = "Begrotingswijziging"
    ws.Cells(1, 3).NumberFormat = "@"
    ws.Cells(1, 3).Value = wijz
    ws.Cells(4, 1).Value = "Grootboek"
    ws.Cells(4, 2).Value = "Econ.Categorie"
    ws.Cells(4, 3).Value = "Verbijzondering"
    ws.Cells(4, 4).Value = "Verdere verbijz."
    ws.Cells(4, 5).Value = "Omschrijving"
    For k = 1 To ny
        ws.Cells(2, 5 + k).Value = "Baten en lasten:"
        ws.Cells(3, 5 + k).NumberFormat = "0"
        ws.Cells(3, 5 + k).Value = yr(k)
        ws.Cells(4, 5 + k).Value = "Bedrag:"
    Next k
    ws.Range(ws.Cells(3, 6), ws.Cells(3, kolMax)).HorizontalAlignment = xlCenter
    ws.Rows(4).Font.Bold = True
    ws.Range("A:D").HorizontalAlignment = xlCenter
    ws.Range("E:E").HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(3, 6), ws.Cells(3, kolMax)).HorizontalAlignment = xlCenter

    'gegevens in een keer wegschrijven
    ws.Range(ws.Cells(5, 1), ws.Cells(laatste, kolMax)).Value = uit
    'balansnummers opnieuw als tekst vastzetten (voorloopnullen)
    For i = 1 To nTot
        If VarType(uit(i, 1)) = vbString Then ws.Cells(4 + i, 1).Value = CStr(uit(i, 1))
    Next i

    'kolombreedtes in punten
    ZetBreedtePt ws, 1, 111
    ZetBreedtePt ws, 2, 90
    ZetBreedtePt ws, 3, 90
    ZetBreedtePt ws, 4, 90
    ZetBreedtePt ws, 5, 325
    For k = 1 To ny
        ZetBreedtePt ws, 5 + k, 109
    Next k

    ws.Range("A1").Select

    'saldo's bewaren voor terugkoppeling
    mSluit = ""
    For k = 1 To ny
        mSluit = mSluit & yr(k) & ": " & Format$(uit(pos, 5 + k), "#,##0") & "   "
    Next k
    mLaatste = laatste
    mAantalRegels = nTot - 1      'inleesregels incl. saldoregel 611004, excl. sluitrekening

    'aantal concepten
    Dim nc As Long
    For i = 1 To n
        If rOmsSt(i) = "concept" Then nc = nc + 1
    Next i
    SchrijfCoda = nc
End Function

Private Sub VulSaldoregel(ByRef uit() As Variant, ByVal pos As Long, ByRef s1() As Double)
    Dim k As Long
    uit(pos, 1) = 611004
    uit(pos, 2) = 43812
    uit(pos, 5) = wijzStreep & " Stelpost nieuw beleid saldo"
    For k = 1 To ny
        uit(pos, 5 + k) = 0 - s1(k)
    Next k
End Sub

'==============================================================================
' Stap 6: tabblad 'controle inlees' (formules naar bron en tbv Coda)
'==============================================================================
Private Sub SchrijfControle(wbk As Workbook, ByVal nConcept As Long)
    Dim ws As Worksheet, wsC As Worksheet
    Dim k As Long, i As Long, r As Long, d As Long
    Dim bq As String, cq As String, last As Long, sluit As Long
    Dim salRij(1 To 4) As Long, kenm(1 To 4) As String, naam(1 To 4) As String
    Dim yc As String, f As String
    Dim hulp() As Variant

    Set wsC = SheetByName(wbk, SH_CODA)
    Set ws = SheetByName(wbk, SH_CONTROLE)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=wsC)
        ws.Name = SH_CONTROLE
    End If
    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10

    bq = "'" & brn.Name & "'!"
    cq = "'" & SH_CODA & "'!"
    last = mLaatste
    sluit = last

    salRij(1) = rSal1: kenm(1) = "1": naam(1) = "Deel 1 Mutaties baten en lasten"
    salRij(2) = rSal3a: kenm(2) = "3a": naam(2) = "Deel 3a Investeringen"
    salRij(3) = rSalRes: kenm(3) = "3b1": naam(3) = "Deel 3b-1 Reserves"
    salRij(4) = rSalVrz: kenm(4) = "3b2": naam(4) = "Deel 3b-2 Voorzieningen"

    ws.Range("A1").Value = "Controle inleesbestand"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "Bronblad"
    ws.Range("B2").Value = brn.Name
    ws.Range("A3").Value = "Status"
    ws.Range("A3").Font.Bold = True
    ws.Range("A5").Value = "Onderdeel"
    ws.Range("B5").Value = "Regel"
    For k = 1 To ny
        ws.Cells(5, 2 + k).NumberFormat = "0"
        ws.Cells(5, 2 + k).Value = yr(k)
    Next k
    ws.Range(ws.Cells(5, 1), ws.Cells(5, 2 + ny)).Font.Bold = True
    ws.Range(ws.Cells(5, 3), ws.Cells(5, 2 + ny)).HorizontalAlignment = xlRight

    'hulpkolommen: bronrij, deelkenmerk (tekstformaat), omschrijvingsstatus; rij = rij op tbv Coda
    ws.Cells(4, HULPKOL - 1).Value = "Bronrij"
    ws.Cells(4, HULPKOL).Value = "Deelkenmerk (hulpkolom)"
    ws.Cells(4, HULPKOL + 1).Value = "Omschrijving"
    ws.Range(ws.Cells(4, HULPKOL - 1), ws.Cells(4, HULPKOL + 1)).Font.Bold = True
    ws.Range(ws.Cells(5, HULPKOL), ws.Cells(last, HULPKOL)).NumberFormat = "@"
    ReDim hulp(1 To last - 4, 1 To 3)
    Dim pos As Long, saldoGezet As Boolean
    pos = 0
    For i = 1 To n
        If rDeel(i) <> "1" And Not saldoGezet Then
            pos = pos + 1: hulp(pos, 1) = "saldo": hulp(pos, 2) = "1": hulp(pos, 3) = "saldoregel 611004"
            saldoGezet = True
        End If
        pos = pos + 1
        hulp(pos, 1) = rRij(i)
        hulp(pos, 2) = rDeel(i)
        hulp(pos, 3) = rOmsSt(i)
    Next i
    If Not saldoGezet Then
        pos = pos + 1: hulp(pos, 1) = "saldo": hulp(pos, 2) = "1": hulp(pos, 3) = "saldoregel 611004"
    End If
    pos = pos + 1
    hulp(pos, 1) = "sluit": hulp(pos, 2) = "": hulp(pos, 3) = "sluitrekening 2999999"
    ws.Range(ws.Cells(5, HULPKOL - 1), ws.Cells(last, HULPKOL + 1)).Value = hulp
    'deelkenmerk als tekst vastzetten
    For i = 1 To pos
        ws.Cells(4 + i, HULPKOL).Value = CStr(hulp(i, 2))
    Next i

    'A. aansluiting per deel en per jaar
    ws.Range("A6").Value = "A. Aansluiting per deel en per jaar"
    ws.Range("A6").Font.Bold = True
    Dim hk As String
    hk = "$" & ColL(HULPKOL) & "$5:$" & ColL(HULPKOL) & "$" & last
    For d = 1 To 4
        r = 7 + (d - 1) * 3
        ws.Cells(r, 1).Value = naam(d)
        ws.Cells(r, 2).Value = "Bedrag volgens de bron (saldoregel, absoluut)"
        ws.Cells(r + 1, 2).Value = "Bedrag volgens het inleesbestand (absoluut)"
        ws.Cells(r + 2, 2).Value = "Verschil"
        ws.Cells(r + 2, 2).Font.Bold = True
        For k = 1 To ny
            yc = ColL(2 + k)
            ws.Cells(r, 2 + k).Formula = "=ABS(N(" & bq & ColL(4 + 2 * k) & salRij(d) & "))"
            f = cq & ColL(5 + k) & "$5:" & ColL(5 + k) & "$" & last
            If d = 1 Then
                ws.Cells(r + 1, 2 + k).Formula = "=ABS(SUMIFS(" & f & "," & hk & ",""1""," & cq & "$A$5:$A$" & last & ",""<>611004""))"
            Else
                ws.Cells(r + 1, 2 + k).Formula = "=ABS(SUMIFS(" & f & "," & hk & ",""" & kenm(d) & """))"
            End If
            ws.Cells(r + 2, 2 + k).Formula = "=ROUND(" & yc & r & "-" & yc & (r + 1) & ",0)"
        Next k
    Next d

    'B. saldocontrole per groep
    ws.Range("A20").Value = "B. Saldocontrole per groep (beide moeten 0 zijn)"
    ws.Range("A20").Font.Bold = True
    ws.Range("B21").Value = "Exploitatie: grootboek 600000 t/m 699999 incl. 611004"
    ws.Range("B22").Value = "Alle regels incl. sluitrekening 2999999"
    ws.Range("B23").Value = "Info: sluitrekening plus (3a + 3b-1 + 3b-2), moet 0 zijn"
    For k = 1 To ny
        f = cq & ColL(5 + k) & "$5:" & ColL(5 + k) & "$" & last
        ws.Cells(21, 2 + k).Formula = "=SUMIFS(" & f & "," & cq & "$A$5:$A$" & last & ","">=600000""," & cq & "$A$5:$A$" & last & ",""<700000"")"
        ws.Cells(22, 2 + k).Formula = "=SUM(" & cq & ColL(5 + k) & "5:" & ColL(5 + k) & sluit & ")"
        ws.Cells(23, 2 + k).Formula = "=" & cq & ColL(5 + k) & sluit & "+SUMIFS(" & f & "," & hk & ",""3a"")+SUMIFS(" & f & "," & hk & ",""3b1"")+SUMIFS(" & f & "," & hk & ",""3b2"")"
    Next k

    'C. statusregel
    Dim rng As String, st As String, c1 As String, c2 As String
    c1 = ColL(3): c2 = ColL(2 + ny)
    st = "=IF(AND("
    For d = 1 To 4
        st = st & "SUMPRODUCT(ABS(" & c1 & (9 + (d - 1) * 3) & ":" & c2 & (9 + (d - 1) * 3) & "))=0,"
    Next d
    st = st & "SUMPRODUCT(ABS(" & c1 & "21:" & c2 & "21))=0,SUMPRODUCT(ABS(" & c1 & "22:" & c2 & "22))=0),""OK - sluit aan op " & brn.Name & """,""LET OP - verschil, niet inlezen"")"
    ws.Range("B3").Formula = st
    ws.Range("B3").Font.Bold = True

    'D. omschrijvingen en overgeslagen regels (vaste waarden, uit de macro)
    Dim nTeLang As Long, nOhv As Long, tel As Long
    For i = 1 To n
        If Len(wijzStreep & " " & rOms(i)) > MAX_OMS Then nTeLang = nTeLang + 1
        If InStr(1, rOms(i), "OHV", vbTextCompare) > 0 And Not HeeftONummer(rOms(i)) Then nOhv = nOhv + 1
    Next i
    ws.Range("A25").Value = "C. Omschrijvingen en overgeslagen regels"
    ws.Range("A25").Font.Bold = True
    ws.Range("B26").Value = "Omschrijvingen langer dan " & MAX_OMS & " tekens (moet 0 zijn)"
    ws.Range("C26").Value = nTeLang
    ws.Range("B27").Value = "OHV zonder O-nummer (moet 0 zijn)"
    ws.Range("C27").Value = nOhv
    ws.Range("B28").Value = "Omschrijvingen als CONCEPT, nakijken op tabblad '" & SH_OMSCHR & "' (moet 0 zijn)"
    ws.Range("C28").Value = nConcept
    ws.Range("B29").Value = "Aantal inleesregels (incl. saldoregel 611004, excl. sluitrekening)"
    ws.Range("C29").Value = mAantalRegels
    ws.Range("B30").Value = "Regel in het inleesbestand komt uit bronrij: zie kolom " & ColL(HULPKOL - 1)
    r = 31
    Dim it As Variant
    For Each it In mInfo
        ws.Cells(r, 2).Value = CStr(it)
        r = r + 1
    Next it
    ws.Cells(r, 2).Value = "Regels zonder bedrag, kop-, totaal- en saldoregels (behalve 611004) leveren geen inleesregel op."

    'opmaak
    ws.Range(ws.Cells(7, 3), ws.Cells(23, 2 + ny)).NumberFormat = "#,##0;-#,##0;-"
    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 62
    For k = 1 To ny
        ws.Columns(2 + k).ColumnWidth = 14
    Next k
    ws.Columns(HULPKOL - 1).ColumnWidth = 9
    ws.Columns(HULPKOL).ColumnWidth = 22
    ws.Columns(HULPKOL + 1).ColumnWidth = 24
    On Error Resume Next
    With ws.Range("B3").FormatConditions
        .Delete
        With .Add(Type:=xlExpression, Formula1:="=LEFT($B$3,2)=""OK""")
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
        End With
        With .Add(Type:=xlExpression, Formula1:="=LEFT($B$3,2)<>""OK""")
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
    End With
    On Error GoTo 0
End Sub

'==============================================================================
' Terugkoppeling (stap 7)
'==============================================================================
Private Sub Terugkoppeling(ByVal st As String, ByVal nConcept As Long)
    Dim t As String, k As Long
    t = "Bronblad: " & brn.Name & vbCrLf
    t = t & "Wijzigingsnummer: " & wijz & " (" & wijzStreep & ")" & vbCrLf
    t = t & "Jaren: "
    For k = 1 To ny
        t = t & yr(k) & IIf(k < ny, ", ", "")
    Next k
    t = t & vbCrLf & "Inleesregels: " & mAantalRegels & " (incl. saldoregel 611004, excl. sluitrekening)" & vbCrLf
    t = t & "Sluitrekening 2999999 per jaar: " & Trim$(mSluit) & vbCrLf
    t = t & "Status controleblad: " & st & vbCrLf
    t = t & "Overgeslagen met tekst in bedragkolom (pm e.d.): " & mInfo.Count & vbCrLf
    If nConcept > 0 Then t = t & vbCrLf & nConcept & " omschrijving(en) zijn CONCEPT. Loop ze na op tabblad '" & SH_OMSCHR & "' (gele regels), zet de status op 'vast' en druk opnieuw op de knop." & vbCrLf
    If Left$(st, 2) <> "OK" Then t = t & vbCrLf & "LET OP: het bestand is NIET inleesklaar. Bekijk de verschillen op '" & SH_CONTROLE & "'." & vbCrLf
    gResultaat = t
    If Not mStil Then
        If Left$(st, 2) = "OK" Then
            MsgBox t, vbInformation, "Inleesbestand " & wijz
        Else
            MsgBox t, vbExclamation, "Inleesbestand " & wijz
        End If
    End If
End Sub

'==============================================================================
' Hulpfuncties
'==============================================================================
Private Sub Meld_stop(ByVal titel As String)
    Dim t As String, it As Variant, i As Long, ws As Worksheet, wbk As Workbook
    Set wbk = ActiveWorkbook
    t = titel & ": er is GEEN inleesbestand gemaakt. De bron moet eerst worden aangepast (het bronblad wordt door deze macro nooit gewijzigd)." & vbCrLf & vbCrLf
    i = 0
    For Each it In mMeld
        i = i + 1
        If i <= 25 Then t = t & "- " & CStr(it) & vbCrLf
    Next it
    If mMeld.Count > 25 Then t = t & "... en " & (mMeld.Count - 25) & " meldingen meer (zie tabblad '" & SH_MELD & "')." & vbCrLf
    gResultaat = t
    'volledige lijst op een apart tabblad (bestaande tbv Coda en controle blijven ongemoeid)
    Set ws = SheetByName(wbk, SH_MELD)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
        ws.Name = SH_MELD
    Else
        ws.Cells.Clear
    End If
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Cells(1, 1).Value = titel & ": geen inleesbestand gemaakt. Pas de bron aan en druk opnieuw op de knop."
    ws.Cells(1, 1).Font.Bold = True
    i = 2
    For Each it In mMeld
        ws.Cells(i, 1).Value = CStr(it)
        i = i + 1
    Next it
    ws.Columns(1).ColumnWidth = 140
    If Not mStil Then MsgBox t, vbExclamation, "Inleesbestand " & wijz
End Sub

Private Sub VerwijderMeldingenBlad(wbk As Workbook)
    Dim ws As Worksheet
    Set ws = SheetByName(wbk, SH_MELD)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub Zeg(ByVal t As String, ByVal stijl As VbMsgBoxStyle)
    gResultaat = t
    If Not mStil Then MsgBox t, stijl, "Inleesbestand"
End Sub

Private Sub AddMeld(ByVal t As String)
    mMeld.Add t
End Sub

Private Function SheetByName(wbk As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wbk.Worksheets
        If LCase$(ws.Name) = LCase$(nm) Then
            Set SheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function Txt(ByVal x As Variant) As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    If IsNull(x) Then Exit Function
    Txt = Trim$(Replace(CStr(x), Chr$(160), " "))
End Function

Private Function IsNum(ByVal x As Variant) As Boolean
    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNum = True
    End Select
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    AllDigits = (Len(s) > 0)
    If AllDigits Then AllDigits = Not (s Like "*[!0-9]*")
End Function

Private Function IsKopNr(ByVal a As String, ByVal letter As String) As Boolean
    If Len(a) >= 2 Then
        If UCase$(Left$(a, 1)) = letter Then IsKopNr = AllDigits(Mid$(a, 2))
    End If
End Function

Private Function IsWijzNr(ByVal t As String) As Boolean
    Dim p As Long
    If Len(t) < 4 Then Exit Function
    If UCase$(Left$(t, 1)) <> "W" Then Exit Function
    p = InStr(t, "/")
    If p < 3 Then Exit Function
    IsWijzNr = AllDigits(Mid$(t, 2, p - 2)) And AllDigits(Mid$(t, p + 1))
End Function

Private Function HeeftONummer(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s) - 1
        If UCase$(Mid$(s, i, 1)) = "O" And Mid$(s, i + 1, 1) Like "#" Then
            If i = 1 Then
                HeeftONummer = True
            ElseIf Mid$(s, i - 1, 1) = " " Then
                HeeftONummer = True
            End If
            If HeeftONummer Then Exit Function
        End If
    Next i
End Function

Private Function ColL(ByVal c As Long) As String
    Dim ad As String
    ad = brn.Cells(1, c).Address(False, False)
    ColL = Left$(ad, Len(ad) - 1)
End Function

Private Function NormTxt(ByVal s As String) As String
    Dim i As Long, ch As String, o As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            o = o & ch
            If Len(o) >= 50 Then Exit For
        End If
    Next i
    NormTxt = o
End Function

Private Sub ZetBreedtePt(ws As Worksheet, ByVal col As Long, ByVal pt As Double)
    Dim i As Long, w As Double
    ws.Columns(col).ColumnWidth = pt / 5.5
    For i = 1 To 8
        w = ws.Columns(col).Width
        If w <= 0 Then Exit For
        If Abs(w - pt) < 0.6 Then Exit For
        ws.Columns(col).ColumnWidth = ws.Columns(col).ColumnWidth * pt / w
    Next i
End Sub

'Voorlopige omschrijving voor regels die nog niet in tabblad 'omschrijvingen' staan.
'Let op: dit is een CONCEPT (afkorten, geen samenvatting). Altijd handmatig nalopen.
Private Function ConceptOms(ByVal s As String) As String
    Dim ab As Variant, i As Long, p As Long, t As String
    ab = Array("onderhoudsvoorziening", "OHV", "BTW-compensatiefonds", "bcf", "BTW-cf", "bcf", "opbrengsten", "opbr", "opbrengst", "opbr", _
        "onvoorziene uitgaven", "onvoorz", "onvoorzien", "onvoorz", "voorzieningen", "vrz", "voorziening", "vrz", "reserves", "res", "reserve", "res", _
        "investeringen", "invest", "investering", "invest", "exploitatie", "expl", "meicirculaire", "Meicirc", "rijksvergoeding", "Rijksverg", _
        "rijksmiddelen", "Rijksmidd", "rijksbijdrage", "Rijksbijdr", "actualisatie", "Actual", "correctie primitieve begroting", "Corr prim begr", _
        "huisvesting", "huisv", "verduurzaming", "verduurz", "openbare ruimte", "OR", "sociale basis", "soc basis", "hogere", "Hoger", "lagere", "Lager", _
        "verschuiving", "Verschuiv")
    t = Trim$(s)
    p = InStr(t, ". ")
    If p > 0 Then t = Left$(t, p - 1)
    For i = LBound(ab) To UBound(ab) Step 2
        t = Replace(t, ab(i), ab(i + 1), 1, -1, vbTextCompare)
    Next i
    If Len(t) > 27 Then
        t = Left$(t, 27)
        p = InStrRev(t, " ")
        If p > 10 Then t = Left$(t, p - 1)
    End If
    If Len(t) > 0 Then t = UCase$(Left$(t, 1)) & Mid$(t, 2)
    ConceptOms = t
End Function
