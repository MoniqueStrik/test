Attribute VB_Name = "modVerschillen"
Option Explicit
'==============================================================================
' Verschillenanalyses 3e financiele rapportage
'
' Knop / macro : MaakOfWerkVerschillenanalyseBij
'   Keuze 1 "Bijwerken"            : bestaande analyse blijft staan; nieuwe, verdwenen
'                                    en gewijzigde bronregels worden verwerkt (stap 9).
'   Keuze 2 "Verschillenanalyse maken" : tabbladen volledig opnieuw opbouwen (stap 0 t/m 8).
'
' Tabbladen : verschillenanalyses firap (presentatie), hulp verschillenanalyse
'             (ankers en controlesommen), Ontwikkeling begrotingssaldo (sectie D).
' Het bronblad "3e firap <jaar>" wordt nooit gewijzigd. Alle bedragen zijn live
' formules naar het bronblad. Rijnummers worden altijd op tekst gezocht.
' Alleen ASCII-tekens in deze module (euro- en trema-tekens via ChrW).
'==============================================================================

Private Const SH_PRES As String = "verschillenanalyses firap"
Private Const SH_HULP As String = "hulp verschillenanalyse"
Private Const SH_ONTW As String = "Ontwikkeling begrotingssaldo"
Private Const GRENS_STD As Double = 50000

Private v As Variant
Private brn As Worksheet
Private bq As String
Private jaar As Long
Private lastR As Long
Private aE1 As Long, aE2 As Long, aSal As Long
Private aI1 As Long, aI2 As Long, aR1 As Long, aR2 As Long, aV1 As Long, aV2 As Long
Private grens As Double
Private mStil As Boolean
Private eur As String, ee As String
Public gResultaatV As String
Private mRap As Collection
Private rapN As String

'per bronrij
Private rKey() As String
Private rKop() As String
Private rKopNm() As String
Private linkN() As Long

'bestaande omschrijvingen (sleutel > tekst)
Private harv As Collection

'groepen (presentatieregels in wording)
Private gN As Long
Private gBlok() As Long          '1 incidenteel, 2 structureel, 3 overhevelingen
Private gVal() As Double
Private gTerm() As String        'formule zonder "="
Private gDesc() As String
Private gHard0() As Boolean
Private gSmall() As Boolean
Private gKey() As String
Private gRij() As Long
Private gKind() As Long          '0 gewoon, 1 budgetneutraal paar, 2 lijst-regel
Private gPart() As Long

'structuur presentatieblad
Private pHdr(1 To 3) As Long, pSub(1 To 3) As Long, pTot As Long
Private pInv As Long, pRes As Long, pVz As Long, pCtl As Long
Private pInvTot As Long, pResTot As Long, pVzTot As Long

'==============================================================================
' Ingangen
'==============================================================================
Public Sub MaakOfWerkVerschillenanalyseBij()
    Dim k As VbMsgBoxResult
    mStil = False
    k = MsgBox("Wat wilt u doen?" & vbCrLf & vbCrLf & _
        "Ja = Alleen bijwerken: de bestaande analyse blijft staan, nieuwe en verdwenen bronregels worden verwerkt." & vbCrLf & _
        "Nee = Verschillenanalyse maken: de tabbladen worden volledig opnieuw opgebouwd." & vbCrLf & _
        "Annuleren = stoppen.", vbYesNoCancel + vbQuestion, "Verschillenanalyses firap")
    If k = vbYes Then Voer_uitV False
    If k = vbNo Then
        If MsgBox("De inhoud van '" & SH_PRES & "' en '" & SH_HULP & "' wordt gewist en opnieuw opgebouwd. Bestaande omschrijvingen worden zoveel mogelijk hergebruikt." & vbCrLf & vbCrLf & "Doorgaan?", vbYesNo + vbQuestion, "Verschillenanalyse maken") = vbYes Then Voer_uitV True
    End If
End Sub

Public Sub VerschillenBijwerkenStil()
    mStil = True
    Voer_uitV False
End Sub

Public Sub VerschillenVolledigStil()
    mStil = True
    Voer_uitV True
End Sub

Private Sub Voer_uitV(ByVal volledig As Boolean)
    Dim wbk As Workbook, ok As Boolean
    On Error GoTo Fout
    Set wbk = ActiveWorkbook
    Set mRap = New Collection
    eur = ChrW(8364): ee = ChrW(235)
    gResultaatV = ""
    Application.ScreenUpdating = False
    If Not LeesBronV(wbk) Then GoTo Meld
    grens = GRENS_STD
    If volledig Then ok = BouwVolledig(wbk) Else ok = Bijwerken(wbk)
    If Not ok Then GoTo Meld
    SectieD wbk
    Application.Calculate
    Afsluiting wbk, volledig
    GoTo Einde
Meld:
    gResultaatV = RapTekst("Geen wijzigingen aangebracht.")
    If Not mStil Then MsgBox gResultaatV, vbExclamation, "Verschillenanalyses firap"
    GoTo Einde
Fout:
    gResultaatV = "Fout " & Err.Number & ": " & Err.Description
    If Not mStil Then MsgBox gResultaatV, vbCritical, "Verschillenanalyses firap"
Einde:
    On Error Resume Next
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' Bronblad inlezen en ankers zoeken (op tekst, nooit op rijnummer)
'==============================================================================
Private Function DashWaarde(ByVal adres As String) As String
    Dim d As Worksheet
    Set d = SheetByNameV(ActiveWorkbook, "Inlees en analyse")
    If d Is Nothing Then Exit Function
    DashWaarde = Trim$(CStr(d.Range(adres).Value))
End Function

Private Sub DashLog(ByVal t As String)
    Dim d As Worksheet
    Set d = SheetByNameV(ActiveWorkbook, "Inlees en analyse")
    If d Is Nothing Then Exit Sub
    d.Range("D16").Value = Format$(Now, "dd-mm-yyyy hh:mm") & "  " & t
End Sub

Private Function LeesBronV(wbk As Workbook) As Boolean
    Dim ws As Worksheet, k As Long, g As Worksheet, r As Long, i As Long
    Dim a As String, c As String, nm As String, s As String, dn As String
    dn = DashWaarde("C4")
    If dn <> "" And DashWaarde("C5") <> "" Then
        dn = dn & " " & DashWaarde("C5")
        Set g = SheetByNameV(wbk, dn)
        If g Is Nothing Then Rap "Bronblad '" & dn & "' (gekozen op het dashboard) bestaat niet. Pas de keuze aan of gebruik de knop Overgang.": Exit Function
        k = 1
    Else
        For Each ws In wbk.Worksheets
            If LCase$(ws.Name) Like "[123]e*firap*" And LCase$(ws.Name) <> "verschillenanalyses firap" Then k = k + 1: Set g = ws
        Next ws
    End If
    If k = 0 Then Rap "Geen bronblad gevonden: een tabblad met '1e', '2e' of '3e firap' en het jaartal in de naam.": Exit Function
    If k > 1 Then
        If LCase$(ActiveSheet.Name) Like "[123]e*firap*" Then
            Set g = ActiveSheet
        Else
            Rap "Er zijn " & k & " bronbladen '3e firap'. Activeer het juiste blad en start opnieuw.": Exit Function
        End If
    End If
    Set brn = g
    rapN = Left$(g.Name, 1)
    bq = "'" & brn.Name & "'!"
    nm = brn.Name
    For i = 1 To Len(nm) - 3
        s = Mid$(nm, i, 4)
        If s Like "20##" Then jaar = CLng(s)
    Next i
    If jaar = 0 Then Rap "Geen jaartal gevonden in de naam van het bronblad.": Exit Function
    lastR = brn.UsedRange.Row + brn.UsedRange.Rows.Count - 1
    If lastR < 20 Then lastR = 20
    v = brn.Range(brn.Cells(1, 1), brn.Cells(lastR, 19)).Value
    aE1 = 0: aE2 = 0: aSal = 0: aI1 = 0: aI2 = 0: aR1 = 0: aR2 = 0: aV1 = 0: aV2 = 0
    For r = 1 To lastR
        a = LCase$(TxtV(v(r, 1)))
        c = LCase$(TxtV(v(r, 3)))
        If aE1 = 0 And a = "grootboek" Then aE1 = r + 1
        If aE1 > 0 And aSal = 0 And c Like "salderingsrekening begrotingssaldo*" Then aSal = r: aE2 = r - 1
        If aI1 = 0 And a Like "3a. investeringen*" Then aI1 = r + 1
        If aI2 = 0 And a Like "totaal 3a*" Then aI2 = r - 1
        If aR1 = 0 And a = "reserves" Then aR1 = r + 1
        If aR2 = 0 And a Like "totaal 3b-1*" Then aR2 = r - 1
        If aV1 = 0 And a = "voorzieningen" Then aV1 = r + 1
        If aV2 = 0 And a Like "totaal 3b-2*" Then aV2 = r - 1
    Next r
    If aE1 = 0 Then Rap "Anker 'Grootboek' (kolom A) niet gevonden."
    If aSal = 0 Then Rap "Anker 'Salderingsrekening begrotingssaldo' (kolom C) niet gevonden."
    If aI1 = 0 Or aI2 = 0 Then Rap "Anker '3a. Investeringen' of 'Totaal 3a' niet gevonden."
    If aR1 = 0 Or aR2 = 0 Then Rap "Anker 'Reserves' of 'Totaal 3b-1' niet gevonden."
    If aV1 = 0 Or aV2 = 0 Then Rap "Anker 'Voorzieningen' of 'Totaal 3b-2' niet gevonden."
    If mRap.Count > 0 Then Exit Function

    'sleutel en kopregel per bronrij (voor hergebruik van omschrijvingen)
    ReDim rKey(1 To lastR): ReDim rKop(1 To lastR): ReDim rKopNm(1 To lastR): ReDim linkN(1 To lastR)
    KopScan aE1, aE2, 6, ""
    KopScan aI1, aI2, 7, ""
    KopScan aR1, aR2, 0, "R"
    KopScan aV1, aV2, 0, "O"
    LeesBronV = True
End Function

Private Sub KopScan(ByVal r1 As Long, ByVal r2 As Long, ByVal lenKop As Long, ByVal letter As String)
    Dim r As Long, a As String, kop As String, nm As String, nb As String, k As Long
    For r = r1 To r2
        a = TxtV(v(r, 1))
        If Not HeeftBedragV(r) Then
            If lenKop > 0 Then
                If Len(a) = lenKop And AllDigitsV(a) Then kop = a: nm = TxtV(v(r, 3))
            Else
                If Len(a) >= 2 Then
                    If UCase$(Left$(a, 1)) = letter And AllDigitsV(Mid$(a, 2)) Then kop = a: nm = TxtV(v(r, 3))
                End If
            End If
        End If
        rKop(r) = kop: rKopNm(r) = nm
        nb = UCase$(TxtV(v(r, 5)))
        If HeeftBedragV(r) Then rKey(r) = kop & "|" & nb & "|" & NormV(TxtV(v(r, 3)))
    Next r
End Sub

Private Function HeeftBedragV(ByVal r As Long) As Boolean
    If IsNumV(v(r, 6)) Then HeeftBedragV = (v(r, 6) <> 0)
End Function

'tekst van de bronregel inclusief doorlopende C-cellen eronder
Private Function VolTekst(ByVal r As Long) As String
    Dim t As String, k As Long
    t = TxtV(v(r, 3))
    k = r + 1
    Do While k <= lastR
        If TxtV(v(k, 1)) <> "" Or IsNumV(v(k, 6)) Or TxtV(v(k, 3)) = "" Then Exit Do
        t = t & " " & TxtV(v(k, 3))
        k = k + 1
    Loop
    VolTekst = t
End Function

'==============================================================================
' Hulpblad (stap 1)
'==============================================================================
Private Sub SchrijfHulpV(wbk As Workbook, ByVal volledig As Boolean)
    Dim ws As Worksheet, g As Variant
    Set ws = SheetByNameV(wbk, SH_HULP)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(After:=wbk.Worksheets(wbk.Worksheets.Count))
        ws.Name = SH_HULP
    End If
    g = ws.Range("B4").Value
    If IsNumV(g) Then grens = CDbl(g) Else grens = GRENS_STD
    If volledig Then ws.Cells.Clear
    ws.Range("A1").Value = "Hulpblad verschillenanalyses " & ChrW(8211) & " 3e firap " & jaar
    ws.Range("A1").Value = "Hulpblad verschillenanalyses, " & rapN & "e firap " & jaar
    ws.Range("A3").Value = "Instellingen"
    ws.Range("A4").Value = "Rapportagegrens exploitatie (" & eur & ")": ws.Range("B4").Value = grens
    ws.Range("A5").Value = "Begrotingsjaar": ws.Range("B5").Value = jaar
    ws.Range("A6").Value = "Bronblad": ws.Range("B6").Value = brn.Name
    ws.Range("A7").Value = "Exploitatie eerste mutatierij": SetF ws.Range("B7"), "=MATCH(""Grootboek""," & bq & "$A:$A,0)+1"
    ws.Range("A8").Value = "Exploitatie laatste mutatierij": SetF ws.Range("B8"), "=MATCH(""Salderingsrekening begrotingssaldo*""," & bq & "$C:$C,0)-1"
    ws.Range("A9").Value = "Volgend begrotingsjaar": SetF ws.Range("B9"), "=B5+1"
    ws.Range("A10").Value = "Investeringen eerste rij": SetF ws.Range("B10"), "=MATCH(""3a. Investeringen*""," & bq & "$A:$A,0)+1"
    ws.Range("A11").Value = "Investeringen laatste rij": SetF ws.Range("B11"), "=MATCH(""Totaal 3a*""," & bq & "$A:$A,0)-1"
    ws.Range("A12").Value = "Reserves eerste rij": SetF ws.Range("B12"), "=MATCH(""Reserves""," & bq & "$A:$A,0)+1"
    ws.Range("A13").Value = "Reserves laatste rij": SetF ws.Range("B13"), "=MATCH(""Totaal 3b-1*""," & bq & "$A:$A,0)-1"
    ws.Range("A14").Value = "Voorzieningen eerste rij": SetF ws.Range("B14"), "=MATCH(""Voorzieningen""," & bq & "$A:$A,0)+1"
    ws.Range("A15").Value = "Voorzieningen laatste rij": SetF ws.Range("B15"), "=MATCH(""Totaal 3b-2*""," & bq & "$A:$A,0)-1"
    ws.Range("A16").ClearContents: ws.Range("B16").ClearContents
    ws.Range("A17").Value = "Controlesommen bron (tekenconventie per sectie)"
    ws.Range("A18").Value = "Exploitatie bron (N positief, " & eur & ")"
    SetF ws.Range("B18"), "=SUMIF(" & Rg("E", aE1, aE2) & ",""N""," & Rg("F", aE1, aE2) & ")-SUMIF(" & Rg("E", aE1, aE2) & ",""B""," & Rg("F", aE1, aE2) & ")"
    ws.Range("A19").Value = "Exploitatie saldoregel bronblad (teken omgekeerd: tegenboeking)"
    SetF ws.Range("B19"), "=IF(TRIM(" & bq & "$E$" & aSal & ")=""N"",-1,1)*" & bq & "$F$" & aSal
    ws.Range("A20").Value = "Investeringen bron (N positief, " & eur & ")"
    SetF ws.Range("B20"), "=SUMIF(" & Rg("E", aI1, aI2) & ",""N""," & Rg("F", aI1, aI2) & ")-SUMIF(" & Rg("E", aI1, aI2) & ",""B""," & Rg("F", aI1, aI2) & ")"
    ws.Range("A21").Value = "Reserves bron (B positief, " & eur & ")"
    SetF ws.Range("B21"), "=SUMIFS(" & Rg("F", aR1, aR2) & "," & Rg("E", aR1, aR2) & ",""B""," & Rg("R", aR1, aR2) & ",""R*"")-SUMIFS(" & Rg("F", aR1, aR2) & "," & Rg("E", aR1, aR2) & ",""N""," & Rg("R", aR1, aR2) & ",""R*"")"
    ws.Range("A22").Value = "Voorzieningen bron (B positief, " & eur & ")"
    SetF ws.Range("B22"), "=SUMIFS(" & Rg("F", aV1, aV2) & "," & Rg("E", aV1, aV2) & ",""B""," & Rg("R", aV1, aV2) & ",""O*"")-SUMIFS(" & Rg("F", aV1, aV2) & "," & Rg("E", aV1, aV2) & ",""N""," & Rg("R", aV1, aV2) & ",""O*"")"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Font.Bold = True
    ws.Range("A17").Font.Bold = True
    ws.Columns(1).ColumnWidth = 60
    ws.Columns(2).ColumnWidth = 16
    ws.Range("B4").NumberFormat = "#,##0"
    ws.Range("B5").NumberFormat = "0"
    ws.Range("B9").NumberFormat = "0"
    ws.Range("B18:B22").NumberFormat = "#,##0;-#,##0;-"
End Sub

Private Function Rg(ByVal kol As String, ByVal r1 As Long, ByVal r2 As Long) As String
    Rg = bq & "$" & kol & "$" & r1 & ":$" & kol & "$" & r2
End Function

'==============================================================================
' Omschrijvingen: hergebruik uit bestaande presentatie, anders kolom Q, anders concept
'==============================================================================
Private Sub OogstOmschrijvingen(wbk As Workbook)
    Dim ws As Worksheet, r As Long, lr As Long, f As String, n As Long, a As String
    Set harv = New Collection
    Set ws = SheetByNameV(wbk, SH_PRES)
    If ws Is Nothing Then Exit Sub
    lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To lr
        f = ""
        On Error Resume Next
        f = ws.Cells(r, 2).Formula
        On Error GoTo 0
        a = ""
        If Not ws.Cells(r, 1).HasFormula Then a = Trim$(CStr(ws.Cells(r, 1).Value))
        If Left$(f, 1) = "=" And a <> "" Then
            n = EersteRef(f)
            If n >= 1 And n <= lastR Then
                If rKey(n) <> "" Then
                    On Error Resume Next
                    harv.Add a, rKey(n)
                    On Error GoTo 0
                End If
            End If
        End If
    Next r
End Sub

Private Function Oms(ByVal r As Long) As String
    Dim t As String, q As String
    t = ""
    If Not harv Is Nothing And rKey(r) <> "" Then
        On Error Resume Next
        t = harv.Item(rKey(r))
        On Error GoTo 0
    End If
    If t <> "" Then Oms = t: Exit Function
    q = TxtV(v(r, 17))
    If q <> "" Then Oms = q: Exit Function
    Oms = Kort(VolTekst(r))
    If Oms = "" Then Oms = rKopNm(r)
    Rap "Concept-omschrijving (nakijken): bronrij " & r & " > " & Oms
End Function

Private Function Kort(ByVal s As String) As String
    Dim t As String, p As Long
    t = Trim$(s)
    p = InStr(t, ". ")
    If p > 0 Then t = Left$(t, p - 1)
    If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1)
    If Len(t) > 90 Then
        t = Left$(t, 90)
        p = InStrRev(t, " ")
        If p > 40 Then t = Left$(t, p - 1)
    End If
    Kort = t
End Function

'==============================================================================
' Exploitatie: groepen vormen (blok, bedrag, formule, omschrijving)
' alleenOnverbonden = True bij bijwerken: alleen bronrijen zonder koppeling
'==============================================================================
Private Sub BouwGroepen(ByVal alleenOnverbonden As Boolean)
    Dim r As Long, i As Long, j As Long, n As Long, m As Long
    Dim ok() As Boolean, gpair() As Long, lijst() As Boolean, nb() As String, di() As String
    Dim sg() As Double, gb() As String, txt() As String, idx() As Long
    Dim p As Long, s As Double, tm As String, blok As Long, groupOf() As Long
    ReDim ok(1 To lastR): ReDim gpair(1 To lastR): ReDim lijst(1 To lastR): ReDim nb(1 To lastR): ReDim di(1 To lastR)
    ReDim sg(1 To lastR): ReDim gb(1 To lastR): ReDim txt(1 To lastR): ReDim groupOf(1 To lastR)
    Dim cap As Long: cap = lastR + 5
    ReDim gBlok(1 To cap): ReDim gVal(1 To cap): ReDim gTerm(1 To cap): ReDim gDesc(1 To cap)
    ReDim gHard0(1 To cap): ReDim gSmall(1 To cap): ReDim gKey(1 To cap): ReDim gRij(1 To cap): ReDim gKind(1 To cap): ReDim gPart(1 To cap)
    gN = 0
    For r = aE1 To aE2
        If HeeftBedragV(r) Then
            nb(r) = UCase$(TxtV(v(r, 5)))
            If nb(r) <> "N" And nb(r) <> "B" Then
                Rap "Bronrij " & r & ": bedrag " & v(r, 6) & " zonder N of B in kolom E (wordt niet meegenomen). Pas het bronblad aan."
            Else
                If Not (alleenOnverbonden And linkN(r) > 0) Then
                    ok(r) = True
                    di(r) = UCase$(TxtV(v(r, 4)))
                    sg(r) = IIf(nb(r) = "N", 1, -1) * v(r, 6)
                    gb(r) = rKop(r)
                    txt(r) = LCase$(VolTekst(r))
                    lijst(r) = (InStr(txt(r), "lijst") > 0 And InStr(txt(r), "overheveling of afram") > 0)
                End If
            End If
        End If
    Next r
    'regels zonder tekst onder dezelfde budgetkop als een lijst-regel horen daar ook bij
    For i = aE1 To aE2
        If ok(i) And Not lijst(i) And Trim$(txt(i)) = "" Then
            For j = aE1 To aE2
                If ok(j) And lijst(j) And gb(j) = gb(i) Then lijst(i) = True: Exit For
            Next j
        End If
    Next i
    'budgetneutrale paren: gelijk bedrag, tegengesteld teken, zelfde grootboek of onderling verwijzend
    For i = aE1 To aE2
        If ok(i) And gpair(i) = 0 And Not lijst(i) Then
            For j = i + 1 To aE2
                If ok(j) And gpair(j) = 0 And Not lijst(j) Then
                    If nb(i) <> nb(j) And Abs(v(i, 6) - v(j, 6)) < 0.005 Then
                        If alleenOnverbonden Or gb(i) = gb(j) Or Verwijst(txt(i), gb(j)) Or Verwijst(txt(j), gb(i)) Or (TxtV(v(i, 17)) <> "" And TxtV(v(i, 17)) = TxtV(v(j, 17))) Then
                            gpair(i) = j: gpair(j) = i: Exit For
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    'bundelkandidaten melden (gelijk bedrag, tegengesteld, zonder aantoonbaar verband)
    For i = aE1 To aE2
        If Not alleenOnverbonden And ok(i) And gpair(i) = 0 And Not lijst(i) And nb(i) = "N" Then
            For j = aE1 To aE2
                If ok(j) And gpair(j) = 0 And Not lijst(j) And nb(j) = "B" And j <> i Then
                    If Abs(v(i, 6) - v(j, 6)) < 0.005 Then Rap "Mogelijke bundel (eerst bevestigen): bronrij " & i & " (N) en " & j & " (B), beide " & Format$(v(i, 6), "#,##0") & ". Er is geen aantoonbaar verband, daarom niet gebundeld."
                End If
            Next j
        End If
    Next i
    'groepen
    For r = aE1 To aE2
        If ok(r) And groupOf(r) = 0 Then
            If gpair(r) > 0 Then
                If nb(r) = "N" Or (nb(r) = "B" And r < gpair(r) And False) Then
                    gN = gN + 1: groupOf(r) = gN: groupOf(gpair(r)) = gN
                    gBlok(gN) = IIf(di(r) = "S", 2, 1)
                    gVal(gN) = 0: gHard0(gN) = True: gKey(gN) = rKey(r): gRij(gN) = r
                    gKind(gN) = 1: gPart(gN) = gpair(r)
                Else
                    'het N-deel van het paar maakt de groep; dit B-deel wordt daar afgehandeld
                End If
            ElseIf lijst(r) Then
                'alle lijst-regels onder dezelfde budgetkop samen
                gN = gN + 1: gBlok(gN) = 3: gKey(gN) = rKey(r): gRij(gN) = r
                s = 0: tm = ""
                For j = r To aE2
                    If ok(j) And lijst(j) And gpair(j) = 0 And groupOf(j) = 0 And gb(j) = gb(r) Then
                        groupOf(j) = gN
                        s = s + sg(j)
                        tm = tm & IIf(tm = "", "", "+") & Term(j, "N")
                    End If
                Next j
                gVal(gN) = s: gTerm(gN) = tm
                gKind(gN) = 2
            Else
                gN = gN + 1: groupOf(r) = gN
                gBlok(gN) = IIf(di(r) = "S", 2, 1)
                gVal(gN) = sg(r): gTerm(gN) = Term(r, "N"): gKey(gN) = rKey(r): gRij(gN) = r
                gSmall(gN) = (Abs(sg(r)) < grens)
            End If
        End If
    Next r
    'B-delen van paren zonder N-partner in dezelfde pass (bijv. beide B): afvangen
    For r = aE1 To aE2
        If ok(r) And groupOf(r) = 0 And gpair(r) > 0 Then groupOf(r) = groupOf(gpair(r))
    Next r
End Sub

Private Function GroepOms(ByVal i As Long) As String
    Select Case gKind(i)
        Case 1: GroepOms = OmsPaar(gRij(i), gPart(i))
        Case 2: GroepOms = OmsLijst(gRij(i))
        Case Else: GroepOms = Oms(gRij(i))
    End Select
End Function

Private Function Verwijst(ByVal tekst As String, ByVal kop As String) As Boolean
    If Len(kop) = 6 Then Verwijst = (InStr(tekst, kop) > 0)
End Function

Private Function OmsPaar(ByVal i As Long, ByVal j As Long) As String
    Dim t As String, bedr As String
    t = Oms(i)
    If InStr(1, t, "budgetneutraal", vbTextCompare) = 0 And InStr(1, t, "saldoneutraal", vbTextCompare) = 0 Then
        bedr = Format$(v(i, 6), "#,##0")
        t = t & ", budgetneutraal, " & eur & " " & bedr
    End If
    OmsPaar = t
End Function

Private Function OmsLijst(ByVal r As Long) As String
    Dim t As String, nm As String, onderwerp As String, d As String
    t = Oms(r)
    If InStr(1, t, "naar " & (jaar + 1), vbTextCompare) > 0 Then
        OmsLijst = t
    Else
        nm = LCase$(rKopNm(r))
        If InStr(1, VolTekst(r), "afram", vbTextCompare) > 0 Then
            onderwerp = "Aframing budget " & nm
        Else
            onderwerp = "Overheveling budget " & nm
        End If
        OmsLijst = onderwerp & " naar " & (jaar + 1)
    End If
    If UCase$(TxtV(v(r, 4))) = "S" And InStr(OmsLijst, "(structureel)") = 0 Then OmsLijst = OmsLijst & " (structureel)"
End Function

'formule-term voor een bronrij; teken "N" = N positief (exploitatie, investeringen), "B" = B positief (reserves, voorzieningen)
Private Function Term(ByVal r As Long, ByVal plus As String) As String
    Term = "IF(TRIM(" & bq & "$E$" & r & ")=""" & plus & """,1,-1)*" & bq & "$F$" & r
End Function

'==============================================================================
' VOLLEDIG OPBOUWEN (stap 0 t/m 8)
'==============================================================================
Private Function BouwVolledig(wbk As Workbook) As Boolean
    Dim ws As Worksheet, i As Long, j As Long, b As Long, rw As Long
    Dim ord() As Long, tmp As Long, nIn As Long, kleinIdx As String
    Dim first As Long, last As Long, bundelTerm As String, bundelN As Long
    Dim rr As Long, hr As Long, p As Long

    OogstOmschrijvingen wbk
    BouwGroepen False
    If mRap.Count > 0 Then
        'meldingen over bronbedragen zonder N/B zijn waarschuwingen, geen blokkade
    End If

    Set ws = SheetByNameV(wbk, SH_PRES)
    If ws Is Nothing Then
        Set ws = wbk.Worksheets.Add(Before:=wbk.Worksheets(1))
        ws.Name = SH_PRES
    End If
    SchrijfHulpV wbk, True
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.EntireRow.Hidden = False
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.ClearOutline
    On Error GoTo 0
    ws.Cells.Font.Name = "Tahoma"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = vbBlack
    ws.Columns(1).ColumnWidth = 66
    ws.Columns(2).ColumnWidth = 16
    On Error Resume Next
    ws.Tab.Color = RGB(255, 165, 0)
    On Error GoTo 0

    ws.Cells(1, 1).Value = "Verschillenanalyses " & rapN & "e financi" & ee & "le rapportage " & jaar
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 14
    ws.Cells(2, 1).Value = "Toelichting op de mutaties in baten en lasten, investeringen, reserves en voorzieningen"
    rw = 4
    SectieKop ws, rw, "1. EXPLOITATIE"
    TabelKop ws, rw + 1, "Omschrijving mutatie"
    Legenda ws, rw + 2, " -/- = voordeel"
    rw = rw + 3

    Dim subRij(1 To 3) As Long, blokNaam(1 To 3) As String, subNaam(1 To 3) As String, subLbl(1 To 3) As String
    blokNaam(1) = "Incidentele mutaties": subNaam(1) = "incidenteel"
    blokNaam(2) = "Structurele mutaties": subNaam(2) = "structureel"
    blokNaam(3) = "Overhevelingen en aframingen naar " & (jaar + 1): subNaam(3) = "overhevelingen en aframingen"

    For b = 1 To 3
        ws.Cells(rw, 1).Value = blokNaam(b): ws.Cells(rw, 1).Font.Bold = True
        Legenda ws, rw, " -/- = voordeel"
        rw = rw + 1
        first = rw
        'groepen van dit blok, aflopend op bedrag; bundelregel onderaan
        ReDim ord(1 To gN + 1): nIn = 0: bundelTerm = "": bundelN = 0
        For i = 1 To gN
            If gBlok(i) = b Then
                If gSmall(i) And b < 3 And Not gHard0(i) Then
                    bundelN = bundelN + 1
                Else
                    nIn = nIn + 1: ord(nIn) = i
                End If
            End If
        Next i
        'enkele kleine post: individueel laten staan
        If bundelN = 1 Then
            For i = 1 To gN
                If gBlok(i) = b And gSmall(i) And Not gHard0(i) Then nIn = nIn + 1: ord(nIn) = i
            Next i
            bundelN = 0
        End If
        For i = 1 To nIn - 1
            For j = i + 1 To nIn
                If gVal(ord(j)) > gVal(ord(i)) Then tmp = ord(i): ord(i) = ord(j): ord(j) = tmp
            Next j
        Next i
        For i = 1 To nIn
            If gHard0(ord(i)) Then
                PutRij ws, rw, GroepOms(ord(i)), "", True
            Else
                PutRij ws, rw, GroepOms(ord(i)), "=" & gTerm(ord(i)), False
            End If
            rw = rw + 1
        Next i
        If bundelN > 1 Then
            For i = 1 To gN
                If gBlok(i) = b And gSmall(i) And Not gHard0(i) Then bundelTerm = bundelTerm & IIf(bundelTerm = "", "", "+") & gTerm(i)
            Next i
            PutRij ws, rw, "Diverse verschillen < " & eur & " " & Format$(grens, "#,##0"), "=" & bundelTerm, False
            rw = rw + 1
        End If
        last = rw - 1
        If last < first Then last = first: PutRij ws, rw, "(geen mutaties)", "", True: rw = rw + 1
        subRij(b) = rw
        SetF ws.Cells(rw, 2), "=SUM($B$" & first & ":$B$" & last & ")"
        SetF ws.Cells(rw, 1), "=IF($B$" & rw & ">0,""Subtotaal nadelig " & subNaam(b) & """,IF($B$" & rw & "<0,""Subtotaal voordelig " & subNaam(b) & """,""Subtotaal " & subNaam(b) & """))"
        TotaalOpmaak ws, rw, 1
        rw = rw + 2
    Next b
    pTot = rw - 1
    SetF ws.Cells(pTot, 2), "=$B$" & subRij(1) & "+$B$" & subRij(2) & "+$B$" & subRij(3)
    SetF ws.Cells(pTot, 1), "=IF($B$" & pTot & ">0,""TOTAAL NADELIG BEGROTINGSSALDO ""&'" & SH_HULP & "'!$B$5,IF($B$" & pTot & "<0,""TOTAAL VOORDELIG BEGROTINGSSALDO ""&'" & SH_HULP & "'!$B$5,""TOTAAL BEGROTINGSSALDO ""&'" & SH_HULP & "'!$B$5))"
    TotaalOpmaak ws, pTot, 1
    rw = pTot + 2

    If Not SchrijfInvResVz(ws, rw) Then BouwVolledig = False: Exit Function
    BouwVolledig = True
End Function

Private Function SchrijfInvResVz(ws As Worksheet, ByRef rw As Long) As Boolean
    Dim r As Long, first As Long, last As Long, n As Long, nP As Long
    Dim kop As String, tot As Long, rub() As String, nr As Long, i As Long, k As Long, key As String
    Dim a As String, fr As String

    '--- 2. INVESTERINGEN
    SectieKop ws, rw, "2. INVESTERINGEN"
    TabelKop ws, rw + 1, "Categorie"
    Legenda ws, rw + 2, " -/- = lagere investering"
    rw = rw + 3: first = rw
    n = 0: nP = 0
    For r = aI1 To aI2
        If HeeftBedragV(r) Then
            If UCase$(TxtV(v(r, 5))) = "N" Or UCase$(TxtV(v(r, 5))) = "B" Then
                n = n + 1
                If TxtV(v(r, 16)) <> "" Then nP = nP + 1
            End If
        End If
    Next r
    If n > 0 And nP > 0 And nP < n Then
        Rap "Investeringen: kolom P (rubriek) is voor " & (n - nP) & " van de " & n & " mutatieregels leeg. Vul kolom P in het bronblad in (ik verzin geen categorie" & ee & "n) en start opnieuw."
        Exit Function
    End If
    If n > 0 And nP = n Then
        ReDim rub(1 To n): nr = 0
        For r = aI1 To aI2
            If HeeftBedragV(r) Then
                If UCase$(TxtV(v(r, 5))) = "N" Or UCase$(TxtV(v(r, 5))) = "B" Then
                    key = LCase$(TxtV(v(r, 16))): k = 0
                    For i = 1 To nr
                        If LCase$(rub(i)) = key Then k = i
                    Next i
                    If k = 0 Then
                        nr = nr + 1: rub(nr) = TxtV(v(r, 16))
                        fr = "=SUMPRODUCT((LOWER(TRIM(" & Rg("P", aI1, aI2) & "))=""" & key & """)*((TRIM(" & Rg("E", aI1, aI2) & ")=""N"")-(TRIM(" & Rg("E", aI1, aI2) & ")=""B""))*" & Rg("F", aI1, aI2) & ")"
                        PutRij ws, rw, UCase$(Left$(rub(nr), 1)) & Mid$(rub(nr), 2), fr, False
                        rw = rw + 1
                    End If
                End If
            End If
        Next r
    Else
        For r = aI1 To aI2
            If HeeftBedragV(r) Then
                If UCase$(TxtV(v(r, 5))) = "N" Or UCase$(TxtV(v(r, 5))) = "B" Then
                    PutRij ws, rw, KopOms(r), "=" & Term(r, "N"), False
                    rw = rw + 1
                End If
            End If
        Next r
    End If
    last = rw - 1
    If last < first Then last = first
    ws.Cells(rw + 0, 1).Value = "": pInvTot = rw + 1
    SchrijfTotaal ws, pInvTot, first, last, "HOGERE INVESTERINGEN", "LAGERE INVESTERINGEN", "INVESTERINGEN"
    rw = pInvTot + 2

    '--- 3. RESERVES
    SectieKop ws, rw, "3. RESERVES"
    TabelKop ws, rw + 1, "Reserve"
    Legenda ws, rw + 2, " -/- = afname reserve"
    rw = rw + 3: first = rw
    For r = aR1 To aR2
        If HeeftBedragV(r) Then
            a = UCase$(TxtV(v(r, 5)))
            If a = "N" Or a = "B" Then
                If Left$(TxtV(v(r, 18)), 1) <> "R" Then Rap "Reserves bronrij " & r & ": kolom R (R-helper) is leeg of geen R-code; sectie D en de controle tellen deze regel dan niet mee."
                PutRij ws, rw, KopOms(r), "=" & Term(r, "B"), False
                rw = rw + 1
            End If
        End If
    Next r
    last = rw - 1
    If last < first Then last = first
    pResTot = rw + 1
    SchrijfTotaal ws, pResTot, first, last, "TOENAME RESERVES", "AFNAME RESERVES", "RESERVES"
    rw = pResTot + 2

    '--- 4. VOORZIENINGEN
    SectieKop ws, rw, "4. VOORZIENINGEN"
    TabelKop ws, rw + 1, "Voorziening"
    Legenda ws, rw + 2, " -/- = afname voorziening"
    rw = rw + 3: first = rw
    For r = aV1 To aV2
        If HeeftBedragV(r) Then
            a = UCase$(TxtV(v(r, 5)))
            If a = "N" Or a = "B" Then
                If Left$(TxtV(v(r, 18)), 1) <> "O" Then Rap "Voorzieningen bronrij " & r & ": kolom R (O-helper) is leeg of geen O-code; sectie D en de controle tellen deze regel dan niet mee."
                PutRij ws, rw, KopOms(r), "=" & Term(r, "B"), False
                rw = rw + 1
            End If
        End If
    Next r
    last = rw - 1
    If last < first Then last = first
    pVzTot = rw + 1
    SchrijfTotaal ws, pVzTot, first, last, "TOENAME VOORZIENINGEN", "AFNAME VOORZIENINGEN", "VOORZIENINGEN"
    rw = pVzTot + 2

    '--- controleblok
    SchrijfControle ws, rw
    SchrijfInvResVz = True
End Function

'omschrijving bij inv/res/vz: naam van de kop + toelichting
Private Function KopOms(ByVal r As Long) As String
    Dim t As String, harvT As String, q As String
    t = ""
    If Not harv Is Nothing And rKey(r) <> "" Then
        On Error Resume Next
        t = harv.Item(rKey(r))
        On Error GoTo 0
    End If
    If t <> "" Then KopOms = t: Exit Function
    q = TxtV(v(r, 17))
    If q = "" Then q = Kort(VolTekst(r))
    If rKopNm(r) <> "" And q <> "" Then
        KopOms = rKopNm(r) & ", " & LCase$(Left$(q, 1)) & Mid$(q, 2)
    ElseIf rKopNm(r) <> "" Then
        KopOms = rKopNm(r)
    Else
        KopOms = q
    End If
    Rap "Concept-omschrijving (nakijken): bronrij " & r & " > " & KopOms
End Function

Private Sub SchrijfTotaal(ws As Worksheet, ByVal rw As Long, ByVal first As Long, ByVal last As Long, ByVal tPos As String, ByVal tNeg As String, ByVal tNeut As String)
    SetF ws.Cells(rw, 2), "=SUM($B$" & first & ":$B$" & last & ")"
    SetF ws.Cells(rw, 1), "=IF($B$" & rw & ">0,""TOTAAL " & tPos & " ""&'" & SH_HULP & "'!$B$5,IF($B$" & rw & "<0,""TOTAAL " & tNeg & " ""&'" & SH_HULP & "'!$B$5,""TOTAAL " & tNeut & " ""&'" & SH_HULP & "'!$B$5))"
    If InStr(tPos, "TOENAME") > 0 Then
        TotaalOpmaak ws, rw, -1
    Else
        TotaalOpmaak ws, rw, 1
    End If
End Sub

Private Sub SchrijfControle(ws As Worksheet, ByVal rw As Long)
    Dim lbl As Variant, tot As Variant, hb As Variant, i As Long, f As String
    ws.Cells(rw, 1).Value = "Controle aansluiting op bronblad"
    ws.Cells(rw, 1).Font.Bold = True
    lbl = Array("Exploitatie", "Investeringen", "Reserves", "Voorzieningen", "Exploitatie detail (alle mutatieregels geteld)")
    tot = Array(pTot, pInvTot, pResTot, pVzTot, pTot)
    hb = Array(19, 20, 21, 22, 18)
    For i = 0 To 4
        ws.Cells(rw + 1 + i, 1).Value = lbl(i)
        f = "=IF(ROUND($B$" & tot(i) & "-'" & SH_HULP & "'!$B$" & hb(i) & ",2)=0,"""",""LET OP verschil ""&TEXT($B$" & tot(i) & "-'" & SH_HULP & "'!$B$" & hb(i) & ",""#.##0""))"
        SetF ws.Cells(rw + 1 + i, 2), f
        ws.Cells(rw + 1 + i, 2).Font.Color = RGB(192, 0, 0)
        ws.Cells(rw + 1 + i, 2).Font.Bold = True
    Next i
End Sub

'------------------------------------------------------------------------------
' Opmaak
'------------------------------------------------------------------------------
Private Sub PutRij(ws As Worksheet, ByVal r As Long, ByVal desc As String, ByVal f As String, ByVal hard0 As Boolean)
    If Left$(desc, 1) = "=" Then desc = "'" & desc
    ws.Cells(r, 1).Value = desc
    If hard0 Then
        ws.Cells(r, 2).Value = 0
    Else
        SetF ws.Cells(r, 2), f
    End If
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Name = "Tahoma"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Size = 10
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Color = vbBlack
    ws.Cells(r, 2).NumberFormat = "#,##0;-#,##0;-"
End Sub

Private Sub SectieKop(ws As Worksheet, ByVal r As Long, ByVal t As String)
    ws.Cells(r, 1).Value = t
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Interior.Color = RGB(7, 61, 136)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 13
    End With
End Sub

Private Sub TabelKop(ws As Worksheet, ByVal r As Long, ByVal t As String)
    ws.Cells(r, 1).Value = t
    ws.Cells(r, 2).Value = "Bedrag " & eur & " " & jaar
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
    ws.Cells(r, 2).HorizontalAlignment = xlRight
End Sub

Private Sub Legenda(ws As Worksheet, ByVal r As Long, ByVal t As String)
    ws.Cells(r, 2).Value = t
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Name = "Tahoma"
    ws.Cells(r, 2).Font.Size = 10
End Sub

'richting +1: > 0 rood, < 0 groen (exploitatie, investeringen); -1: > 0 groen, < 0 rood
Private Sub TotaalOpmaak(ws As Worksheet, ByVal r As Long, ByVal richting As Long)
    Dim rg As Range, cPos As Long, cNeg As Long
    Set rg = ws.Cells(r, 2)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Name = "Calibri"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Size = 11
    ws.Cells(r, 2).NumberFormat = "#,##0;-#,##0;-"
    If richting = 1 Then
        cPos = RGB(192, 0, 0): cNeg = RGB(0, 128, 0)
    Else
        cPos = RGB(0, 128, 0): cNeg = RGB(192, 0, 0)
    End If
    On Error Resume Next
    rg.FormatConditions.Delete
    With rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = cPos
    End With
    With rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = cNeg
    End With
    ws.Cells(r, 1).FormatConditions.Delete
    With ws.Cells(r, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=$B$" & r & ">0")
        .Font.Color = cPos
    End With
    With ws.Cells(r, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=$B$" & r & "<0")
        .Font.Color = cNeg
    End With
    On Error GoTo 0
End Sub

'==============================================================================
' BIJWERKEN (stap 8 en 9)
'==============================================================================
Private Function Bijwerken(wbk As Workbook) As Boolean
    Dim ws As Worksheet, r As Long, i As Long, b As Long
    Set ws = SheetByNameV(wbk, SH_PRES)
    If ws Is Nothing Then Rap "Het tabblad '" & SH_PRES & "' bestaat nog niet. Kies 'Verschillenanalyse maken'.": Exit Function
    If Not Structuur(ws) Then Rap "De opbouw van '" & SH_PRES & "' is niet herkend (koppen 1. EXPLOITATIE t/m 4. VOORZIENINGEN, blokken en totaalregels). Kies 'Verschillenanalyse maken'.": Exit Function
    SchrijfHulpV wbk, False
    OogstOmschrijvingen wbk

    '--- stap 1: koppelingen inlezen, verdwenen bronregels opruimen
    ReDim linkN(1 To lastR)
    Dim nDood As Long
    nDood = RuimOp(ws)
    If Not Structuur(ws) Then Rap "Opbouw niet meer herkend na opruimen.": Exit Function

    '--- stap 2: koppelingen opnieuw tellen
    TelKoppelingen ws
    For r = 1 To lastR
        If linkN(r) > 1 Then Rap "Bronrij " & r & " komt " & linkN(r) & " keer voor in de analyse (dubbeltelling). Corrigeer handmatig."
    Next r

    '--- stap 3: nieuwe bronregels toevoegen
    VoegToe ws
    If Not Structuur(ws) Then Rap "Opbouw niet meer herkend na toevoegen.": Exit Function
    HerstelTotalen ws
    Bijwerken = True
End Function

'zoekt kop- en totaalregels op tekst
Private Function Structuur(ws As Worksheet) As Boolean
    Dim r As Long, lr As Long, a As String, f As String, b As Long
    lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For b = 1 To 3: pHdr(b) = 0: pSub(b) = 0: Next b
    pTot = 0: pInv = 0: pRes = 0: pVz = 0: pCtl = 0: pInvTot = 0: pResTot = 0: pVzTot = 0
    For r = 1 To lr
        a = Trim$(CStr(ws.Cells(r, 1).Value))
        f = ""
        If ws.Cells(r, 1).HasFormula Then f = UCase$(ws.Cells(r, 1).Formula)
        If a = "Incidentele mutaties" Then pHdr(1) = r
        If a = "Structurele mutaties" Then pHdr(2) = r
        If LCase$(a) Like "overhevelingen en aframingen naar*" And f = "" Then pHdr(3) = r
        If a = "2. INVESTERINGEN" Then pInv = r
        If a = "3. RESERVES" Then pRes = r
        If a = "4. VOORZIENINGEN" Then pVz = r
        If a = "Controle aansluiting op bronblad" Then pCtl = r
        If f <> "" Then
            If InStr(f, "SUBTOTAAL INCIDENTEEL") > 0 Then pSub(1) = r
            If InStr(f, "SUBTOTAAL STRUCTUREEL") > 0 Then pSub(2) = r
            If InStr(f, "SUBTOTAAL OVERHEVELINGEN") > 0 Then pSub(3) = r
            If InStr(f, "BEGROTINGSSALDO") > 0 And InStr(f, "TOTAAL") > 0 Then pTot = r
            If InStr(f, "TOTAAL") > 0 And InStr(f, "INVESTERINGEN") > 0 Then pInvTot = r
            If InStr(f, "TOTAAL") > 0 And InStr(f, "RESERVES") > 0 Then pResTot = r
            If InStr(f, "TOTAAL") > 0 And InStr(f, "VOORZIENINGEN") > 0 Then pVzTot = r
        End If
    Next r
    Structuur = (pHdr(1) > 0 And pHdr(2) > 0 And pHdr(3) > 0 And pSub(1) > 0 And pSub(2) > 0 And pSub(3) > 0 And pTot > 0 And pInv > 0 And pRes > 0 And pVz > 0 And pInvTot > 0 And pResTot > 0 And pVzTot > 0)
End Function

'datarijen van een onderdeel op het presentatieblad
Private Sub DataBereik(ByVal soort As Long, ByRef r1 As Long, ByRef r2 As Long)
    Select Case soort
        Case 1, 2, 3: r1 = pHdr(soort) + 1: r2 = pSub(soort) - 1
        Case 4: r1 = pInv + 3: r2 = pInvTot - 1
        Case 5: r1 = pRes + 3: r2 = pResTot - 1
        Case 6: r1 = pVz + 3: r2 = pVzTot - 1
    End Select
    Do While r2 >= r1
        If Trim$(CStr(Application.WorksheetFunction.CountA(0))) = "" Then Exit Do
        Exit Do
    Loop
End Sub

Private Function LaatsteData(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long
    LaatsteData = r1 - 1
    For r = r1 To r2
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" And Not ws.Cells(r, 1).HasFormula Then LaatsteData = r
    Next r
End Function

'verdwenen bronregels: verwijder termen en lege rijen, van onder naar boven
Private Function RuimOp(ws As Worksheet) As Long
    Dim soort As Long, r1 As Long, r2 As Long, r As Long, n As Long
    Dim f As String, parts() As String, keep As String, i As Long, ref As Long, verw As Boolean, tot As Long
    Dim del() As Long, nDel As Long
    ReDim del(1 To 2000)
    For soort = 1 To 6
        DataBereik soort, r1, r2
        r2 = LaatsteData(ws, r1, r2)
        For r = r1 To r2
            f = ws.Cells(r, 2).Formula
            If Left$(f, 4) = "=IF(" And InStr(f, "-1)*") > 0 Then
                parts = Split(Mid$(f, 2), "+IF(")
                keep = ""
                verw = False
                For i = 0 To UBound(parts)
                    If i > 0 Then parts(i) = "IF(" & parts(i)
                    ref = EersteRef(parts(i))
                    If IsDood(ref, soort) Then
                        verw = True
                        Rap "Verdwenen bronregel opgeruimd: bronrij " & ref & " (presentatierij " & r & ": " & Left$(CStr(ws.Cells(r, 1).Value), 50) & ")"
                    Else
                        keep = keep & IIf(keep = "", "", "+") & parts(i)
                    End If
                Next i
                If verw Then
                    If keep = "" Then
                        nDel = nDel + 1: del(nDel) = r
                    Else
                        SetF ws.Cells(r, 2), "=" & keep
                    End If
                End If
            End If
        Next r
    Next soort
    'verwijderen van onder naar boven
    For i = nDel To 1 Step -1
        ws.Rows(del(i)).Delete
    Next i
    RuimOp = nDel
End Function

Private Function IsDood(ByVal ref As Long, ByVal soort As Long) As Boolean
    Dim nb As String
    If ref < 1 Or ref > lastR Then IsDood = True: Exit Function
    nb = UCase$(TxtV(v(ref, 5)))
    If Not IsNumV(v(ref, 6)) Then IsDood = True: Exit Function
    If v(ref, 6) = 0 Then IsDood = True: Exit Function
    If nb <> "N" And nb <> "B" Then IsDood = True
End Function

Private Sub TelKoppelingen(ws As Worksheet)
    Dim soort As Long, r1 As Long, r2 As Long, r As Long, f As String, k As Long, p As Long, n As Long, c As String
    ReDim linkN(1 To lastR)
    For soort = 1 To 6
        DataBereik soort, r1, r2
        r2 = LaatsteData(ws, r1, r2)
        For r = r1 To r2
            f = ws.Cells(r, 2).Formula
            If UCase$(Left$(f, 12)) = "=SUMPRODUCT(" Then
                For k = aI1 To aI2: linkN(k) = linkN(k) + 1: Next k
            Else
                p = 1
                Do
                    p = InStr(p, f, "$F$")
                    If p = 0 Then Exit Do
                    n = 0: c = Mid$(f, p + 3, 1): k = p + 3
                    Do While c Like "#"
                        n = n * 10 + CLng(c): k = k + 1: c = Mid$(f, k, 1)
                    Loop
                    If n >= 1 And n <= lastR Then linkN(n) = linkN(n) + 1
                    p = p + 3
                Loop
            End If
        Next r
    Next soort
End Sub

'nieuwe bronregels in de juiste blokken zetten
Private Sub VoegToe(ws As Worksheet)
    Dim i As Long, b As Long, r As Long, nieuw As Long
    Dim rp As Long, f As String, a As String, bund As Long, r1 As Long, r2 As Long
    Dim paren As Long
    'exploitatie
    BouwGroepen True
    For i = 1 To gN
        If gHard0(i) Then
            paren = paren + 1
        Else
            Structuur ws
            DataBereik gBlok(i), r1, r2
            If gSmall(i) And gBlok(i) < 3 Then
                bund = ZoekBundel(ws, r1, r2)
                If bund > 0 Then
                    SetF ws.Cells(bund, 2), ws.Cells(bund, 2).Formula & "+" & gTerm(i)
                    Rap "Nieuwe bronrij " & gRij(i) & " toegevoegd aan 'Diverse verschillen' (bedrag " & Format$(gVal(i), "#,##0") & ")."
                Else
                    NieuweRij ws, r2 + 1, "Diverse verschillen < " & eur & " " & Format$(grens, "#,##0"), "=" & gTerm(i)
                    Rap "Nieuwe bronrij " & gRij(i) & " in nieuwe regel 'Diverse verschillen'."
                End If
            Else
                NieuweRij ws, r2 + 1, GroepOms(i), "=" & gTerm(i)
                Rap "Nieuwe bronrij " & gRij(i) & " toegevoegd: " & GroepOms(i) & " (" & Format$(gVal(i), "#,##0") & ")"
            End If
        End If
    Next i
    If paren > 0 Then Rap paren & " onverbonden budgetneutrale paren (gelijk bedrag, N en B) zijn in orde en krijgen geen eigen regel."
    'investeringen, reserves, voorzieningen
    NieuwSectie ws, 4, aI1, aI2, "N"
    NieuwSectie ws, 5, aR1, aR2, "B"
    NieuwSectie ws, 6, aV1, aV2, "B"
End Sub

Private Function ZoekBundel(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Long
    Dim r As Long
    For r = r1 To r2
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) Like "diverse verschillen*" Then ZoekBundel = r
    Next r
End Function

Private Sub NieuweRij(ws As Worksheet, ByVal pos As Long, ByVal desc As String, ByVal f As String)
    ws.Rows(pos).Insert
    On Error Resume Next
    ws.Rows(pos - 1).Copy
    ws.Rows(pos).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error GoTo 0
    ws.Cells(pos, 1).Value = desc
    SetF ws.Cells(pos, 2), f
    Structuur ws
End Sub

Private Sub NieuwSectie(ws As Worksheet, ByVal soort As Long, ByVal r1b As Long, ByVal r2b As Long, ByVal plus As String)
    Dim r As Long, nb As String, a As Long, b As Long, d1 As Long, d2 As Long
    For r = r1b To r2b
        If HeeftBedragV(r) Then
            nb = UCase$(TxtV(v(r, 5)))
            If (nb = "N" Or nb = "B") And linkN(r) = 0 Then
                Structuur ws
                DataBereik soort, d1, d2
                If soort = 4 Then
                    NieuweRij ws, LaatsteData(ws, d1, d2) + 1, KopOms(r), "=" & Term(r, "N")
                Else
                    NieuweRij ws, LaatsteData(ws, d1, d2) + 1, KopOms(r), "=" & Term(r, "B")
                End If
                Rap "Nieuwe bronrij " & r & " toegevoegd aan " & IIf(soort = 4, "investeringen", IIf(soort = 5, "reserves", "voorzieningen")) & ": " & Left$(CStr(ws.Cells(LaatsteData(ws, d1, d2 + 1), 1).Value), 60)
            End If
        End If
    Next r
End Sub

'sorteren (blokken 1 t/m 3), bereiken van subtotalen/totalen herstellen
Private Sub HerstelTotalen(ws As Worksheet)
    Dim b As Long, r1 As Long, r2 As Long, soort As Long, lr As Long
    Structuur ws
    For b = 1 To 3
        DataBereik b, r1, r2
        SorteerBlok ws, r1, LaatsteData(ws, r1, r2)
        Structuur ws
    Next b
    For b = 1 To 3
        DataBereik b, r1, r2
        r2 = LaatsteData(ws, r1, r2)
        If r2 >= r1 Then SetF ws.Cells(pSub(b), 2), "=SUM($B$" & r1 & ":$B$" & r2 & ")"
    Next b
    SetF ws.Cells(pTot, 2), "=$B$" & pSub(1) & "+$B$" & pSub(2) & "+$B$" & pSub(3)
    For soort = 4 To 6
        DataBereik soort, r1, r2
        r2 = LaatsteData(ws, r1, r2)
        If r2 >= r1 Then
            Select Case soort
                Case 4: SetF ws.Cells(pInvTot, 2), "=SUM($B$" & r1 & ":$B$" & r2 & ")"
                Case 5: SetF ws.Cells(pResTot, 2), "=SUM($B$" & r1 & ":$B$" & r2 & ")"
                Case 6: SetF ws.Cells(pVzTot, 2), "=SUM($B$" & r1 & ":$B$" & r2 & ")"
            End Select
        End If
    Next soort
    'controleregels herschrijven
    If pCtl > 0 Then
        Dim tot As Variant, hb As Variant, i As Long, f As String, rr As Long
        tot = Array(pTot, pInvTot, pResTot, pVzTot, pTot)
        hb = Array(19, 20, 21, 22, 18)
        For i = 0 To 4
            rr = pCtl + 1 + i
            f = "=IF(ROUND($B$" & tot(i) & "-'" & SH_HULP & "'!$B$" & hb(i) & ",2)=0,"""",""LET OP verschil ""&TEXT($B$" & tot(i) & "-'" & SH_HULP & "'!$B$" & hb(i) & ",""#.##0""))"
            SetF ws.Cells(rr, 2), f
        Next i
    End If
End Sub

Private Sub SorteerBlok(ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Dim n As Long, i As Long, j As Long, bund As Long
    Dim ta() As String, tf() As String, tv() As Double, ts() As String, hx() As Boolean
    Dim sa As String, sf As String, sv As Double, hz As Boolean, k As Long, rr As Long, bundTxt As String, bundF As String, bundV As Double, bundH As Boolean
    If r2 <= r1 Then Exit Sub
    n = r2 - r1 + 1
    ReDim ta(1 To n): ReDim tf(1 To n): ReDim tv(1 To n): ReDim hx(1 To n)
    k = 0
    For rr = r1 To r2
        If LCase$(Trim$(CStr(ws.Cells(rr, 1).Value))) Like "diverse verschillen*" Then
            bundTxt = CStr(ws.Cells(rr, 1).Value): bundF = ws.Cells(rr, 2).Formula: bundH = Not ws.Cells(rr, 2).HasFormula: bund = 1
        Else
            k = k + 1
            ta(k) = CStr(ws.Cells(rr, 1).Value)
            tf(k) = ws.Cells(rr, 2).Formula
            hx(k) = Not ws.Cells(rr, 2).HasFormula
            If IsNumeric(ws.Cells(rr, 2).Value) Then tv(k) = CDbl(ws.Cells(rr, 2).Value)
        End If
    Next rr
    For i = 1 To k - 1
        For j = i + 1 To k
            If tv(j) > tv(i) Then
                sa = ta(i): ta(i) = ta(j): ta(j) = sa
                sf = tf(i): tf(i) = tf(j): tf(j) = sf
                sv = tv(i): tv(i) = tv(j): tv(j) = sv
                hz = hx(i): hx(i) = hx(j): hx(j) = hz
            End If
        Next j
    Next i
    For i = 1 To k
        ws.Cells(r1 + i - 1, 1).Value = ta(i)
        If hx(i) Then
            ws.Cells(r1 + i - 1, 2).Value = CDbl(tf(i))
        Else
            SetF ws.Cells(r1 + i - 1, 2), tf(i)
        End If
    Next i
    If bund = 1 Then
        ws.Cells(r1 + k, 1).Value = bundTxt
        SetF ws.Cells(r1 + k, 2), bundF
    End If
End Sub

'==============================================================================
' Ontwikkeling begrotingssaldo, sectie D (stap 10)
'==============================================================================
Private Sub SectieD(wbk As Workbook)
    Dim ws As Worksheet, r As Long, lr As Long, a As String, rKopR As Long, rKopV As Long, rD As Long
    Set ws = SheetByNameV(wbk, SH_ONTW)
    If ws Is Nothing Then Rap "Tabblad '" & SH_ONTW & "' niet gevonden: sectie D niet bijgewerkt.": Exit Sub
    lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = 1 To lr
        a = Trim$(CStr(ws.Cells(r, 1).Value))
        If a Like "D. Ontwikkeling reserves en voorzieningen*" Then rD = r
        If a = "Mutaties reserves" Then rKopR = r
        If a = "Mutaties voorzieningen" Then rKopV = r
    Next r
    If rD = 0 Or rKopR = 0 Or rKopV = 0 Then
        Rap "Sectie D (D. Ontwikkeling reserves en voorzieningen, Mutaties reserves, Mutaties voorzieningen) niet herkend: niet bijgewerkt."
        Exit Sub
    End If
    ws.Cells(rD, 1).Value = "D. Ontwikkeling reserves en voorzieningen " & jaar
    'reserves: +2 toevoeging, +3 onttrekking, +4 saldo
    ws.Cells(rKopR + 2, 1).Value = "Saldo toevoeging aan reserves"
    SetF ws.Cells(rKopR + 2, 4), "=SUMIFS(" & Rg("F", aR1, aR2) & "," & Rg("E", aR1, aR2) & ",""B""," & Rg("R", aR1, aR2) & ",""R*"")"
    ws.Cells(rKopR + 3, 1).Value = "Saldo onttrekking aan reserves"
    SetF ws.Cells(rKopR + 3, 4), "=SUMIFS(" & Rg("F", aR1, aR2) & "," & Rg("E", aR1, aR2) & ",""N""," & Rg("R", aR1, aR2) & ",""R*"")"
    SetF ws.Cells(rKopR + 4, 4), "=D" & (rKopR + 2) & "-D" & (rKopR + 3)
    SetF ws.Cells(rKopR + 4, 1), "=IF(D" & (rKopR + 4) & ">0,""Saldo toevoeging aan reserves"",IF(D" & (rKopR + 4) & "<0,""Saldo onttrekking aan reserves"",""Saldo mutatie reserves""))"
    'voorzieningen: +2 onttrekking, +3 dotatie, +4 saldo
    ws.Cells(rKopV + 2, 1).Value = "Saldo onttrekking voorzieningen"
    SetF ws.Cells(rKopV + 2, 4), "=SUMIFS(" & Rg("F", aV1, aV2) & "," & Rg("E", aV1, aV2) & ",""N""," & Rg("R", aV1, aV2) & ",""O*"")"
    ws.Cells(rKopV + 3, 1).Value = "Saldo dotatie voorzieningen"
    SetF ws.Cells(rKopV + 3, 4), "=SUMIFS(" & Rg("F", aV1, aV2) & "," & Rg("E", aV1, aV2) & ",""B""," & Rg("R", aV1, aV2) & ",""O*"")"
    SetF ws.Cells(rKopV + 4, 4), "=D" & (rKopV + 3) & "-D" & (rKopV + 2)
    SetF ws.Cells(rKopV + 4, 1), "=IF(D" & (rKopV + 4) & ">0,""Saldo dotatie voorzieningen"",IF(D" & (rKopV + 4) & "<0,""Saldo onttrekking voorzieningen"",""Saldo mutatie voorzieningen""))"
    SaldoKleur ws.Cells(rKopR + 4, 4)
    SaldoKleur ws.Cells(rKopV + 4, 4)
End Sub

Private Sub SaldoKleur(c As Range)
    c.Font.Color = vbBlack
    On Error Resume Next
    c.FormatConditions.Delete
    With c.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 176, 80)
    End With
    With c.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(255, 0, 0)
    End With
    On Error GoTo 0
End Sub

'==============================================================================
' Afsluitende controle en terugkoppeling
'==============================================================================
Private Sub Afsluiting(wbk As Workbook, ByVal volledig As Boolean)
    Dim ws As Worksheet, i As Long, t As String, nLet As Long, rr As Long, wsO As Worksheet
    Dim sd As Double, tot As Double, sg As Long
    Set ws = SheetByNameV(wbk, SH_PRES)
    Structuur ws
    t = IIf(volledig, "Verschillenanalyse opnieuw opgebouwd", "Verschillenanalyse bijgewerkt") & " uit '" & brn.Name & "'." & vbCrLf
    If pCtl > 0 Then
        For i = 1 To 5
            If CStr(ws.Cells(pCtl + i, 2).Value) <> "" Then
                nLet = nLet + 1
                t = t & "- " & CStr(ws.Cells(pCtl + i, 1).Value) & ": " & CStr(ws.Cells(pCtl + i, 2).Value) & vbCrLf
            End If
        Next i
    End If
    If nLet = 0 Then t = t & "Controle aansluiting op bronblad: alle vijf regels leeg (sluit aan)." & vbCrLf Else t = t & "LET OP: de analyse sluit niet aan op het bronblad. Zie Stap 8 van de skill (niet-gekoppelde regels)." & vbCrLf
    t = t & "Totalen: begrotingssaldo " & Format$(ws.Cells(pTot, 2).Value, "#,##0") & ", investeringen " & Format$(ws.Cells(pInvTot, 2).Value, "#,##0") & ", reserves " & Format$(ws.Cells(pResTot, 2).Value, "#,##0") & ", voorzieningen " & Format$(ws.Cells(pVzTot, 2).Value, "#,##0") & vbCrLf
    'verborgen rijen of groepering opsporen
    For rr = 1 To pCtl + 6
        If ws.Rows(rr).Hidden Then t = t & "LET OP: rij " & rr & " is verborgen." & vbCrLf
    Next rr
    If mRap.Count > 0 Then
        t = t & vbCrLf & "Meldingen:" & vbCrLf & RapTekst("")
    End If
    gResultaatV = t
    DashLog IIf(volledig, "Analyse opnieuw opgebouwd", "Analyse bijgewerkt") & " uit " & brn.Name & IIf(nLet = 0, ": alle controles leeg", ": LET OP, zie controleblok")
    If Not mStil Then
        ws.Activate
        ws.Range("A1").Select
        MsgBox t, IIf(nLet = 0, vbInformation, vbExclamation), "Verschillenanalyses firap"
    End If
End Sub

'==============================================================================
' Hulpfuncties
'==============================================================================
'formule schrijven; cel met tekstopmaak eerst op Standaard zetten (anders wordt de formule tekst)
Private Sub SetF(c As Range, ByVal f As String)
    If c.NumberFormat = "@" Then c.NumberFormat = "General"
    c.Formula = f
End Sub

Private Sub Rap(ByVal t As String)
    Dim it As Variant
    For Each it In mRap
        If CStr(it) = t Then Exit Sub
    Next it
    mRap.Add t
End Sub

Private Function RapTekst(ByVal kop As String) As String
    Dim t As String, it As Variant, i As Long
    If kop <> "" Then t = kop & vbCrLf
    For Each it In mRap
        i = i + 1
        If i <= 30 Then t = t & "- " & CStr(it) & vbCrLf
    Next it
    If mRap.Count > 30 Then t = t & "... en " & (mRap.Count - 30) & " meldingen meer." & vbCrLf
    RapTekst = t
End Function

Private Function EersteRef(ByVal f As String) As Long
    Dim p As Long, n As Long, k As Long, c As String
    p = InStr(f, "$F$")
    If p = 0 Then Exit Function
    k = p + 3: c = Mid$(f, k, 1)
    Do While c Like "#"
        n = n * 10 + CLng(c): k = k + 1: c = Mid$(f, k, 1)
    Loop
    EersteRef = n
End Function

Private Function SheetByNameV(wbk As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wbk.Worksheets
        If LCase$(ws.Name) = LCase$(nm) Then
            Set SheetByNameV = ws
            Exit Function
        End If
    Next ws
End Function

Private Function TxtV(ByVal x As Variant) As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    If IsNull(x) Then Exit Function
    TxtV = Trim$(Replace(CStr(x), Chr$(160), " "))
End Function

Private Function IsNumV(ByVal x As Variant) As Boolean
    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            IsNumV = True
    End Select
End Function

Private Function AllDigitsV(ByVal s As String) As Boolean
    AllDigitsV = (Len(s) > 0)
    If AllDigitsV Then AllDigitsV = Not (s Like "*[!0-9]*")
End Function

Private Function NormV(ByVal s As String) As String
    Dim i As Long, ch As String, o As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            o = o & ch
            If Len(o) >= 50 Then Exit For
        End If
    Next i
    NormV = o
End Function
