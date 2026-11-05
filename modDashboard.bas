Attribute VB_Name = "modDashboard"
Option Explicit
'==============================================================================
' Dashboard "Dashboard"
'   C4 = rapportage (1e firap, 2e firap of 3e firap)   C5 = begrotingsjaar
'   C6 = begrotingswijzigingsnummer (bijv. W16/26)
'
' Knop "Overgang naar nieuwe firap" = OvergangNaarNieuweFirap:
'   zet de werkmap om naar de gekozen rapportage en het gekozen jaar:
'   - hernoemt het bronblad (bijv. 3e firap 2026 wordt 1e firap 2027);
'     alle formules die naar het bronblad wijzen worden door Excel meegenomen
'   - zet de jaartallen in rij 10 van het bronblad (eerste jaar + 4 vervolgjaren)
'   - zet het wijzigingsnummer in het bronblad op de plek van het huidige nummer
'   - past alleen de kopteksten aan op de tabbladen 'Ontwikkeling begrotingssaldo',
'     'verschillenanalyses firap' en 'hulp verschillenanalyse'
'   Mutatiegegevens, omschrijvingen van mutaties en bedragen worden NIET gewijzigd
'   of gewist; een lijst met punten voor handmatige controle volgt na afloop.
' Alleen ASCII-tekens (trema via ChrW).
'==============================================================================

Private Const SH_DASH As String = "Dashboard"
Public gResultaatD As String
Private mWis As Boolean
Private mWisVast As Boolean
Private mStilD As Boolean
Private mMeldD As Collection
Private mAantal As Long
Private oRap As Long, oJr As Long, nRap As Long, nJr As Long
Private oPRap As Long, oPJr As Long, nPRap As Long, nPJr As Long
Private ee As String

Public Sub OvergangNaarNieuweFirap()
    mStilD = False: mWisVast = False
    DoOvergang
End Sub

Public Sub OvergangStil()
    mStilD = True: mWisVast = True: mWis = True
    DoOvergang
End Sub

Public Sub OvergangStilBehoud()
    mStilD = True: mWisVast = True: mWis = False
    DoOvergang
End Sub

Private Sub DoOvergang()
    Dim wbk As Workbook, d As Worksheet, src As Worksheet, ws As Worksheet
    Dim rap As String, jr As String, wz As String, target As String, k As Long
    Dim lr As Long, r As Long, c As Long, t As String, nieuw As String
    Dim antw As VbMsgBoxResult, nm As String, i As Long, sBeg As String, sOnv As String
    On Error GoTo Fout
    Set wbk = ActiveWorkbook
    Set mMeldD = New Collection
    ee = ChrW(235): mAantal = 0: gResultaatD = ""
    Set d = SheetBy(wbk, SH_DASH)
    If d Is Nothing Then Stopt "Het dashboard (tabblad '" & SH_DASH & "') is niet gevonden.": Exit Sub
    rap = Trim$(CStr(d.Range("C4").Value)): jr = Trim$(CStr(d.Range("C5").Value)): wz = Trim$(CStr(d.Range("C6").Value))
    If Not (LCase$(rap) Like "[123]e firap") Then Stopt "Kies bij Rapportage (C4) 1e firap, 2e firap of 3e firap.": Exit Sub
    If Not (IsNumeric(jr) And Len(jr) = 4) Then Stopt "Vul bij Begrotingsjaar (C5) een jaartal van vier cijfers in.": Exit Sub
    If Not IsWijz(wz) Then Stopt "Vul bij Begrotingswijzigingsnummer (C6) een nummer in zoals W16/26.": Exit Sub
    nRap = CLng(Left$(rap, 1)): nJr = CLng(jr)
    sBeg = Trim$(CStr(d.Range("C8").Value)): sOnv = Trim$(CStr(d.Range("C9").Value))
    If nRap = 1 Then
        If Not IsNumeric(sBeg) Or sBeg = "" Then Stopt "Bij een 1e firap moet het begrotingssaldo (C8) worden ingevuld: dit komt in Ontwikkeling begrotingssaldo D6.": Exit Sub
        If Not IsNumeric(sOnv) Or sOnv = "" Then Stopt "Bij een 1e firap moet het saldo onvoorzien (C9) worden ingevuld: dit komt in Ontwikkeling begrotingssaldo D20.": Exit Sub
    ElseIf (sBeg <> "" And Not IsNumeric(sBeg)) Or (sOnv <> "" And Not IsNumeric(sOnv)) Then
        Stopt "Begrotingssaldo (C8) en saldo onvoorzien (C9) moeten getallen zijn.": Exit Sub
    End If
    target = Left$(rap, 1) & "e firap " & jr

    'huidig bronblad zoeken
    For Each ws In wbk.Worksheets
        If LCase$(ws.Name) Like "[123]e*firap*" And LCase$(ws.Name) <> "verschillenanalyses firap" Then
            k = k + 1: Set src = ws
        End If
    Next ws
    If k = 0 Then Stopt "Geen bronblad gevonden (een tabblad zoals '3e firap 2026').": Exit Sub
    If k > 1 Then Stopt "Er zijn meerdere bronbladen; verwijder of hernoem de overbodige en start opnieuw.": Exit Sub
    nm = src.Name
    oRap = CLng(Left$(nm, 1))
    If Not (Right$(nm, 4) Like "20##") Then Stopt "Het jaartal ontbreekt aan het einde van de naam van het bronblad '" & nm & "'.": Exit Sub
    oJr = CLng(Right$(nm, 4))
    If LCase$(nm) <> LCase$(target) Then
        If Not SheetBy(wbk, target) Is Nothing Then Stopt "Er bestaat al een tabblad '" & target & "'.": Exit Sub
    End If
    'voorgaande rapportage (oud en nieuw)
    If oRap = 1 Then oPRap = 3: oPJr = oJr - 1 Else oPRap = oRap - 1: oPJr = oJr
    If nRap = 1 Then nPRap = 3: nPJr = nJr - 1 Else nPRap = nRap - 1: nPJr = nJr

    If Not mStilD Then
        antw = MsgBox("Overgang van '" & nm & "' naar '" & target & "'." & vbCrLf & vbCrLf & _
            "- bronblad wordt hernoemd" & vbCrLf & "- jaartallen in rij 10 worden " & nJr & " t/m " & (nJr + 4) & vbCrLf & _
            "- wijzigingsnummer wordt " & wz & vbCrLf & "- kopteksten op de analysebladen worden aangepast" & vbCrLf & _
            "- mutatiegegevens worden NIET gewist" & vbCrLf & vbCrLf & _
            IIf(sBeg <> "" Or sOnv <> "", "- begrotingssaldo en saldo onvoorzien komen in Ontwikkeling begrotingssaldo (D6 en D20)" & vbCrLf, "") & vbCrLf & _
            "Tip: sla eerst een kopie van het bestand op. Doorgaan?", vbYesNo + vbQuestion, "Overgang naar nieuwe firap")
        If antw <> vbYes Then Exit Sub
        antw = MsgBox("Ook alle mutaties van de vorige rapportage wissen?" & vbCrLf & vbCrLf & _
            "Ja = bedragen, N/B, categorie en toelichting van de mutaties in het bronblad worden gewist; de structuur, kopregels en formules blijven staan. De analyse (verschillenanalyses firap) wordt leeg opgebouwd en de inhoud van tbv adm pakket wordt gewist." & vbCrLf & _
            "Nee = mutaties laten staan.", vbYesNo + vbQuestion, "Mutaties wissen")
        mWis = (antw = vbYes)
    End If
    Application.ScreenUpdating = False

    '--- bronblad -------------------------------------------------------------
    src.Name = target
    'jaartallen rij 10
    c = 5: k = 0
    Do While c + 1 <= 19
        If UCase$(Trim$(CStr(src.Cells(10, c).Value))) = "N/B" And IsNumeric(src.Cells(10, c + 1).Value) Then
            If k = 0 Then
                src.Cells(10, c + 1).Value = nJr
            ElseIf Not src.Cells(10, c + 1).HasFormula Then
                src.Cells(10, c + 1).Value = nJr + k
            End If
            k = k + 1: c = c + 2
        Else
            Exit Do
        End If
    Loop
    Meld "Jaartallen in rij 10 van '" & target & "': " & k & " kolommen gezet (" & nJr & " t/m " & (nJr + k - 1) & ")."
    'wijzigingsnummer
    Dim gevonden As Boolean, oudWz As String
    For r = 1 To 9
        For c = 1 To 19
            t = Trim$(CStr(src.Cells(r, c).Value))
            If IsWijz(t) And Not gevonden Then
                oudWz = t: src.Cells(r, c).Value = wz: gevonden = True
                Meld "Wijzigingsnummer in bronblad " & src.Cells(r, c).Address(False, False) & ": " & oudWz & " wordt " & wz & "."
            End If
        Next c
    Next r
    If Not gevonden Then
        src.Cells(1, 6).Value = wz
        Meld "Geen bestaand wijzigingsnummer gevonden; " & wz & " gezet in F1 van het bronblad."
    End If
    'kopteksten bronblad
    t = CStr(src.Cells(1, 1).Value)
    If InStr(1, t, "RAPPORTAGE BEGROTING", vbTextCompare) > 0 Then
        nieuw = nRap & "e FINANCIELE RAPPORTAGE BEGROTING " & nJr & " "
        If nieuw <> t Then src.Cells(1, 1).Value = nieuw: mAantal = mAantal + 1
    End If
    lr = src.UsedRange.Row + src.UsedRange.Rows.Count - 1
    For r = 2 To lr
        t = CStr(src.Cells(r, 1).Value)
        If Not src.Cells(r, 1).HasFormula And t <> "" Then
            If InStr(1, t, "na vaststelling", vbTextCompare) > 0 Then Zet src.Cells(r, 1)
        End If
    Next r
    'punten voor handmatige controle in het bronblad
    For r = 1 To 9
        For c = 1 To 19
            t = CStr(src.Cells(r, c).Value)
            If t Like "BL##-*" Then Meld "Controleer zaaknummer in " & src.Cells(r, c).Address(False, False) & " van het bronblad: " & t
        Next c
    Next r

    '--- analysebladen ----------------------------------------------------------
    Set ws = SheetBy(wbk, "Ontwikkeling begrotingssaldo")
    If Not ws Is Nothing Then
        lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        For r = 1 To lr
            If Not ws.Cells(r, 1).HasFormula And CStr(ws.Cells(r, 1).Value) <> "" Then
                If Not (CStr(ws.Cells(r, 1).Value) Like "W##-##*") Then Zet ws.Cells(r, 1)
            End If
        Next r
        Meld "Ontwikkeling begrotingssaldo: controleer de vaste bedragen en wijzigingen in D6 en A10:D14 (die horen bij de vorige rapportage)."
    Else
        Meld "Tabblad 'Ontwikkeling begrotingssaldo' niet gevonden: niet aangepast."
    End If
    Set ws = SheetBy(wbk, "verschillenanalyses firap")
    If Not ws Is Nothing Then
        lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        For r = 1 To lr
            For c = 1 To 2
                If Not ws.Cells(r, c).HasFormula Then
                    t = CStr(ws.Cells(r, c).Value)
                    If t Like "Verschillenanalyses *" Or t Like "Bedrag * 20##" Or LCase$(t) Like "overhevelingen en aframingen naar *" Then Zet ws.Cells(r, c)
                End If
            Next c
        Next r
        Meld "verschillenanalyses firap: koppen aangepast. De omschrijvingen van de mutaties zelf horen bij de vorige rapportage; werk bij of bouw opnieuw op met de knop Verschillenanalyse."
    End If
    Set ws = SheetBy(wbk, "tbv adm pakket")
    If Not ws Is Nothing Then
        If CStr(ws.Range("A1").Value) = "Begrotingswijziging" And CStr(ws.Range("A4").Value) = "Grootboek" Then
            ws.Range("C1").Value = wz
            c = 6
            Do While CStr(ws.Cells(3, c).Value) <> "" And c <= 15
                ws.Cells(3, c).Value = nJr + (c - 6)
                c = c + 1
            Loop
            Meld "tbv adm pakket: jaarkolommen beginnen nu bij " & nJr & " en het wijzigingsnummer is " & wz & "."
        End If
    End If
    Set ws = SheetBy(wbk, "hulp verschillenanalyse")
    If Not ws Is Nothing Then
        ws.Range("B5").Value = nJr
        ws.Range("B6").Value = target
        If Not ws.Range("A1").HasFormula Then ws.Range("A1").Value = "Hulpblad verschillenanalyses, " & Left$(target, 1) & "e firap " & nJr
        Meld "hulp verschillenanalyse: begrotingsjaar (B5) en bronblad (B6) bijgewerkt."
    End If
    Meld "Tabblad 'tbv adm pakket' en 'controle inlees' worden bij de volgende run van de knop Inleesbestand opnieuw opgebouwd met de nieuwe jaren en het nieuwe wijzigingsnummer."
    If Not mWis Then Meld "Mutatiegegevens in het bronblad (bedragen, N/B, omschrijvingen, kolommen P en Q) zijn niet gewist: verwijder of vervang de mutaties van de vorige rapportage zelf."

    'begrotingssaldo en saldo onvoorzien uit het dashboard
    Set ws = SheetBy(wbk, "Ontwikkeling begrotingssaldo")
    If mWis Then WisMutaties wbk, src, ws
    If Not ws Is Nothing Then
        If sBeg <> "" Then ws.Range("D6").Value = CDbl(sBeg): Meld "Begrotingssaldo " & sBeg & " gezet in Ontwikkeling begrotingssaldo D6."
        If sOnv <> "" Then
            ws.Range("D20").Value = CDbl(sOnv)
            Meld "Saldo onvoorzien " & sOnv & " gezet in Ontwikkeling begrotingssaldo D20."
            Zet335 src, CDbl(sOnv)
        End If
    End If
    Application.ScreenUpdating = True
    Application.Calculate
    DashLogD wbk, "Overgang naar " & target & " (" & wz & ")"
    gResultaatD = "Overgang voltooid: '" & nm & "' wordt '" & target & "'. " & mAantal & " kopteksten aangepast." & vbCrLf & vbCrLf & Join2()
    If Not mStilD Then MsgBox gResultaatD, vbInformation, "Overgang naar nieuwe firap"
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    gResultaatD = "Fout " & Err.Number & ": " & Err.Description
    If Not mStilD Then MsgBox gResultaatD, vbCritical, "Overgang naar nieuwe firap"
End Sub

'past een kopcel aan: voorgaande en huidige rapportage, daarna jaartallen
Private Sub Zet(c As Range)
    Dim t As String, v As String
    t = CStr(c.Value)
    v = Vervang(t)
    If v <> t Then c.Value = v: mAantal = mAantal + 1
End Sub

Private Function Vervang(ByVal t As String) As String
    Dim f As Variant, forms As Variant
    forms = Array("financi" & ee & "le rapportage", "financiele rapportage", "firap")
    For Each f In forms
        t = Replace(t, "tot en met de " & oRap & "e " & f, "tot en met de " & Chr$(5) & f & Chr$(5), 1, -1, vbTextCompare)
        t = Replace(t, oPRap & "e " & f & " " & oPJr, Chr$(1) & f & Chr$(1), 1, -1, vbTextCompare)
        t = Replace(t, oRap & "e " & f & " " & oJr, Chr$(2) & f & Chr$(2), 1, -1, vbTextCompare)
    Next f
    t = Replace(t, CStr(oJr + 1), Chr$(3))
    t = Replace(t, CStr(oJr), Chr$(4))
    t = Replace(t, Chr$(3), CStr(nJr + 1))
    t = Replace(t, Chr$(4), CStr(nJr))
    For Each f In forms
        t = Replace(t, Chr$(1) & f & Chr$(1), nPRap & "e " & f & " " & nPJr)
        t = Replace(t, Chr$(2) & f & Chr$(2), nRap & "e " & f & " " & nJr)
        t = Replace(t, Chr$(5) & f & Chr$(5), nRap & "e " & f)
    Next f
    Vervang = t
End Function

Private Function IsWijz(ByVal t As String) As Boolean
    Dim p As Long
    If Len(t) < 4 Then Exit Function
    If UCase$(Left$(t, 1)) <> "W" Then Exit Function
    p = InStr(t, "/")
    If p < 3 Then Exit Function
    If Mid$(t, 2, p - 2) Like "*[!0-9]*" Then Exit Function
    If Mid$(t, p + 1) Like "*[!0-9]*" Then Exit Function
    IsWijz = (Len(Mid$(t, p + 1)) > 0)
End Function

Private Sub Meld(ByVal t As String)
    mMeldD.Add t
End Sub

Private Function Join2() As String
    Dim it As Variant, s As String
    For Each it In mMeldD
        s = s & "- " & CStr(it) & vbCrLf
    Next it
    Join2 = s
End Function

Private Sub Stopt(ByVal t As String)
    gResultaatD = t
    If Not mStilD Then MsgBox t, vbExclamation, "Overgang naar nieuwe firap"
End Sub

Private Function SheetBy(wbk As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wbk.Worksheets
        If LCase$(ws.Name) = LCase$(nm) Then Set SheetBy = ws: Exit Function
    Next ws
End Function

Private Sub DashLogD(wbk As Workbook, ByVal t As String)
    Dim d As Worksheet
    Set d = SheetBy(wbk, SH_DASH)
    If d Is Nothing Then Exit Sub
    d.Range("D17").Value = Format$(Now, "dd-mm-yyyy hh:mm") & "  " & t
End Sub

Public Function ResultaatD() As String
    ResultaatD = gResultaatD
End Function

'budget onvoorziene uitgaven (sectie 2 van het bronblad) = saldo onvoorzien van het dashboard
Private Sub Zet335(src As Worksheet, ByVal bedrag As Double)
    Dim r As Long, lr As Long
    lr = src.UsedRange.Row + src.UsedRange.Rows.Count - 1
    For r = 1 To lr
        If LCase$(Trim$(CStr(src.Cells(r, 1).Value))) Like "budget voor onvoorziene uitgaven na vaststelling*" Then
            If Not src.Cells(r, 6).HasFormula Then
                src.Cells(r, 5).Value = IIf(bedrag >= 0, "N", "B")
                src.Cells(r, 6).Value = Abs(bedrag)
                Meld "Bronblad F" & r & ": budget onvoorziene uitgaven gelijk gezet aan het saldo onvoorzien van het dashboard."
            End If
            Exit For
        End If
    Next r
End Sub

Private Function Bedr(src As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long, v As Variant
    For c = 6 To 14 Step 2
        v = src.Cells(r, c).Value
        If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
            If v <> 0 Then Bedr = True: Exit Function
        End If
    Next c
End Function

'wist alle mutaties: bronblad (structuur en formules blijven), analyse (leeg opnieuw), Ontwikkeling, tbv adm pakket
Private Sub WisMutaties(wbk As Workbook, src As Worksheet, wo As Worksheet)
    Dim r As Long, lr As Long, a As String, c As String, t As String
    Dim aE1 As Long, aSal As Long, aI1 As Long, aI2 As Long, aR1 As Long, aR2 As Long, aV1 As Long, aV2 As Long
    Dim k2 As Long, t2 As Long, i As Long, n As Long
    Dim r1(1 To 5) As Long, r2(1 To 5) As Long
    Dim isAmt() As Boolean, kop As Boolean, ws As Worksheet, nCl As Long
    lr = src.UsedRange.Row + src.UsedRange.Rows.Count - 1
    For r = 1 To lr
        a = LCase$(Trim$(CStr(src.Cells(r, 1).Value)))
        c = LCase$(Trim$(CStr(src.Cells(r, 3).Value)))
        If aE1 = 0 And a = "grootboek" Then aE1 = r + 1
        If aE1 > 0 And aSal = 0 And c Like "salderingsrekening begrotingssaldo*" Then aSal = r
        If k2 = 0 And aE1 > 0 And r >= aE1 And a Like "2. mutaties in onvoorziene*" Then k2 = r
        If k2 > 0 And t2 = 0 And r > k2 And a Like "totaal 2*" Then t2 = r
        If aI1 = 0 And a Like "3a. investeringen*" Then aI1 = r + 1
        If aI2 = 0 And a Like "totaal 3a*" Then aI2 = r - 1
        If aR1 = 0 And a = "reserves" Then aR1 = r + 1
        If aR2 = 0 And a Like "totaal 3b-1*" Then aR2 = r - 1
        If aV1 = 0 And a = "voorzieningen" Then aV1 = r + 1
        If aV2 = 0 And a Like "totaal 3b-2*" Then aV2 = r - 1
    Next r
    If aE1 = 0 Or aSal = 0 Or k2 = 0 Or t2 = 0 Or aI1 = 0 Or aI2 = 0 Or aR1 = 0 Or aR2 = 0 Or aV1 = 0 Or aV2 = 0 Then
        Meld "Mutaties NIET gewist: de ankers van het bronblad zijn niet allemaal gevonden."
        Exit Sub
    End If
    r1(1) = aE1: r2(1) = aSal - 1: r1(2) = aI1: r2(2) = aI2: r1(3) = aR1: r2(3) = aR2
    r1(4) = aV1: r2(4) = aV2: r1(5) = k2 + 1: r2(5) = t2 - 1
    ReDim isAmt(1 To lr + 2)
    For i = 1 To 5
        For r = r1(i) To r2(i)
            isAmt(r) = Bedr(src, r)
        Next r
    Next i
    For i = 1 To 5
        For r = r1(i) To r2(i)
            a = Trim$(CStr(src.Cells(r, 1).Value))
            c = Trim$(CStr(src.Cells(r, 3).Value))
            If isAmt(r) Then
                kop = (a Like "######" Or a Like "#######")
                src.Range(src.Cells(r, 4), src.Cells(r, 14)).ClearContents
                src.Range(src.Cells(r, 16), src.Cells(r, 17)).ClearContents
                If Not kop Then
                    src.Cells(r, 1).ClearContents
                    src.Cells(r, 3).ClearContents
                End If
                n = n + 1
            ElseIf a = "" And c <> "" And Not src.Cells(r, 3).HasFormula Then
                t = LCase$(c)
                If Not (t Like "programma*" Or t Like "taakveld*" Or t = "grondexploitaties") Then
                    If (r > r1(i) And isAmt(r - 1)) Or (r < r2(i) And isAmt(r + 1)) Then src.Cells(r, 3).ClearContents: nCl = nCl + 1
                End If
            End If
        Next r
    Next i
    Meld "Bronblad: " & n & " mutatieregels gewist (en " & nCl & " doorlopende toelichtingen). Kopregels, saldoregels en formules zijn niet aangeraakt."
    'Ontwikkeling begrotingssaldo
    If Not wo Is Nothing Then
        wo.Range("D6").ClearContents
        wo.Range("D20").ClearContents
        For r = 1 To 14
            t = Trim$(CStr(wo.Cells(r, 1).Value))
            If t Like "W##-##*" Then wo.Cells(r, 1).ClearContents: wo.Cells(r, 4).ClearContents
        Next r
        For r = 18 To 24
            If Trim$(CStr(wo.Cells(r, 1).Value)) Like "Af:*" And Not wo.Cells(r, 1).HasFormula Then wo.Cells(r, 1).Value = "Af: mutaties onvoorziene uitgaven in deze rapportage"
        Next r
    End If
    'tbv adm pakket: alleen de inleesregels wissen
    Set ws = SheetBy(wbk, "tbv adm pakket")
    If Not ws Is Nothing Then
        If CStr(ws.Range("A1").Value) = "Begrotingswijziging" And CStr(ws.Range("A4").Value) = "Grootboek" Then
            lr = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If lr >= 5 Then ws.Range(ws.Cells(5, 1), ws.Cells(lr, 16)).Clear
            Meld "tbv adm pakket: inleesregels gewist (koppen blijven)."
        End If
    End If
    'analyse leeg opnieuw opbouwen
    VerschillenVolledigStil
    Meld "verschillenanalyses firap: leeg opnieuw opgebouwd."
End Sub
