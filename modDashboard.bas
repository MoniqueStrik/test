Attribute VB_Name = "modDashboard"
Option Explicit
'==============================================================================
' Dashboard "Inlees en analyse"
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

Private Const SH_DASH As String = "Inlees en analyse"
Public gResultaatD As String
Private mStilD As Boolean
Private mMeldD As Collection
Private mAantal As Long
Private oRap As Long, oJr As Long, nRap As Long, nJr As Long
Private oPRap As Long, oPJr As Long, nPRap As Long, nPJr As Long
Private ee As String

Public Sub OvergangNaarNieuweFirap()
    mStilD = False
    DoOvergang
End Sub

Public Sub OvergangStil()
    mStilD = True
    DoOvergang
End Sub

Private Sub DoOvergang()
    Dim wbk As Workbook, d As Worksheet, src As Worksheet, ws As Worksheet
    Dim rap As String, jr As String, wz As String, target As String, k As Long
    Dim lr As Long, r As Long, c As Long, t As String, nieuw As String
    Dim antw As VbMsgBoxResult, nm As String, i As Long
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
            "Tip: sla eerst een kopie van het bestand op. Doorgaan?", vbYesNo + vbQuestion, "Overgang naar nieuwe firap")
        If antw <> vbYes Then Exit Sub
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
        Meld "Ontwikkeling begrotingssaldo: controleer de datums in A1 en A15 en de vaste bedragen en wijzigingen in D6 en A10:D14 (die horen bij de vorige rapportage)."
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
    Set ws = SheetBy(wbk, "hulp verschillenanalyse")
    If Not ws Is Nothing Then
        ws.Range("B5").Value = nJr
        ws.Range("B6").Value = target
        If Not ws.Range("A1").HasFormula Then ws.Range("A1").Value = "Hulpblad verschillenanalyses, " & Left$(target, 1) & "e firap " & nJr
        Meld "hulp verschillenanalyse: begrotingsjaar (B5) en bronblad (B6) bijgewerkt."
    End If
    Meld "Tabblad 'tbv Coda' en 'controle inlees' worden bij de volgende run van de knop Inleesbestand opnieuw opgebouwd met de nieuwe jaren en het nieuwe wijzigingsnummer."
    Meld "Mutatiegegevens in het bronblad (bedragen, N/B, omschrijvingen, kolommen P en Q) zijn niet gewist: verwijder of vervang de mutaties van de vorige rapportage zelf."

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
