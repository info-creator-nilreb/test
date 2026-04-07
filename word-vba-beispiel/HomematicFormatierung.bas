Option Explicit

' ============================================================================
' Homematic-Dokumentation formatieren und als neue Datei speichern
' Word Desktop VBA
' ============================================================================
'
' Nutzung:
' 1) Beispiel-Dokument in denselben Ordner wie diese .bas-Datei legen
' 2) Makro in Word importieren und ausführen
' 3) Das aktive Dokument wird formatiert und als neue Datei gespeichert:
'    <Originalname>_formatiert.docx
' ============================================================================

Public Sub FormatHomematicDokumentationUndSpeichern()
    Dim doc As Document
    Dim sourcePath As String
    Dim targetPath As String

    Set doc = ActiveDocument

    If Len(doc.Path) = 0 Then
        MsgBox "Bitte speichere das Dokument zuerst als .docx, bevor du das Makro ausführst.", vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanFail

    sourcePath = doc.FullName
    targetPath = BuildFormattedPath(sourcePath)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone

    ' 1) Dokumentweite Formatierungen
    FormatHeaderLines doc
    FormatObjectLines doc

    ' 2) Makrobereiche verarbeiten
    ProcessMacroSections doc

    ' 3) Als neue Datei speichern
    doc.SaveAs2 FileName:=targetPath, FileFormat:=wdFormatXMLDocument

    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True

    MsgBox "Formatierung abgeschlossen. Neue Datei gespeichert unter:" & vbCrLf & targetPath, vbInformation
    Exit Sub

CleanFail:
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    MsgBox "Fehler bei der Formatierung: " & Err.Description, vbCritical
End Sub

Private Function BuildFormattedPath(ByVal sourcePath As String) As String
    Dim dotPos As Long

    dotPos = InStrRev(sourcePath, ".")

    If dotPos > 0 Then
        BuildFormattedPath = Left$(sourcePath, dotPos - 1) & "_formatiert.docx"
    Else
        BuildFormattedPath = sourcePath & "_formatiert.docx"
    End If
End Function

' ============================================================================
' Dokumentweite Kopfzeilen
' ============================================================================

Private Sub FormatHeaderLines(ByVal doc As Document)
    Dim i As Long
    Dim paraText As String
    Dim rng As Range
    Dim paraCount As Long

    paraCount = doc.Paragraphs.Count

    For i = 1 To paraCount
        Set rng = doc.Paragraphs(i).Range.Duplicate
        paraText = CleanParagraphText(rng.Text)

        If StartsWith(paraText, "Projekt:") _
        Or StartsWith(paraText, "Datum:") _
        Or StartsWith(paraText, "Anzahl Objekte:") _
        Or StartsWith(paraText, "Liste Objekte") Then

            With rng.Font
                .Name = "Tahoma"
                .Size = 11
                .Bold = True
                .Italic = False
                .Underline = wdUnderlineNone
                .Color = wdColorBlue
            End With
        End If
    Next i
End Sub

' ============================================================================
' Dokumentweite Objektzeilen
' ============================================================================

Private Sub FormatObjectLines(ByVal doc As Document)
    Dim i As Long
    Dim txt As String
    Dim rng As Range
    Dim rngLabel As Range
    Dim paraCount As Long

    paraCount = doc.Paragraphs.Count

    For i = 1 To paraCount
        Set rng = doc.Paragraphs(i).Range.Duplicate
        txt = CleanParagraphText(rng.Text)

        If StartsWith(LTrim$(txt), "Objekt:") Then
            With rng.Font
                .Name = "Tahoma"
                .Size = 11
                .Bold = True
                .Italic = False
                .Underline = wdUnderlineNone
                .Color = GetAccentRedColor()
            End With

            Set rngLabel = rng.Duplicate
            rngLabel.End = rngLabel.Start + Len("Objekt:")

            With rngLabel.Font
                .Underline = wdUnderlineSingle
            End With
        End If
    Next i
End Sub

' ============================================================================
' Makrobereiche erkennen und verarbeiten
' ============================================================================

Private Sub ProcessMacroSections(ByVal doc As Document)
    Dim i As Long
    Dim paraCount As Long
    Dim startIdx As Long
    Dim endIdx As Long
    Dim txt As String

    paraCount = doc.Paragraphs.Count
    i = 1

    Do While i <= paraCount
        txt = CleanParagraphText(doc.Paragraphs(i).Range.Text)

        If IsMacroHeader(txt) Then
            startIdx = FindMacroContentStart(doc, i + 1)

            If startIdx > 0 Then
                endIdx = FindMacroContentEnd(doc, startIdx)

                If endIdx >= startIdx Then
                    ProcessSingleMacroBlock doc, startIdx, endIdx
                    i = endIdx
                End If
            End If
        End If

        i = i + 1
    Loop
End Sub

Private Function IsMacroHeader(ByVal txt As String) As Boolean
    Dim t As String

    t = Trim$(txt)

    IsMacroHeader = (t = "Makro") _
                 Or (t = "Makro:") _
                 Or (t = "Macro") _
                 Or (t = "Macro:")
End Function

Private Function FindMacroContentStart(ByVal doc As Document, ByVal fromIdx As Long) As Long
    Dim i As Long
    Dim txt As String
    Dim t As String

    FindMacroContentStart = 0

    For i = fromIdx To doc.Paragraphs.Count
        txt = CleanParagraphText(doc.Paragraphs(i).Range.Text)
        t = Trim$(txt)

        If Len(t) = 0 Then
            ' Leere Absätze überspringen

        ElseIf StartsWith(t, "Ausführung bei") Then
            ' Kopfzeile innerhalb Makroblock überspringen

        ElseIf StartsWith(t, "Verbunden mit Anschluss") _
            Or StartsWith(t, "Raum:") _
            Or StartsWith(t, "Objekt:") _
            Or StartsWith(t, "Typ:") _
            Or StartsWith(t, "Definition") _
            Or StartsWith(t, "Definitionen") _
            Or StartsWith(t, "Hardware-Modul:") _
            Or IsSeparatorLine(t) Then

            Exit Function

        Else
            FindMacroContentStart = i
            Exit Function
        End If
    Next i
End Function

Private Function FindMacroContentEnd(ByVal doc As Document, ByVal startIdx As Long) As Long
    Dim i As Long
    Dim txt As String
    Dim t As String
    Dim lastContent As Long

    lastContent = startIdx

    For i = startIdx To doc.Paragraphs.Count
        txt = CleanParagraphText(doc.Paragraphs(i).Range.Text)
        t = Trim$(txt)

        If StartsWith(t, "Verbunden mit Anschluss") _
        Or StartsWith(t, "Raum:") _
        Or StartsWith(t, "Objekt:") _
        Or StartsWith(t, "Typ:") _
        Or StartsWith(t, "Definition") _
        Or StartsWith(t, "Definitionen") _
        Or StartsWith(t, "Hardware-Modul:") _
        Or IsMacroHeader(t) Then
            Exit For
        Else
            lastContent = i
        End If
    Next i

    FindMacroContentEnd = lastContent
End Function

Private Sub ProcessSingleMacroBlock(ByVal doc As Document, ByVal startIdx As Long, ByVal endIdx As Long)
    Dim i As Long

    For i = startIdx To endIdx
        FormatMacroCommentLine doc.Paragraphs(i)
    Next i

    For i = startIdx To endIdx
        FormatKeywordsInParagraph doc.Paragraphs(i)
    Next i

    ResetIndentation doc, startIdx, endIdx
    RecalculateIndentation doc, startIdx, endIdx
End Sub

' ============================================================================
' Kommentarzeilen im Makrocode
' ============================================================================

Private Sub FormatMacroCommentLine(ByVal para As Paragraph)
    Dim txt As String
    Dim t As String
    Dim rng As Range

    Set rng = para.Range.Duplicate
    txt = CleanParagraphText(rng.Text)
    t = LTrim$(txt)

    If StartsWith(t, "//") Then
        With rng.Font
            .Name = "Tahoma"
            .Size = 8
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
            .Color = GetCommentBlueColor()
        End With
    End If
End Sub

' ============================================================================
' Schlüsselwörter im Makrocode
' ============================================================================

Private Sub FormatKeywordsInParagraph(ByVal para As Paragraph)
    Dim rng As Range

    Set rng = para.Range.Duplicate

    FormatWordInRange rng, "wenn"
    FormatWordInRange rng, "dann"
    FormatWordInRange rng, "sonst"
    FormatWordInRange rng, "endewenn"
End Sub

Private Sub FormatWordInRange(ByVal targetRange As Range, ByVal keyword As String)
    Dim rngFind As Range

    Set rngFind = targetRange.Duplicate

    With rngFind.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = keyword
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
    End With

    Do While rngFind.Find.Execute
        With rngFind.Font
            .Name = "Tahoma"
            .Size = 9
            .Bold = True
            .Italic = True
            .Underline = wdUnderlineNone
            .Color = GetKeywordGreenColor()
        End With

        rngFind.Collapse wdCollapseEnd
    Loop
End Sub

' ============================================================================
' Einrückung
' ============================================================================

Private Sub ResetIndentation(ByVal doc As Document, ByVal startIdx As Long, ByVal endIdx As Long)
    Dim i As Long

    For i = startIdx To endIdx
        With doc.Paragraphs(i).Range.ParagraphFormat
            .LeftIndent = 0
            .FirstLineIndent = 0
        End With
    Next i
End Sub

Private Sub RecalculateIndentation(ByVal doc As Document, ByVal startIdx As Long, ByVal endIdx As Long)
    Dim i As Long
    Dim level As Long
    Dim txt As String
    Dim t As String

    Const INDENT_STEP As Single = 18

    level = 0

    For i = startIdx To endIdx
        txt = CleanParagraphText(doc.Paragraphs(i).Range.Text)
        t = LTrim$(txt)

        If Len(Trim$(t)) = 0 Then
            ApplyIndent doc.Paragraphs(i), level * INDENT_STEP

        ElseIf StartsWith(t, "endewenn") Then
            If level > 0 Then level = level - 1
            ApplyIndent doc.Paragraphs(i), level * INDENT_STEP

        ElseIf StartsWith(t, "sonst") Then
            If level > 0 Then level = level - 1
            ApplyIndent doc.Paragraphs(i), level * INDENT_STEP
            level = level + 1

        Else
            ApplyIndent doc.Paragraphs(i), level * INDENT_STEP

            If StartsWith(t, "wenn") Then
                level = level + 1
            End If
        End If
    Next i
End Sub

Private Sub ApplyIndent(ByVal para As Paragraph, ByVal leftIndentPoints As Single)
    With para.Range.ParagraphFormat
        .LeftIndent = leftIndentPoints
        .FirstLineIndent = 0
    End With
End Sub

' ============================================================================
' Hilfsfunktionen
' ============================================================================

Private Function CleanParagraphText(ByVal s As String) As String
    s = Replace$(s, vbCr, "")
    s = Replace$(s, Chr$(13), "")
    s = Replace$(s, Chr$(7), "")
    s = Replace$(s, Chr$(11), "")
    s = Replace$(s, Chr$(12), "")
    CleanParagraphText = Trim$(s)
End Function

Private Function StartsWith(ByVal text As String, ByVal prefix As String) As Boolean
    If Len(text) < Len(prefix) Then
        StartsWith = False
    Else
        StartsWith = (Left$(text, Len(prefix)) = prefix)
    End If
End Function

Private Function IsSeparatorLine(ByVal txt As String) As Boolean
    Dim t As String

    t = Replace$(Trim$(txt), "=", "")
    IsSeparatorLine = (Len(Trim$(txt)) > 0 And Len(t) = 0)
End Function

Private Function GetAccentRedColor() As Long
    ' Näherung für "Rot Akzent 2 dunkler 25%"
    GetAccentRedColor = RGB(192, 80, 77)
End Function

Private Function GetKeywordGreenColor() As Long
    ' Näherung für "Olivgrün Akzent 3 dunkler 25%"
    GetKeywordGreenColor = RGB(118, 146, 60)
End Function

Private Function GetCommentBlueColor() As Long
    ' Näherung für "Dunkelblau Text 2 heller 40%"
    GetCommentBlueColor = RGB(79, 129, 189)
End Function
