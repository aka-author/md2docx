'===========================================================
' bookmarx.vbs — html2docx: part_* insert, style_*, label removal
' Optional 4th arg: log file (default: <html-dir>\<html-base>.bookmarx.log)
'===========================================================

Option Explicit

Dim gLogPath

Sub LogAppend(line)
  Dim fso, ts
  If Len(gLogPath) = 0 Then Exit Sub
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.OpenTextFile(gLogPath, 8, True)
  ts.WriteLine line
  ts.Close
End Sub

Sub LogMsg(level, msg)
  Dim line
  line = Now & " [" & level & "] " & msg
  WScript.Echo line
  LogAppend line
End Sub

Sub LogOpen(path)
  Dim fso, folder, ts
  Set fso = CreateObject("Scripting.FileSystemObject")
  folder = fso.GetParentFolderName(path)
  If Len(folder) > 0 And Not fso.FolderExists(folder) Then
    fso.CreateFolder folder
  End If
  gLogPath = path
  Set ts = fso.CreateTextFile(path, True)
  ts.Close
  LogMsg "INFO", "Log file: " & path
End Sub

Sub LogClose()
  gLogPath = ""
End Sub

Sub LogFileCheck(label, path)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FileExists(path) Then
    LogMsg "ERROR", label & " missing: " & path
    WScript.Quit 10
  End If
  LogMsg "INFO", label & " OK size=" & fso.GetFile(path).Size & " path=" & path
End Sub

Function GetArg(i, name)
  If WScript.Arguments.Count <= i Then
    WScript.Echo "Missing argument: " & name
    WScript.Quit 1
  End If
  GetArg = WScript.Arguments(i)
End Function

Function ResolvePath(p)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  If Left(p, 1) = "\" Or InStr(p, ":") > 0 Then
    ResolvePath = fso.GetAbsolutePathName(p)
  Else
    ResolvePath = fso.GetAbsolutePathName(".\" & p)
  End If
End Function

Function DefaultLogPath(htmlFile)
  Dim fso, base
  Set fso = CreateObject("Scripting.FileSystemObject")
  base = fso.GetBaseName(htmlFile)
  DefaultLogPath = fso.BuildPath(fso.GetParentFolderName(htmlFile), base & ".bookmarx.log")
End Function

Function CopyForWordOpen(sourcePath, prefix)
  Dim fso, tempPath
  Set fso = CreateObject("Scripting.FileSystemObject")
  tempPath = fso.BuildPath(fso.GetSpecialFolder(2), prefix & "_" & fso.GetBaseName(sourcePath))
  If LCase(fso.GetExtensionName(sourcePath)) <> "" Then
    tempPath = tempPath & "." & fso.GetExtensionName(sourcePath)
  End If
  fso.CopyFile sourcePath, tempPath, True
  LogMsg "INFO", "Copy for Open (avoid recovery prompt): " & tempPath
  CopyForWordOpen = tempPath
End Function

Function CreateWord()
  Dim w
  LogMsg "INFO", "Word COM: CreateObject(Word.Application) — wait..."
  Set w = CreateObject("Word.Application")
  w.Visible = False
  w.DisplayAlerts = 0
  On Error Resume Next
  w.ScreenUpdating = False
  w.Options.SaveNormalPrompt = False
  w.Options.ConfirmConversions = False
  w.Options.UpdateLinksAtOpen = False
  w.Options.CheckSpellingAsYouType = False
  w.Options.CheckGrammarAsYouType = False
  On Error GoTo 0
  LogMsg "INFO", "Word COM: ready version=" & w.Version
  Set CreateWord = w
End Function

Sub QuitWord(w)
  On Error Resume Next
  w.Quit
  On Error GoTo 0
End Sub

Function OpenDoc(w, path, readOnly)
  LogMsg "INFO", "Documents.Open readOnly=" & readOnly & " path=" & path
  Set OpenDoc = w.Documents.Open(path, False, readOnly)
  LogMsg "INFO", "Documents.Open OK"
End Function

Sub CloseDoc(doc, saveChanges)
  On Error Resume Next
  doc.Close saveChanges
  On Error GoTo 0
End Sub

Sub SaveDocAs(doc, outPath)
  Const wdFormatXMLDocument = 12
  On Error Resume Next
  doc.SaveAs2 outPath, wdFormatXMLDocument
  If Err.Number <> 0 Then
    LogMsg "ERROR", "SaveAs failed: " & outPath & " — " & Err.Description
    WScript.Quit 3
  End If
  On Error GoTo 0
  LogMsg "INFO", "Saved: " & outPath
End Sub

Function IsPartBookmarkName(name)
  IsPartBookmarkName = (Len(name) >= 6 And LCase(Left(name, 5)) = "part_")
End Function

' Closing bookmark names start with "_", which Word treats as hidden:
' Bookmarks.ShowHidden must be True to find them.
Function CloseName(name)
  CloseName = "_" & name
End Function

Function IsServiceBookmarkName(name)
  Dim base
  base = name
  If Left(base, 1) = "_" Then base = Mid(base, 2)
  IsServiceBookmarkName = IsPartBookmarkName(base) Or IsStyleBookmarkOpening(base)
End Function

Function IsStyleBookmarkOpening(name)
  If Len(name) < 7 Then
    IsStyleBookmarkOpening = False
    Exit Function
  End If
  If Left(name, 1) = "_" Then
    IsStyleBookmarkOpening = False
    Exit Function
  End If
  IsStyleBookmarkOpening = (LCase(Left(name, 6)) = "style_")
End Function

Function BookmarkLabelText(bm)
  Dim t
  t = bm.Range.Text
  t = Replace(t, vbCr, "")
  t = Replace(t, vbLf, "")
  BookmarkLabelText = t
End Function

Function PartTemplateSuffix(templatePartName)
  PartTemplateSuffix = Mid(templatePartName, 6)
End Function

Function FindPartHtmlOpeningName(htmlDoc, templatePartName)
  Dim bm, want
  FindPartHtmlOpeningName = ""
  want = PartTemplateSuffix(templatePartName)
  For Each bm In htmlDoc.Bookmarks
    If IsPartBookmarkName(bm.Name) Then
      If BookmarkLabelText(bm) = want Then
        If htmlDoc.Bookmarks.Exists(CloseName(bm.Name)) Then
          FindPartHtmlOpeningName = bm.Name
          Exit Function
        End If
      End If
    End If
  Next
End Function

Function PairRange(doc, name)
  Set PairRange = Nothing
  If Not doc.Bookmarks.Exists(name) Then Exit Function
  If Not doc.Bookmarks.Exists(CloseName(name)) Then Exit Function
  Set PairRange = doc.Range(doc.Bookmarks(name).Range.Start, doc.Bookmarks(CloseName(name)).Range.End)
End Function

Sub CollectTemplateZonePartNames(doc, templateZoneParts)
  Dim bm
  For Each bm In doc.Bookmarks
    If IsPartBookmarkName(bm.Name) Then
      templateZoneParts(bm.Name) = True
    End If
  Next
End Sub

Function IsHtmlAnchorBookmark(name, templateZoneParts)
  Dim base
  If Not IsServiceBookmarkName(name) Then
    IsHtmlAnchorBookmark = False
    Exit Function
  End If
  base = name
  If Left(base, 1) = "_" Then base = Mid(base, 2)
  If templateZoneParts.Exists(base) Then
    IsHtmlAnchorBookmark = False
    Exit Function
  End If
  IsHtmlAnchorBookmark = True
End Function

Function ParagraphWhollyInside(pStart, pEnd, rStart, rEnd)
  ParagraphWhollyInside = (pStart >= rStart And pEnd <= rEnd)
End Function

Sub RemoveServiceLabels(doc, templateZoneParts)
  Dim bm, names(), starts(), ends(), count, i, j, tmp, rng, para, removed, emptied

  On Error Resume Next
  count = 0
  For Each bm In doc.Bookmarks
    If IsHtmlAnchorBookmark(bm.Name, templateZoneParts) Then
      If bm.Range.End > bm.Range.Start Then
        ReDim Preserve names(count)
        ReDim Preserve starts(count)
        ReDim Preserve ends(count)
        names(count) = bm.Name
        starts(count) = bm.Range.Start
        ends(count) = bm.Range.End
        count = count + 1
      End If
    End If
  Next

  For i = 0 To count - 2
    For j = i + 1 To count - 1
      If starts(j) > starts(i) Then
        tmp = starts(i): starts(i) = starts(j): starts(j) = tmp
        tmp = ends(i): ends(i) = ends(j): ends(j) = tmp
        tmp = names(i): names(i) = names(j): names(j) = tmp
      End If
    Next
  Next

  removed = 0
  emptied = 0
  For i = 0 To count - 1
    Set rng = doc.Range(starts(i), ends(i))
    Set para = rng.Paragraphs(1).Range
    rng.Delete
    removed = removed + 1
    If Trim(Replace(para.Text, vbCr, "")) = "" Then
      para.Delete
      emptied = emptied + 1
    End If
    If Err.Number <> 0 Then
      LogMsg "ERROR", "RemoveServiceLabels " & names(i) & " at " & starts(i) & ": Err=" & Err.Number & " " & Err.Description
      Err.Clear
    End If
  Next
  On Error GoTo 0

  LogMsg "INFO", "Html anchor bookmarks cleared=" & removed & " empty paragraphs removed=" & emptied
End Sub

Sub CollectPartBookmarkNamesDesc(doc, names)
  Dim bm, count, starts(), i, j, tmpN, tmpS
  count = 0
  For Each bm In doc.Bookmarks
    If IsPartBookmarkName(bm.Name) Then
      ReDim Preserve names(count)
      ReDim Preserve starts(count)
      names(count) = bm.Name
      starts(count) = bm.Range.Start
      count = count + 1
      LogMsg "INFO", "Template part bookmark: " & bm.Name & " start=" & bm.Range.Start & " end=" & bm.Range.End
    End If
  Next
  If count = 0 Then
    ReDim names(-1)
    Exit Sub
  End If
  For i = 0 To count - 2
    For j = i + 1 To count - 1
      If starts(j) > starts(i) Then
        tmpS = starts(i)
        starts(i) = starts(j)
        starts(j) = tmpS
        tmpN = names(i)
        names(i) = names(j)
        names(j) = tmpN
      End If
    Next
  Next
  LogMsg "INFO", "Insert order (desc): " & Join(names, ", ")
End Sub

Sub PastePartFromHtmlDoc(htmlDoc, doc, partName)
  Dim insertStart, insertEnd, src, rng, htmlOpenName

  On Error Resume Next

  If Not doc.Bookmarks.Exists(partName) Then
    LogMsg "WARN", "Part " & partName & ": no bookmark in template — skip"
    Exit Sub
  End If

  htmlOpenName = FindPartHtmlOpeningName(htmlDoc, partName)
  If Len(htmlOpenName) = 0 Then
    LogMsg "WARN", "Part " & partName & ": no HTML pair with label [" & PartTemplateSuffix(partName) & "] — skip"
    Exit Sub
  End If

  Set src = PairRange(htmlDoc, htmlOpenName)
  If src Is Nothing Then
    LogMsg "WARN", "Part " & partName & ": incomplete HTML pair " & htmlOpenName & "/" & CloseName(htmlOpenName) & " — skip"
    Exit Sub
  End If

  insertStart = doc.Bookmarks(partName).Range.Start
  insertEnd = doc.Bookmarks(partName).Range.End
  LogMsg "INFO", "Part " & partName & ": HTML " & htmlOpenName & " src " & src.Start & ".." & src.End & " -> target " & insertStart & ".." & insertEnd

  If src.End <= src.Start Then
    LogMsg "WARN", "Part " & partName & ": empty source range — skip"
    Exit Sub
  End If

  If insertEnd > insertStart Then
    If doc.Range(insertEnd - 1, insertEnd).Text = vbCr And src.Characters.Last.Text <> vbCr Then
      insertEnd = insertEnd - 1
      LogMsg "INFO", "Part " & partName & ": template paragraph mark kept, target " & insertStart & ".." & insertEnd
    End If
  End If

  Set rng = doc.Range(insertStart, insertEnd)
  rng.FormattedText = src.FormattedText
  If Err.Number <> 0 Then
    LogMsg "ERROR", "Part " & partName & ": insert failed Err=" & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
  End If

  LogMsg "INFO", "Part " & partName & ": insert OK"
End Sub

'-----------------------------------------------------------
' Style bookmarks
'-----------------------------------------------------------

Const wdStyleTypeParagraph = 1
Const wdStyleTypeCharacter = 2
Const wdMainTextStory = 1

Function TryGetStyle(doc, styleName)
  Dim st
  On Error Resume Next
  Set st = doc.Styles(styleName)
  If Err.Number <> 0 Then
    Set st = Nothing
    Err.Clear
  End If
  On Error GoTo 0
  Set TryGetStyle = st
End Function

Sub ApplyStyleAtBookmark(doc, bmName, applied, failed)
  Dim st, styleName, rng, p, openStart, openEnd, closeStart, closeEnd, pStart, pEnd, applyRng

  On Error Resume Next

  Set rng = PairRange(doc, bmName)
  If rng Is Nothing Then
    LogMsg "WARN", "Style " & bmName & ": no bookmark pair " & bmName & "/" & CloseName(bmName) & " — skip"
    failed = failed + 1
    Exit Sub
  End If

  openStart = doc.Bookmarks(bmName).Range.Start
  openEnd = doc.Bookmarks(bmName).Range.End
  closeStart = doc.Bookmarks(CloseName(bmName)).Range.Start
  closeEnd = doc.Bookmarks(CloseName(bmName)).Range.End

  styleName = BookmarkLabelText(doc.Bookmarks(bmName))
  Set st = TryGetStyle(doc, styleName)
  If st Is Nothing Then
    LogMsg "WARN", "Style " & bmName & ": no style [" & styleName & "] in document — skip"
    failed = failed + 1
    Exit Sub
  End If

  Select Case st.Type
    Case wdStyleTypeParagraph
      For Each p In rng.Paragraphs
        pStart = p.Range.Start
        pEnd = p.Range.End
        If Not ParagraphWhollyInside(pStart, pEnd, openStart, openEnd) And Not ParagraphWhollyInside(pStart, pEnd, closeStart, closeEnd) Then
          p.Range.Style = st
        End If
      Next
    Case wdStyleTypeCharacter
      If closeStart > openEnd Then
        Set applyRng = doc.Range(openEnd, closeStart)
        applyRng.Style = st
      End If
    Case Else
      LogMsg "WARN", "Style " & bmName & ": unsupported style type " & st.Type & " — skip"
      failed = failed + 1
      Exit Sub
  End Select

  If Err.Number <> 0 Then
    LogMsg "ERROR", "Style " & bmName & " [" & styleName & "]: apply failed Err=" & Err.Number & " " & Err.Description
    Err.Clear
    failed = failed + 1
    Exit Sub
  End If

  applied = applied + 1
End Sub

Sub ProcessStyleBookmarks(doc)
  Dim bm, names(), count, i, applied, failed

  count = 0
  On Error Resume Next
  For Each bm In doc.Bookmarks
    If bm.Range.StoryType = wdMainTextStory Then
      If IsStyleBookmarkOpening(bm.Name) Then
        ReDim Preserve names(count)
        names(count) = bm.Name
        count = count + 1
      End If
    End If
  Next
  Err.Clear
  On Error GoTo 0

  LogMsg "INFO", "Style bookmarks found: " & count

  applied = 0
  failed = 0
  For i = 0 To count - 1
    ApplyStyleAtBookmark doc, names(i), applied, failed
  Next

  LogMsg "INFO", "Style bookmarks applied=" & applied & " skipped=" & failed
End Sub

Sub InsertPartsOnly(htmlDoc, doc)
  Dim names(), i, bmName
  ReDim names(-1)
  CollectPartBookmarkNamesDesc doc, names
  If UBound(names) < 0 Then
    LogMsg "WARN", "No part_* bookmarks in template — nothing to insert"
    Exit Sub
  End If
  For i = 0 To UBound(names)
    bmName = names(i)
    PastePartFromHtmlDoc htmlDoc, doc, bmName
  Next
End Sub

'--- MAIN ---
Dim htmlPath, inDocx, outDocx, logPath, w, doc, htmlDoc, workDocx, workHtml, templateZoneParts

WScript.Echo "bookmarx | host: " & WScript.FullName

htmlPath = ResolvePath(GetArg(0, "HTML file"))
If WScript.Arguments.Count > 3 Then
  logPath = ResolvePath(GetArg(3, "Log file"))
Else
  logPath = DefaultLogPath(htmlPath)
End If
LogOpen logPath

inDocx   = ResolvePath(GetArg(1, "Input DOCX"))
outDocx  = ResolvePath(GetArg(2, "Output DOCX"))
LogMsg "INFO", "bookmarx start"
LogMsg "INFO", "HTML=" & htmlPath
LogMsg "INFO", "Template=" & inDocx
LogMsg "INFO", "Output=" & outDocx

LogFileCheck "HTML", htmlPath
LogFileCheck "Template", inDocx

Set w = CreateWord()
workDocx = CopyForWordOpen(inDocx, "bookmarx_tpl")
workHtml = CopyForWordOpen(htmlPath, "bookmarx_html")
Set doc = OpenDoc(w, workDocx, False)
Set htmlDoc = OpenDoc(w, workHtml, True)
doc.Bookmarks.ShowHidden = True
htmlDoc.Bookmarks.ShowHidden = True
LogMsg "INFO", "Template+HTML opened from TEMP copies (same flow as PasteHtmlIntoDoc)"

Set templateZoneParts = CreateObject("Scripting.Dictionary")
CollectTemplateZonePartNames doc, templateZoneParts

LogMsg "INFO", "InsertPartsOnly begin"
On Error Resume Next
InsertPartsOnly htmlDoc, doc
If Err.Number <> 0 Then
  LogMsg "ERROR", "InsertPartsOnly aborted Err=" & Err.Number & " " & Err.Description & " — output still saved"
  Err.Clear
End If
On Error GoTo 0
LogMsg "INFO", "InsertPartsOnly end"
CloseDoc htmlDoc, False

LogMsg "INFO", "ProcessStyleBookmarks begin"
On Error Resume Next
ProcessStyleBookmarks doc
If Err.Number <> 0 Then
  LogMsg "ERROR", "ProcessStyleBookmarks aborted Err=" & Err.Number & " " & Err.Description & " — output still saved"
  Err.Clear
End If
On Error GoTo 0
LogMsg "INFO", "ProcessStyleBookmarks end"

RemoveServiceLabels doc, templateZoneParts

SaveDocAs doc, outDocx
CloseDoc doc, False
LogMsg "INFO", "Word COM: Quit"
QuitWord w
On Error Resume Next
CreateObject("Scripting.FileSystemObject").DeleteFile workHtml
CreateObject("Scripting.FileSystemObject").DeleteFile workDocx
On Error GoTo 0

LogMsg "INFO", "bookmarx done"
LogClose
