'===========================================================
' bookmarx.vbs — part_* insert only (diagnostic); 3 args from pipeline
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

' Each service anchor holds MARKER so that Word imports it as a non-empty
' bookmark and keeps it when the fragment is copied.
Const MARKER = "@#$$#@"

Function IsServiceBookmarkName(name)
  Dim base
  base = name
  If Left(base, 1) = "_" Then base = Mid(base, 2)
  IsServiceBookmarkName = IsPartBookmarkName(base) Or (LCase(Left(base, 6)) = "style_")
End Function

Function PairRange(doc, name)
  Set PairRange = Nothing
  If Not doc.Bookmarks.Exists(name) Then Exit Function
  If Not doc.Bookmarks.Exists(CloseName(name)) Then Exit Function
  Set PairRange = doc.Range(doc.Bookmarks(name).Range.Start, doc.Bookmarks(CloseName(name)).Range.End)
End Function

Sub RemoveMarkers(doc)
  Dim bm, starts(), ends(), count, i, j, tmp, rng, para, removed, emptied

  On Error Resume Next
  count = 0
  For Each bm In doc.Bookmarks
    If IsServiceBookmarkName(bm.Name) Then
      If doc.Range(bm.Range.Start, bm.Range.Start + Len(MARKER)).Text = MARKER Then
        ReDim Preserve starts(count)
        ReDim Preserve ends(count)
        starts(count) = bm.Range.Start
        ends(count) = bm.Range.Start + Len(MARKER)
        count = count + 1
      End If
    End If
  Next
  For i = 0 To count - 2
    For j = i + 1 To count - 1
      If starts(j) > starts(i) Then
        tmp = starts(i): starts(i) = starts(j): starts(j) = tmp
        tmp = ends(i): ends(i) = ends(j): ends(j) = tmp
      End If
    Next
  Next

  removed = 0
  emptied = 0
  For i = 0 To count - 1
    Set rng = doc.Range(starts(i), ends(i))
    If rng.Text = MARKER Then
      Set para = rng.Paragraphs(1).Range
      rng.Delete
      removed = removed + 1
      If Trim(Replace(para.Text, vbCr, "")) = "" Then
        para.Delete
        emptied = emptied + 1
      End If
    End If
    If Err.Number <> 0 Then
      LogMsg "ERROR", "RemoveMarkers at " & starts(i) & ": Err=" & Err.Number & " " & Err.Description
      Err.Clear
    End If
  Next
  On Error GoTo 0

  LogMsg "INFO", "Markers removed=" & removed & " empty paragraphs removed=" & emptied
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

Sub PastePartFromHtmlDoc(w, htmlDoc, doc, partName)
  Dim insertStart, insertEnd, src, rng

  On Error Resume Next

  If Not doc.Bookmarks.Exists(partName) Then
    LogMsg "WARN", "Part " & partName & ": no bookmark in template — skip"
    Exit Sub
  End If
  Set src = PairRange(htmlDoc, partName)
  If src Is Nothing Then
    LogMsg "WARN", "Part " & partName & ": no bookmark pair " & partName & "/" & CloseName(partName) & " in opened HTML — skip"
    Exit Sub
  End If

  insertStart = doc.Bookmarks(partName).Range.Start
  insertEnd = doc.Bookmarks(partName).Range.End
  LogMsg "INFO", "Part " & partName & ": src " & src.Start & ".." & src.End & " -> target " & insertStart & ".." & insertEnd

  If src.End <= src.Start Then
    LogMsg "WARN", "Part " & partName & ": empty source range — skip"
    Exit Sub
  End If

  ' A zone ending with a paragraph mark keeps it when the fragment does not end
  ' with one; otherwise the fragment tail merges into the next template paragraph.
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

Function IsStyleBookmarkName(name)
  Dim base
  IsStyleBookmarkName = False
  If Len(name) < 8 Then Exit Function
  If LCase(Left(name, 6)) <> "style_" Then Exit Function
  base = Mid(name, 7)
  IsStyleBookmarkName = (InStrRev(base, "_") > 1)
End Function

Function ExtractStyleName(bmName)
  Dim base, pos
  base = Mid(bmName, Len("style_") + 1)
  pos = InStrRev(base, "_")
  If pos = 0 Then
    ExtractStyleName = ""
    Exit Function
  End If
  base = Left(base, pos - 1)
  base = Replace(base, "__", "_")
  base = Replace(base, "_s", " ")
  ExtractStyleName = base
End Function

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
  Dim st, styleName, rng, p

  On Error Resume Next

  Set rng = PairRange(doc, bmName)
  If rng Is Nothing Then
    LogMsg "WARN", "Style " & bmName & ": no bookmark pair " & bmName & "/" & CloseName(bmName) & " — skip"
    failed = failed + 1
    Exit Sub
  End If

  styleName = ExtractStyleName(bmName)
  Set st = TryGetStyle(doc, styleName)
  If st Is Nothing Then
    LogMsg "WARN", "Style " & bmName & ": no style [" & styleName & "] in document — skip"
    failed = failed + 1
    Exit Sub
  End If

  Select Case st.Type
    Case wdStyleTypeParagraph
      For Each p In rng.Paragraphs
        p.Range.Style = st
      Next
    Case wdStyleTypeCharacter
      rng.Style = st
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
      If IsStyleBookmarkName(bm.Name) Then
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

Sub InsertPartsOnly(w, htmlDoc, doc)
  Dim names(), i, bmName
  ReDim names(-1)
  CollectPartBookmarkNamesDesc doc, names
  If UBound(names) < 0 Then
    LogMsg "WARN", "No part_* bookmarks in template — nothing to insert"
    Exit Sub
  End If
  For i = 0 To UBound(names)
    bmName = names(i)
    PastePartFromHtmlDoc w, htmlDoc, doc, bmName
  Next
End Sub

'--- MAIN ---
Dim htmlPath, inDocx, outDocx, logPath, w, doc, htmlDoc, workDocx, workHtml

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

LogMsg "INFO", "InsertPartsOnly begin"
On Error Resume Next
InsertPartsOnly w, htmlDoc, doc
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

RemoveMarkers doc

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
