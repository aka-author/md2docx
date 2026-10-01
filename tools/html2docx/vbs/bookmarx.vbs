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

Function ReadTextFileUtf8(path)
  Dim stream
  LogMsg "INFO", "ReadTextFileUtf8: " & path
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile path
  ReadTextFileUtf8 = stream.ReadText
  stream.Close
End Function

Sub WriteTextFileUtf8(path, text)
  Dim stream
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText text
  stream.SaveToFile path, 2
  stream.Close
End Sub

Function IsPartBookmarkName(name)
  IsPartBookmarkName = (Len(name) >= 6 And LCase(Left(name, 5)) = "part_")
End Function

Function FindAnchorOpenTag(html, partName)
  Dim patterns(2), i, p, tagEnd
  patterns(0) = "<a name=""" & partName & """"
  patterns(1) = "<a name=""" & partName & " """
  patterns(2) = "<a name='" & partName & "'"
  For i = 0 To 2
    p = InStr(1, html, patterns(i), vbTextCompare)
    If p > 0 Then
      tagEnd = InStr(p, html, ">")
      If tagEnd = 0 Then Exit Function
      FindAnchorOpenTag = tagEnd + 1
      Exit Function
    End If
  Next
  FindAnchorOpenTag = 0
End Function

Function IsTagOpenA(html, pos)
  Dim ch
  If InStr(pos, html, "<a", vbTextCompare) <> pos Then
    IsTagOpenA = False
    Exit Function
  End If
  ch = LCase(Mid(html, pos + 2, 1))
  IsTagOpenA = (ch = " " Or ch = ">" Or ch = vbCr Or ch = vbLf Or ch = "/")
End Function

Function ExtractPartFragment(html, partName)
  Dim startContent, depth, i, openPos, closePos, closeTag
  startContent = FindAnchorOpenTag(html, partName)
  If startContent = 0 Then
    ExtractPartFragment = ""
    Exit Function
  End If
  depth = 1
  i = startContent
  closeTag = "</a>"
  Do While i <= Len(html) And depth > 0
    openPos = InStr(i, html, "<a", vbTextCompare)
    closePos = InStr(i, html, closeTag, vbTextCompare)
    If closePos = 0 And openPos = 0 Then Exit Do
    If closePos > 0 And (openPos = 0 Or closePos < openPos) Then
      depth = depth - 1
      If depth = 0 Then
        ExtractPartFragment = Mid(html, startContent, closePos - startContent)
        Exit Function
      End If
      i = closePos + Len(closeTag)
    Else
      If IsTagOpenA(html, openPos) Then depth = depth + 1
      i = openPos + 2
    End If
  Loop
  ExtractPartFragment = ""
End Function

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
  Dim insertStart, insertEnd, srcStart, srcEnd, rng

  On Error Resume Next

  If Not doc.Bookmarks.Exists(partName) Then
    LogMsg "WARN", "Part " & partName & ": no bookmark in template — skip"
    Exit Sub
  End If
  If Not htmlDoc.Bookmarks.Exists(partName) Then
    LogMsg "WARN", "Part " & partName & ": no bookmark in opened HTML — skip"
    Exit Sub
  End If

  insertStart = doc.Bookmarks(partName).Range.Start
  insertEnd = doc.Bookmarks(partName).Range.End

  srcStart = htmlDoc.Bookmarks(partName).Range.Start
  srcEnd = NextPartStart(htmlDoc, srcStart)
  LogMsg "INFO", "Part " & partName & ": src " & srcStart & ".." & srcEnd & " -> target " & insertStart & ".." & insertEnd

  If srcEnd <= srcStart Then
    LogMsg "WARN", "Part " & partName & ": empty source range — skip"
    Exit Sub
  End If

  Set rng = doc.Range(insertStart, insertEnd)
  rng.FormattedText = htmlDoc.Range(srcStart, srcEnd).FormattedText
  If Err.Number <> 0 Then
    LogMsg "ERROR", "Part " & partName & ": insert failed Err=" & Err.Number & " " & Err.Description
    Err.Clear
    Exit Sub
  End If

  LogMsg "INFO", "Part " & partName & ": insert OK"
End Sub

Function NextPartStart(htmlDoc, afterStart)
  Dim bm, best
  best = htmlDoc.Content.End
  For Each bm In htmlDoc.Bookmarks
    If IsPartBookmarkName(bm.Name) Then
      If bm.Range.Start > afterStart And bm.Range.Start < best Then
        best = bm.Range.Start
      End If
    End If
  Next
  NextPartStart = best
End Function

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
  Dim bm, st, styleName, rng, p

  On Error Resume Next

  Set bm = doc.Bookmarks(bmName)
  If Err.Number <> 0 Then
    LogMsg "WARN", "Style " & bmName & ": bookmark lost — " & Err.Description
    Err.Clear
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

  Set rng = bm.Range.Duplicate

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
