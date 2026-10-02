'===========================================================
' mdmeta2xml.vbs — YAML front matter from Markdown to XML
' Usage: cscript mdmeta2xml.vbs <input.md> <output.xml>
'===========================================================

Option Explicit

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

Function ReadTextFileUtf8(path)
  Dim stream
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.LoadFromFile path
  ReadTextFileUtf8 = stream.ReadText
  stream.Close
End Function

Sub WriteTextFileUtf8(path, text)
  Dim stream, fso, folder
  Set fso = CreateObject("Scripting.FileSystemObject")
  folder = fso.GetParentFolderName(path)
  If Len(folder) > 0 And Not fso.FolderExists(folder) Then
    fso.CreateFolder folder
  End If
  Set stream = CreateObject("ADODB.Stream")
  stream.Type = 2
  stream.Charset = "utf-8"
  stream.Open
  stream.WriteText text
  stream.SaveToFile path, 2
  stream.Close
End Sub

Function IsBlankLine(s)
  IsBlankLine = (Len(Trim(s)) = 0)
End Function

Function XmlEscape(s)
  Dim t
  t = s
  t = Replace(t, "&", "&amp;")
  t = Replace(t, "<", "&lt;")
  t = Replace(t, ">", "&gt;")
  t = Replace(t, """", "&quot;")
  XmlEscape = t
End Function

Function XmlAttrEscape(s)
  XmlAttrEscape = XmlEscape(s)
End Function

Function SplitLines(text)
  Dim normalized
  normalized = Replace(text, vbCrLf, vbLf)
  normalized = Replace(normalized, vbCr, vbLf)
  SplitLines = Split(normalized, vbLf)
End Function

Function ExtractFrontMatter(lines)
  Dim i, n, startIdx, inBlock, body(), count, line
  n = UBound(lines)
  startIdx = -1
  For i = 0 To n
    If Not IsBlankLine(lines(i)) Then
      startIdx = i
      Exit For
    End If
  Next
  If startIdx < 0 Then
    ExtractFrontMatter = ""
    Exit Function
  End If
  If Trim(lines(startIdx)) <> "---" Then
    ExtractFrontMatter = ""
    Exit Function
  End If
  count = 0
  ReDim body(-1)
  For i = startIdx + 1 To n
    line = lines(i)
    If Trim(line) = "---" Then
      If count = 0 Then
        ReDim body(-1)
      Else
        ReDim Preserve body(count - 1)
      End If
      ExtractFrontMatter = Join(body, vbLf)
      Exit Function
    End If
    ReDim Preserve body(count)
    body(count) = line
    count = count + 1
  Next
  ExtractFrontMatter = ""
End Function

Sub ParseMetaLine(line, nameOut, valueOut)
  Dim pos, key, val
  nameOut = ""
  valueOut = ""
  line = RTrim(line)
  If Len(line) = 0 Then Exit Sub
  pos = InStr(line, ":")
  If pos < 2 Then Exit Sub
  key = Trim(Left(line, pos - 1))
  val = Trim(Mid(line, pos + 1))
  If Len(key) = 0 Then Exit Sub
  nameOut = key
  valueOut = val
End Sub

Function BuildMdmetaXml(frontMatter)
  Dim lines, i, n, name, val, xml
  lines = SplitLines(frontMatter)
  n = UBound(lines)
  xml = "<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
  xml = xml & "<mdmeta>" & vbCrLf
  For i = 0 To n
    ParseMetaLine lines(i), name, val
    If Len(name) > 0 Then
      xml = xml & "   <meta name=""" & XmlAttrEscape(name) & """>" & XmlEscape(val) & "</meta>" & vbCrLf
    End If
  Next
  xml = xml & "</mdmeta>" & vbCrLf
  BuildMdmetaXml = xml
End Function

'--- MAIN ---
Dim inPath, outPath, text, fm, xmlOut

inPath = ResolvePath(GetArg(0, "input Markdown file"))
outPath = ResolvePath(GetArg(1, "output XML file"))

text = ReadTextFileUtf8(inPath)
fm = ExtractFrontMatter(SplitLines(text))
xmlOut = BuildMdmetaXml(fm)
WriteTextFileUtf8 outPath, xmlOut

WScript.Echo "Wrote: " & outPath
