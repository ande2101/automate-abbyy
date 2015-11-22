Function HTML_Decode(byVal encodedstring)
  Dim tmp, i
  tmp = encodedstring
  tmp = Replace(tmp, "&quot;", chr(34) )
  tmp = Replace(tmp, "&apos;", chr(39))
  tmp = Replace(tmp, "&lt;"  , chr(60) )
  tmp = Replace(tmp, "&gt;"  , chr(62) )
  tmp = Replace(tmp, "&amp;" , chr(38) )
  tmp = Replace(tmp, "&nbsp;", chr(32) )
  For i = 160 to 255
    tmp = Replace(tmp, "&#" & i & ";", chr(i))
  Next
  HTML_Decode = tmp
End Function



set fso = createobject("scripting.filesystemobject")
Set objShell = WScript.CreateObject( "WScript.Shell" )

rootPath  = "C:\Documents and Settings\anonymous\My Documents\PDF-maker"
inputPath = rootpath & "\input"
outputPath = rootpath & "\output"

sourcePath = "\\server\root\disk\input"
destPath   = "\\server\root\disk\output"

set sourcedir  = fso.getFolder(sourcePath)

while sourcedir.subfolders.count > 0 and fso.fileExists(rootpath & "\stop.txt") = False

  wscript.echo sourcedir.subfolders.count & " books to go..."

  if fso.getFolder(inputPath).files.Count > 0 then
    wscript.echo "Waiting for job to complete"
    
    while fso.getFolder(inputPath).files.Count > 0
      wscript.sleep 1000
    wend
  end if

  set CurrentBookDir = Nothing

  For Each bookdir In sourcedir.subfolders
    if fso.fileExists(bookdir.path & "\page0001.png") = True and fso.fileExists(bookdir.path & "\lock") = False Then
      Set currentBookDir = bookdir
      set fout = fso.createTextFile(currentbookdir.path & "\lock")
      fout.close()
      set fout = Nothing
      Exit For
    End If
  Next

  if CurrentBookDir is Nothing Then  
    wscript.echo "Run out of dirs to process"
    WScript.Quit 0
  End If


  wscript.echo "Moving images for " & currentbookdir.name

  fso.MoveFile currentbookdir.path & "\page*.png", inputPath



  wscript.echo "Getting name for " & currentbookdir.name

  set fin = fso.openTextFile(currentbookdir.path & "\details.xml")
  s = fin.readall()
  fin.close()

  begin = InStr(s, "<Title>") + 7
  title = Mid(s, begin, InStr(s, "</Title") -begin)

  filename = HTML_Decode(title)



  wscript.echo "Starting task for book: " & filename
  
  set fout = fso.createTextFile(inputPath & "\filename.txt")
  fout.write(filename)
  fout.close()

  fso.MoveFile inputPath & "\filename.txt", inputpath & "\start.txt"


  wscript.echo "Waiting for a new output file"

  startingcount = fso.getfolder(outputPath).files.count
  while startingcount = fso.getfolder(outputPath).files.count
    wscript.sleep 1000
  wend

  wscript.sleep 1000


  wscript.echo "Moving PDF to book dir"
  fso.MoveFile outputPath & "\*.pdf", currentbookdir.path


  wscript.echo "Moving finished book to output dir"
  currentbookdir.move destPath & "\" & currentbookdir.name

  set sourcedir  = fso.getFolder(sourcePath)

wend

