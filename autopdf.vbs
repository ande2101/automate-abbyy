set fso = createobject("scripting.filesystemobject")
Set objShell = WScript.CreateObject( "WScript.Shell" )

rootPath = "C:\Documents and Settings\anonymous\My Documents\PDF-maker"

while True

  wscript.echo "Waiting for start.txt"

  while fso.fileExists(rootPath & "\input\start.txt") = False
    wscript.sleep 10000
  wend

  set fin = fso.OpenTextFile(rootPath & "\input\start.txt")
  outputFileName = fin.readLine()
  outputfilename = Replace(outputFileName, ":", "-")
  outputfilename = Replace(outputFileName, "\", "-")
  outputfilename = Replace(outputFileName, "/", "-")
  outputfilename = Replace(outputFileName, ">", "_")
  outputfilename = Replace(outputFileName, "<", "_")
  outputfilename = Replace(outputFileName, "|", "_")
  outputfilename = Replace(outputFileName, "?", "_")
  outputfilename = Replace(outputFileName, "*", "_")
  outputfilename = Replace(outputFileName, chr(34), "_")
  outputfilename = RTrim(LTrim(outputFileName))
  fin.close()

  wscript.echo "Found job: " & outputFileName

  if fso.fileExists(rootPath & "\tmp\output.pdf") then
    wscript.echo "Temp output file already exists. Pausing until it is removed."
    while fso.fileExists(rootPath & "\tmp\output.pdf")
      wscript.sleep 10000
    wend
  end if

  if fso.fileExists(rootPath & "\output\" & outputFileName & ".pdf") then
    count = 1
    while fso.fileExists(rootPath & "\output\" & outputFileName & "(" & count & ").pdf")
      count = count + 1
    wend

    outputFileName = outputFileName & "(" & count & ")"

    wscript.echo "Output file already exists, renaming to " & outputfilename
  end if

  wscript.echo "Opening OCR software"

  wscript.sleep 5000

  objShell.Run("""C:\Program Files\ABBYY FineReader 11\FineCmd.exe""")

  wscript.sleep 5000

  wscript.echo "Starting ORC job"

  objShell.AppActivate("ABBYY FineReader 11 Corporate Edition")

  wscript.sleep 500
  objShell.SendKeys "%N"
  wscript.sleep 500
  objShell.SendKeys "{ESC}"
  wscript.sleep 500
  objShell.SendKeys "%FW"
  wscript.sleep 500
  objShell.SendKeys "{Tab}"
  wscript.sleep 100
  objShell.SendKeys "{Tab}"
  wscript.sleep 100
  objShell.SendKeys "{Tab}"
  wscript.sleep 100
  objShell.SendKeys "{Tab}"
  wscript.sleep 100
  objShell.SendKeys " "
  wscript.sleep 100

  wscript.echo "Waiting for output file"

  while fso.fileExists(rootPath & "\tmp\output.pdf") = False
    wscript.sleep 1000
  wend

  wscript.sleep 5000

  wscript.echo "Closing OCR software"


  objShell.AppActivate("Convert from PNGs")
  wscript.sleep 100
  objShell.SendKeys "%C"

  wscript.sleep 100
  objShell.SendKeys "%{F4}"
  wscript.sleep 100
  objShell.SendKeys "%N"


  wscript.echo "Moving output file"

  fso.MoveFile rootPath & "\tmp\output.pdf", rootPath & "\output\" & outputFileName & ".pdf"

  wscript.echo "Removing input files"

  fso.DeleteFile(rootPath & "\input\*.*")

wend

