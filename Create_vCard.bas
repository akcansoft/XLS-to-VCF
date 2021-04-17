Sub Create_vCard_File()
' Excel VBA Macro Code
' v1.0.191116
' Use this macro code in akcanSoft XLS to VCF.xlsm file ( https://github.com/akcansoft/XLS-to-VCF )

' 16/11/2019
' Mesut Akcan
' makcan@gmail.com
' akcansof.blogspot.com
' https://github.com/akcansoft/XLS-to-VCF
' youtube.com/mesutakcan

Dim fso, fs
Dim nRow As Long, lastRow As Long, txtFileName
Dim langNo As Integer, lData As Range
Dim r, tRows As Long, totalc As Long, cname As String, msg As String
Dim rec1 As Shape, rec2 As Shape, rec1w As Double
Dim rat As Double, n As Long, cell_v As String

With ActiveSheet
lastRow = WorksheetFunction.Max(.Cells(Rows.Count, 1).End(xlUp).Row, .Cells(Rows.Count, 2).End(xlUp).Row, .Cells(Rows.Count, 3).End(xlUp).Row)
End With
langNo = Range("lang_no") 'Language no
Set lData = Range("lang_data") 'language data table
tRows = lastRow - 2 'number of contacts
If tRows < 1 Then ' if no contact
    MsgBox lData.Cells(28, langNo) '28: No contact data found
    Exit Sub
End If

'29: vCard Files
'30: Please specify the file to save
'get text file name
txtFileName = Application.GetSaveAsFilename(, lData.Cells(29, langNo) & " (*.vcf), *.vcf", , lData.Cells(30, langNo))

If txtFileName = False Then Exit Sub

Set fso = CreateObject("Scripting.FileSystemObject") ' Create File System Object
If fso.FileExists(txtFileName) Then
    '31: The file already exists!
    '32: Do you want to overwrite it?
    'Message: The file already exists! Do you want to overwrite it?
    r = MsgBox(txtFileName & vbCrLf & lData.Cells(31, langNo) & vbCrLf & lData.Cells(32, langNo), vbYesNo)
    'if response No then exit
    If r = vbNo Then Exit Sub
End If
'Application.ScreenUpdating = False

'Progress bar
Set rec1 = ActiveSheet.Shapes("shp_rec1") 'rectangle shape 1 = grey
Set rec2 = ActiveSheet.Shapes("shp_rec2") 'rectangle shape 2 = green
rec2.Left = rec1.Left
rec2.Height = rec1.Height
rec2.Top = rec1.Top
rec2.Width = 0
rec1.Visible = True
rec2.Visible = True
rec2.TextFrame.Characters.Text = ""
rec1w = rec1.Width

Set fs = fso.CreateTextFile(txtFileName, True, False) 'Create Text File (filename, overwrite, unicode)

rat = 100 / tRows 'ratio

With ActiveSheet
For nRow = 3 To lastRow
    cname = Trim(.Range("C" & nRow)) & Trim(.Range("A" & nRow)) & Trim(.Range("B" & nRow))
    If cname <> "" Then
        fs.WriteLine "BEGIN:VCARD"
        fs.WriteLine "VERSION:3.0"
        cname = Trim(.Range("C" & nRow)) & ";" & Trim(.Range("A" & nRow)) & ";" & Trim(.Range("B" & nRow))
        
        fs.WriteLine "N:" & cname
        cell_v = .Range("D" & nRow)
        If cell_v <> "" Then fs.WriteLine "BDAY:" & Year(cell_v) & "-" & Month(cell_v) & "-" & Day(cell_v)
        cell_v = .Range("E" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=CELL:" & cell_v
        cell_v = .Range("F" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=CELL:" & cell_v
        cell_v = .Range("G" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=CELL:" & cell_v
        cell_v = .Range("H" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=HOME:" & cell_v
        cell_v = .Range("I" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=HOME:" & cell_v
        
        cell_v = .Range("J" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=WORK:" & cell_v
        cell_v = .Range("K" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=WORK:" & cell_v
        cell_v = .Range("L" & nRow): If cell_v <> "" Then fs.WriteLine "TEL;TYPE=FAX:" & cell_v
        cell_v = .Range("M" & nRow): If cell_v <> "" Then fs.WriteLine "EMAIL;TYPE=HOME;TYPE=INTERNET:" & cell_v
        cell_v = .Range("N" & nRow): If cell_v <> "" Then fs.WriteLine "EMAIL;TYPE=HOME;TYPE=INTERNET:" & cell_v
        cell_v = .Range("O" & nRow): If cell_v <> "" Then fs.WriteLine "EMAIL;TYPE=HOME;TYPE=INTERNET:" & cell_v
        cell_v = .Range("P" & nRow): If cell_v <> "" Then fs.WriteLine "EMAIL;TYPE=WORK;TYPE=INTERNET:" & cell_v
        cell_v = .Range("Q" & nRow): If cell_v <> "" Then fs.WriteLine "EMAIL;TYPE=WORK;TYPE=INTERNET:" & cell_v
        cell_v = .Range("R" & nRow): If cell_v <> "" Then fs.WriteLine "ADR;TYPE=HOME:" & cell_v
        cell_v = .Range("S" & nRow): If cell_v <> "" Then fs.WriteLine "ADR;TYPE=WORK:" & cell_v
        cell_v = .Range("T" & nRow): If cell_v <> "" Then fs.WriteLine "ORG:" & cell_v
        cell_v = .Range("U" & nRow): If cell_v <> "" Then fs.WriteLine "TITLE:" & cell_v
        cell_v = .Range("V" & nRow): If cell_v <> "" Then fs.WriteLine "URL:" & cell_v
        cell_v = .Range("W" & nRow): If cell_v <> "" Then fs.WriteLine "URL:" & cell_v
        cell_v = .Range("X" & nRow): If cell_v <> "" Then fs.WriteLine "CATEGORIES:" & cell_v
        cell_v = .Range("Y" & nRow): If cell_v <> "" Then fs.WriteLine "NOTE:" & cell_v
        
        fs.WriteLine "END:VCARD"
        fs.WriteLine
        totalc = totalc + 1 ' total contacts
    End If
    n = nRow - 2
    rec2.TextFrame.Characters.Text = Round(n * rat) & "%" 'write ratio
    rec2.Width = n * rec1w / tRows
    DoEvents
Next
End With
fs.Close
DoEvents
Beep
'Application.ScreenUpdating = True
msg = totalc & " " & lData.Cells(33, langNo) & vbCrLf & txtFileName '33: contacts are exported to
If tRows > totalc Then msg = msg & vbCrLf & tRows - totalc & " " & lData.Cells(34, langNo) '34: contacts not saved because they do not have name information
MsgBox msg

rec1.Visible = False
rec2.Visible = False
End Sub

Sub Language_Change()
    ActiveSheet.Buttons("button1").Text = Range("lang_data").Cells(27, Range("lang_no")) '27: Create vCard File (.vcf)
End Sub
