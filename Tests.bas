Attribute VB_Name = "Tests"
Private Sub NoUpdates()


Dim IE As Object
Dim nDate As String
Dim nFilm As String
Dim nType As String
Dim chkFilm As String
Dim getFilm As String
Dim drp As String
Dim car As String
Dim sls As String
Dim strPath As String
Dim wb As Workbook
Dim ws As Worksheet
Dim i As Long
Dim CIK As String
Dim chat As String


Dim x As Long
Dim z As Long
Dim y As Long

Dim kr As Long
Dim k As Long
Dim mb As Workbook, ms As Worksheet, gs As Worksheet
Dim nuname As String
Dim WinHttpReq As Object
Dim oStream As Object

Dim bb As Workbook, bs As Worksheet, br As Long
Dim nameRng As Name
Dim FSO As Object

Const sSortOrder As String = "Exit,New,Delta"


Set wb = Workbooks("CIK.xlsm")
Set ws = wb.Sheets("Manual")

'Find last row in CIK
lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

If Application.WorksheetFunction.CountIf(ws.Range("H2:H" & lr), 0) = lr - 1 Then
MsgBox "No updates"
Else
MsgBox "Updates"
End If


End Sub

Private Sub MakePivot()

Dim lr As Long
Dim i As Long
Dim wb As Workbook, ws As Worksheet
Dim car As String
Dim sls As String

Dim PB As Workbook, ps As Worksheet
Dim ub As Workbook, us As Worksheet
Dim p As Long, pr As Long, u As Long, ur As Long

Dim newb As Workbook


Const sSortOrder As String = "Exit,New,Delta"


Dim FindString As String
Dim Rng As Range

Dim drp As String, CIK As String
Dim bI As String
Dim nDate As String, nFilm As String, nType As String, chkFilm As String, getFilm As String

Dim strPath As String

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim pf As PivotField
Dim pf_Name As String

Application.ScreenUpdating = False

'Workbooks.Open ("C:\AM\AMDeltas.xlsm")
Set PB = Workbooks("AMDeltas.xlsm")
Set ps = PB.Worksheets("Filings")

'Find out how many rows we have in AMDeltas.xlsm
pr = ps.Cells(ps.Rows.Count, 1).End(xlUp).Row

'Add a new worksheet to hold the data

With PB
        Set sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        sht.Name = "PVTA"
End With

'Determine the data range to pivot
SrcData = ps.Name & "!" & Range("A1:S" & pr).Address(ReferenceStyle:=xlR1C1)

'Put where to start the pivot
  StartPvt = sht.Name & "!" & ps.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
Set pvtCache = PB.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="PivotTable1")

'Make pvt as the PivotTable1
Set pvt = PB.Worksheets("PVTA").PivotTables("PivotTable1")

'Add Status to the Report Filter
pvt.PivotFields("Status").Orientation = xlPageField

'Make a pivot field as Status so we can filter for New and Delta only, this means Excluding Exits
Set pf = pvt.PivotFields("Status")

'Enable filtering on multiple items
pf.EnableMultiplePageItems = True

'De-select Exits
pf.PivotItems("Exit").Visible = False
  
'Add field to the Column Labels
pvt.PivotFields("Year").Orientation = xlColumnField
pvt.PivotFields("Quarter").Orientation = xlColumnField
    
'Add fields to the Row Labels
pvt.PivotFields("Name of Issuer").Orientation = xlRowField
pvt.PivotFields("CUSIP").Orientation = xlRowField

'Make pf_Name to be the text "Count of Activists" which we need to pass into the Data field
pf_Name = "Count of Activist"

'Add the Data field
pvt.AddDataField pvt.PivotFields("Activist"), pf_Name, xlCount

'Add the Activist field
pvt.PivotFields("Activist").Orientation = xlRowField

'Turn off the Grand Totals
pvt.ColumnGrand = False
pvt.RowGrand = False

'That's the pivot table created.

'Now write the x-Activist formulae to the Filings Sheet.

ps.Range("T2:T" & pr).Value = "=IF(Q2=" & Chr(34) & "Exit" & Chr(34) & ", " & Chr(34) & Chr(34) & ", IF(GETPIVOTDATA(" & Chr(34) & "Count of Activist" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Name of Issuer" & Chr(34) & ", A2, " & Chr(34) & "CUSIP" & Chr(34) & ", C2, " & Chr(34) & "Year" & Chr(34) & ", R2 , " & Chr(34) & "Quarter" & Chr(34) & ", S2)=1, " & Chr(34) & Chr(34) & ", GETPIVOTDATA(" & Chr(34) & "Count of Activist" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Name of Issuer" & Chr(34) & ", A2, " & Chr(34) & "CUSIP" & Chr(34) & ", C2, " & Chr(34) & "Year" & Chr(34) & ", R2, " & Chr(34) & "Quarter" & Chr(34) & ", S2)))"
Application.Wait (Now + TimeValue("0:00:10"))
ps.Range("T2:T" & pr).Value = ps.Range("T2:T" & pr).Value

'Now write the x-Activist ISSUER formuale to the Filings Sheet

'Seems we need to pause Excel here to let it catch up
Application.Wait (Now + TimeValue("0:00:10"))

ps.Range("V2:V" & pr).Value = "=IF(Q2=" & Chr(34) & "Exit" & Chr(34) & ", " & Chr(34) & Chr(34) & ", IF(GETPIVOTDATA(" & Chr(34) & "Count of Activist" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Name of Issuer" & Chr(34) & ", A2, " & Chr(34) & "Year" & Chr(34) & ", R2, " & Chr(34) & "Quarter" & Chr(34) & ", S2) = 1, " & Chr(34) & Chr(34) & ", GETPIVOTDATA(" & Chr(34) & "Count of Activist" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Name of Issuer" & Chr(34) & ", A2, " & Chr(34) & "Year" & Chr(34) & ", R2, " & Chr(34) & "Quarter" & Chr(34) & ", S2)))"
Application.Wait (Now + TimeValue("0:00:10"))
ps.Range("V2:V" & pr).Value = ps.Range("V2:V" & pr).Value

'Now need to rejig the table to do the m-Pos

'Remove Row Labels
pvt.PivotFields("CUSIP").Orientation = xlHidden
pvt.PivotFields("Activist").Orientation = xlHidden
pvt.PivotFields("Name of Issuer").Orientation = xlHidden
    
'Removing Values
pvt.PivotFields("Count of Activist").Orientation = xlHidden

'Add fields to the Row Labels
pvt.PivotFields("Activist").Orientation = xlRowField

'Make pf_Name to be the text "Count of NAME OF ISSUER"
pf_Name = "Count of NAME OF ISSUER"

'Add the Data field
pvt.AddDataField pvt.PivotFields("NAME OF ISSUER"), pf_Name, xlCount

'Add the Activist field
pvt.PivotFields("NAME OF ISSUER").Orientation = xlRowField

'Turn off the Grand Totals
pvt.ColumnGrand = False
pvt.RowGrand = False

'Refresh the pivot cache
pvt.PivotCache.Refresh

'Write in the formula to get the M-Pos

ps.Range("U2:U" & pr).Value = "=IF(Q2=" & Chr(34) & "Exit" & Chr(34) & ", " & Chr(34) & Chr(34) & ", IF(GETPIVOTDATA(" & Chr(34) & "Count of Name of Issuer" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Activist" & Chr(34) & ", P2, " & Chr(34) & "NAME OF ISSUER" & Chr(34) & ", A2," & Chr(34) & "Year" & Chr(34) & ", R2, " & Chr(34) & "Quarter" & Chr(34) & ", S2) = 1, " & Chr(34) & Chr(34) & ", GETPIVOTDATA(" & Chr(34) & "Count of Name of Issuer" & Chr(34) & ", PVTA!$A$3, " & Chr(34) & "Activist" & Chr(34) & ", P2, " & Chr(34) & "NAME OF ISSUER" & Chr(34) & ", A2," & Chr(34) & "Year" & Chr(34) & ", R2, " & Chr(34) & "Quarter" & Chr(34) & ", S2)))"
Application.Wait (Now + TimeValue("0:00:10"))
ps.Range("U2:U" & pr).Value = ps.Range("U2:U" & pr).Value

'Delete the Pivot table
Application.DisplayAlerts = False
PB.Sheets(2).Delete

'Warnings back on
Application.DisplayAlerts = True

'All done
Application.ScreenUpdating = True

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub


Private Sub SmartCIK()


Dim IE As Object
Dim nDate As String
Dim nFilm As String
Dim nType As String
Dim chkFilm As String
Dim getFilm As String
Dim drp As String
Dim car As String
Dim sls As String
Dim strPath As String
Dim wb As Workbook
Dim ws As Worksheet
Dim i As Long
Dim CIK As String
Dim chat As String


Dim x As Long
Dim z As Long
Dim y As Long

Dim kr As Long
Dim k As Long
Dim mb As Workbook, ms As Worksheet, gs As Worksheet
Dim nuname As String
Dim WinHttpReq As Object
Dim oStream As Object

Dim bb As Workbook, bs As Worksheet, br As Long
Dim nameRng As Name
Dim FSO As Object

Const sSortOrder As String = "Exit,New,Delta"

Dim wbks As Long
'Dim z As Long
Dim DText As String

Set wb = Workbooks("CIK.xlsm")
Set ws = wb.Sheets("Activists")

'Find last row in CIK
lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

i = 2


'Run checks that CIK is a valid number, Master File exists, Folder path is valid, and 1P Film number is a valid number
' Check CIK is valid
' the Asci chars 49 to 57 represent the numbers 0 to 9 so this checks that each char in the CIK string is one of those numbers.




'End of Initialisation
Application.ScreenUpdating = False

'Back up CIK files

sls = Format(Now, "yyyy-mm-dd hh-mm-ss")
strPath = "C:\AM\zz_LiveBackUp\" & sls & "_CIK.xlsm"

Set bb = Workbooks.Add

    Set wb = Workbooks("CIK.xlsm")
    Set ws = wb.Sheets("Activists")
  

Application.DisplayAlerts = False

ws.Copy After:=bb.Sheets(1)
bb.Sheets(1).Delete
wb.Sheets(2).Copy After:=bb.Sheets(1)
bb.SaveAs Filename:=strPath, _
            FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
            Password:="", _
            WriteResPassword:="", _
            ReadOnlyRecommended:=False, _
            CreateBackup:=False
bb.Close

Application.DisplayAlerts = True

'Back up CUSIPS files


'Make ms the CUSIPSfilesworksheet
Set ms = wb.Sheets(2)

'Find last row on CUSIPfiles worksheet
z = ms.Cells(ms.Rows.Count, 1).End(xlUp).Row
Debug.Print z

'''
'''For y = 2 To z
'''
''''Copy each CUSIP file whose filename is in column A of CUSIPfiles worksheet
'''strPath = "C:\AM\zz_LiveBackUp\" & sls & ms.Range("A" & y).Value
'''FileCopy ms.Range("A" & y).Value, strPath
'''strPath = ""
'''
'''
''''Now need to open each one and update the Named Ranges for each
''''Each has the Ticker, CUSIP and Shares Outstanding data. When we process the filing we write in a lookup formula to get the shares outstanding
''''from the the relevant CUSIPS file.
'''
''''Make strPath the Filepath as in Column B
'''strPath = ms.Range("B" & y).Value
'''Workbooks.Open (strPath)
'''
''''Make strPath the name of the workbook as in Column A
'''strPath = ms.Range("A" & y).Value
'''
''''Check CUSIPS for formulae, trim data etc, format as numbers
'''Set bb = Workbooks(strPath)
'''Set bs = bb.Sheets("TickersShares")
'''bI = bs.Cells(bs.Rows.Count, 1).End(xlUp).Row
'''
''''Formula for Unique Identifier
'''bs.Range("F2:F" & bI).Value = "=IF(ISBLANK(C2)," & Chr(34) & Chr(34) & ", TRIM(C2) & " & Chr(34) & ": COM" & Chr(34) & ")"
'''
''''Set Shares Outstanding to be number format
'''bs.Range("D2:D" & bI).NumberFormat = "0"
'''
''''Update the Include and Exclude ranges, ShareTbls and Uniq ranges in CUSIPS. These allow us to do lookups on Shares Outstanding, CUSIPS and what is and isn't defined as Com = Y
'''Set bs = bb.Sheets("UniqueClass")
'''
''''Set Stringpath for Include column last row, col F, which is 6th Column to find out how many rows are in it
'''bI = bs.Cells(bs.Rows.Count, 6).End(xlUp).Row
'''
''''Set Excluded as named range
'''Set nameRng = bb.Names.Item("Included")
'''nameRng.RefersTo = "=UniqueClass!$F$2:$F$" & bI
'''
''''Set Stringpath for Exclude column last row, col H, which is 8th Column to find out how many rows are in it
'''bI = bs.Cells(bs.Rows.Count, 8).End(xlUp).Row
'''
''''Set Excluded as named range
'''Set nameRng = bb.Names.Item("Excluded")
'''nameRng.RefersTo = "=UniqueClass!$H$2:$H$" & bI
'''
''''Set Stringpath for Name of Issuer column on sheet TickersShares last row, Column A
'''Set bs = bb.Sheets("TickersShares")
'''bI = bs.Cells(bs.Rows.Count, 1).End(xlUp).Row
'''
''''Set ShareTbls as named range
'''Set nameRng = bb.Names.Item("ShareTbl")
'''nameRng.RefersTo = "=TickersShares!$A$2:$F$" & bI
'''
''''Set Uniq as Name range
'''Set nameRng = bb.Names.Item("Uniq")
'''nameRng.RefersTo = "=TickersShares!$F$2:$F$" & bI
'''
''''Save CUSIPS
'''bb.Save
'''bb.Close
'''
'''Set bb = Nothing
'''bI = ""
'''
'''Next y

'NWC - need to figure out which CUSIPS file to open
'If filing is a 13F-HR we open the one based on the Filing Date in K, otherwise we have to open the Master file and find the filing date of the previous HR


Dim abI As String
Dim aCIK As String
Dim aStrPath As String
Dim awbks As Long
Dim abb As Workbook
Dim aDText As String
Dim az As Long

If ws.Range("M" & i).Value = "13F-HR" Then

'Set J2 on CUSIPSfile to be the 1P Filing Date, then J3 will generate the quarter, and J4 the year. Pass these into bI and CIK, and use strPath to have the correct CUSIP filename
wb.Sheets(2).Range("J2").NumberFormat = "mm/dd/yyyy"
wb.Sheets(2).Range("J2").Value = ws.Range("K" & i).Value
wb.Sheets(2).Range("J3").Value = "=IF(ROUNDUP(MONTH(J2)/3, 0) = 1," & Chr(34) & "Q4" & Chr(34) & ", " & Chr(34) & "Q" & Chr(34) & " & ROUNDUP(MONTH(J2)/3, 0)-1)"
wb.Sheets(2).Range("J4").Value = "=IF(J3 = " & Chr(34) & "Q4" & Chr(34) & ", YEAR(J2)-1, YEAR(J2))"
abI = wb.Sheets(2).Range("J3").Value 'Q
aCIK = wb.Sheets(2).Range("J4").Value 'Y
abI = abI & aCIK
aStrPath = "C:\AM\" & abI & "CUSIPS.xlsx"
'Clear the calcs
wb.Sheets(2).Range("J2:J4").Clear

'Open the CUSIPS file
Workbooks.Open (aStrPath)

'Set its filename into aStrPath
aStrPath = abI & "CUSIPS.xlsx"
Set abb = Workbooks(aStrPath)

Else ' It's an HRA, need to open the Master and go through to find the previous HRA

aStrPath = ws.Range("G" & i).Value & ws.Range("F" & i).Value
Workbooks.Open (aStrPath)

'Set strPath as the filename, and set the master into BB

aStrPath = ws.Range("F" & i).Value
Set abb = Workbooks(aStrPath)

awbks = abb.Worksheets.Count

For az = 3 To awbks
If abb.Worksheets(az).Name Like "*HRA*" Then
'Do nothing, it's an HRA, go to the next

Else

'This is the previous 13F-HR, get the last 8 chars and manipulate it to make a date field in format mm/dd/yyyy
aDText = Mid(Right(abb.Worksheets(az).Name, 8), 5, 2) & "/"
aDText = aDText & Right(abb.Worksheets(az).Name, 2) & "/"
aDText = aDText & Left(Right(abb.Worksheets(az).Name, 8), 4)
abb.Close False
Set abb = Nothing

Exit For


End If

Next az

'Put DText into a field, then use the Q and Y calcs to generate the Q and Y.

'Set J2 on CUSIPSfile to be the 1P Filing Date, then J3 will generate the quarter, and J4 the year. Pass these into bI and CIK, and use strPath to have the correct CUSIP filename
wb.Sheets(2).Range("J2").NumberFormat = "mm/dd/yyyy"
wb.Sheets(2).Range("J2").Value = aDText
wb.Sheets(2).Range("J3").Value = "=IF(ROUNDUP(MONTH(J2)/3, 0) = 1," & Chr(34) & "Q4" & Chr(34) & ", " & Chr(34) & "Q" & Chr(34) & " & ROUNDUP(MONTH(J2)/3, 0)-1)"
wb.Sheets(2).Range("J4").Value = "=IF(J3 = " & Chr(34) & "Q4" & Chr(34) & ", YEAR(J2)-1, YEAR(J2))"
abI = wb.Sheets(2).Range("J3").Value 'Q
aCIK = wb.Sheets(2).Range("J4").Value 'Y
abI = abI & aCIK
aStrPath = "C:\AM\" & abI & "CUSIPS.xlsx"
'Clear the calcs
wb.Sheets(2).Range("J2:J4").Clear

'Open the CUSIPS file
Workbooks.Open (aStrPath)

'Need to pass its name into a String.
aStrPath = abI & "CUSIPS.xlsx"
Set abb = Workbooks(aStrPath)


End If


'Open Tester

Workbooks.Open ("C:\AM\Blue Harbour Group 0001325256\Tester.xlsx")
Set ms = Workbooks("Tester.xlsx").Sheets(1)
'calculate kr
kr = ms.Cells(ms.Rows.Count, 1).End(xlUp).Row

'put formulae in
'Col H unique entity, Trim CUSIP and add :COM on the end.
'=TRIM(C2) & ": COM"
car = "=TRIM(C2) & " & Chr(34) & ": COM" & Chr(34)

With ms.Range("H2:H" & kr)
.Formula = car
.Value = .Value
End With

car = ""

'CUSIPS.xlsx
' " & aStrPath & "
'Put in formula for Shrs Out in Col L. 'This formula looks at the file holding all the Ticker data and Shares Outstanding, located 'C:\AM\CUSIPS.xlsx'.
'That file has named ranges, Uniq - a column with the same calculation as Unique Identifier to match the Issuer.
'Range ShareTbl which is all the Ticker info, column 4 has the Shares Outstanding value.
car = "=IF(ISNA(MATCH(H2, " & aStrPath & "!Uniq,0))," & Chr(34) & "No Shrs" & Chr(34) & ", IF(INDEX(" & aStrPath & "!ShareTbl, MATCH(H2," & aStrPath & "!Uniq,0),4)=" & Chr(34) & Chr(34) & "," & Chr(34) & "No Value" & Chr(34) & ",INDEX(" & aStrPath & "!ShareTbl,MATCH(H2," & aStrPath & "!Uniq,0),4)))"
With ms.Range("L2:L" & kr)
.Formula = car
.Value = .Value
End With
car = ""

'need to put in Col N a value Y to identify what is considered Common stock or instruments we are targeting.
'there are 2 lists in " & aStrPath & ", Included and Excluded. The Title of Class must contain text belonging to Included, but not on Excluded.
'=IF(AND(SUMPRODUCT(--ISNUMBER(SEARCH(" & aStrPath & "!Excluded,B2)))<1, SUMPRODUCT(--ISNUMBER(SEARCH(" & aStrPath & "!Included,B2)))>0), "Y", "")
car = "=IF(OR(ISNUMBER(L2), AND(SUMPRODUCT(--ISNUMBER(SEARCH(" & aStrPath & "!Excluded,B2)))<1, SUMPRODUCT(--ISNUMBER(SEARCH(" & aStrPath & "!Included,B2)))>0)), " & Chr(34) & "Y" & Chr(34) & ", " & Chr(34) & Chr(34) & ")"
With ms.Range("N2:N" & kr)
.Formula = car
.Value = .Value
End With

car = ""


'Put in correct formula for Total Shares in Col I
'This sums the values in Column E (SHRS OR PRN AMT) for matching CUSIPS, excluding Puts and Calls
car = "=SUMIFS($E$2:$E$" & kr & ", $H$2:$H$" & kr & ", H2, $G$2:$G$" & kr & ", " & Chr(34) & "<>Put" & Chr(34) & ", $G$2:$G$" & kr & ", " & Chr(34) & "<>Call" & Chr(34) & ")"
With ms.Range("I2:I" & kr)
.Formula = car
.Value = .Value
End With
car = ""

'Col J Total Puts, this sums the values of SHRS OR PRN AMT for matching CUSIPS which have Put in Column G
car = "=IF(SUMIFS($E$2:$E$" & kr & ",$H$2:$H$" & kr & ", H2, $G$2:$G$" & kr & ", " & Chr(34) & "PUT" & Chr(34) & ")=0, " & Chr(34) & Chr(34) & ",SUMIFS($E$2:$E$" & kr & ",$H$2:$H$" & kr & ", H2, $G$2:$G$" & kr & ", " & Chr(34) & "PUT" & Chr(34) & "))"
With ms.Range("J2:J" & kr)
.Formula = car
End With

car = ""

'Col K Final Total =IF(J2 = "", I2, I2-J2)
car = "=IF(J2=" & Chr(34) & Chr(34) & ", I2, I2-J2)"
With ms.Range("K2:K" & kr)
.Formula = car
End With
car = ""


'Col M %Stake; =IF(OR(L2="No Shrs", L2 = "No Value"),"", ROUND((I2/L2)*100, 2))
car = "=IF(OR(L2=" & Chr(34) & "No Shrs" & Chr(34) & ", L2 = " & Chr(34) & "No Value" & Chr(34) & ")," & Chr(34) & Chr(34) & ", ROUND((I2/L2)*100, 2))"
With ms.Range("M2:M" & kr)
.Formula = car
.Value = .Value
End With
car = ""

'Need to close the CUSIP file
Workbooks(aStrPath).Close False

'Now format the Final Total to show negatives in red
ms.Range("K2:K" & kr).NumberFormat = "#,##0_ ;[Red]-#,##0"

'Now format shares columns as numbers with separator
ms.Range("I2:J" & kr).NumberFormat = "#,##0_ ;[Red]-#,##0"
ms.Range("L2:L" & kr).NumberFormat = "#,##0_ ;[Red]-#,##0"


'close it
Workbooks("Tester.xlsx").Close False


'FOR Manual N1

Dim mabi As String
Dim maCIK As String
Dim maStrPath As String
Dim mawbks As Long
Dim mabb As Workbook
Dim maDText As String
Dim maz As Long

Dim jr As Long
Dim ns As Worksheet


If ws.Range("M" & i).Value = "13F-HR" Then

'Set J2 on CUSIPSfile to be the 1P Filing Date, then J3 will generate the quarter, and J4 the year. Pass these into bI and CIK, and use strPath to have the correct CUSIP filename
wb.Sheets(2).Range("J2").NumberFormat = "mm/dd/yyyy"
wb.Sheets(2).Range("J2").Value = ws.Range("K" & i).Value
wb.Sheets(2).Range("J3").Value = "=IF(ROUNDUP(MONTH(J2)/3, 0) = 1," & Chr(34) & "Q4" & Chr(34) & ", " & Chr(34) & "Q" & Chr(34) & " & ROUNDUP(MONTH(J2)/3, 0)-1)"
wb.Sheets(2).Range("J4").Value = "=IF(J3 = " & Chr(34) & "Q4" & Chr(34) & ", YEAR(J2)-1, YEAR(J2))"
mabi = wb.Sheets(2).Range("J3").Value 'Q
maCIK = wb.Sheets(2).Range("J4").Value 'Y
mabi = mabi & maCIK
maStrPath = "C:\AM\" & mabi & "CUSIPS.xlsx"
'Clear the calcs
wb.Sheets(2).Range("J2:J4").Clear

'Open the CUSIPS file
Workbooks.Open (maStrPath)

'Set its filename into maStrPath
maStrPath = mabi & "CUSIPS.xlsx"
Set mabb = Workbooks(maStrPath)

Else ' It's an HRA, need to open the Master and go through to find the previous HRA

maStrPath = ws.Range("G" & i).Value & ws.Range("F" & i).Value
Workbooks.Open (maStrPath)

'Set strPath as the filename, and set the master into BB

maStrPath = ws.Range("F" & i).Value
Set mabb = Workbooks(maStrPath)

mawbks = mabb.Worksheets.Count

For maz = 3 To mawbks
If mabb.Worksheets(maz).Name Like "*HRA*" Then
'Do nothing, it's an HRA, go to the next

Else

'This is the previous 13F-HR, get the last 8 chars and manipulate it to make a date field in format mm/dd/yyyy
maDText = Mid(Right(mabb.Worksheets(maz).Name, 8), 5, 2) & "/"
maDText = maDText & Right(mabb.Worksheets(maz).Name, 2) & "/"
maDText = maDText & Left(Right(mabb.Worksheets(maz).Name, 8), 4)
mabb.Close False
Set mabb = Nothing

Exit For


End If

Next maz

'Put DText into a field, then use the Q and Y calcs to generate the Q and Y.

'Set J2 on CUSIPSfile to be the 1P Filing Date, then J3 will generate the quarter, and J4 the year. Pass these into bI and CIK, and use strPath to have the correct CUSIP filename
wb.Sheets(2).Range("J2").NumberFormat = "mm/dd/yyyy"
wb.Sheets(2).Range("J2").Value = maDText
wb.Sheets(2).Range("J3").Value = "=IF(ROUNDUP(MONTH(J2)/3, 0) = 1," & Chr(34) & "Q4" & Chr(34) & ", " & Chr(34) & "Q" & Chr(34) & " & ROUNDUP(MONTH(J2)/3, 0)-1)"
wb.Sheets(2).Range("J4").Value = "=IF(J3 = " & Chr(34) & "Q4" & Chr(34) & ", YEAR(J2)-1, YEAR(J2))"
mabi = wb.Sheets(2).Range("J3").Value 'Q
maCIK = wb.Sheets(2).Range("J4").Value 'Y
mabi = mabi & maCIK
maStrPath = "C:\AM\" & mabi & "CUSIPS.xlsx"
'Clear the calcs
wb.Sheets(2).Range("J2:J4").Clear

'Open the CUSIPS file
Workbooks.Open (maStrPath)

'Need to pass its name into a String.
maStrPath = mabi & "CUSIPS.xlsx"
Set mabb = Workbooks(maStrPath)


End If





Workbooks.Open ("C:\AM\Blue Harbour Group 0001325256\Tester.xlsx")
Set ns = Workbooks("Tester.xlsx").Sheets(1)

'recalculate kr
jr = ns.Cells(ns.Rows.Count, 1).End(xlUp).Row


'Col H unique entity, concetanate the 6 left chars of cusip as that is unique to an instrument, and if the Title of Class contains COM or SHS, then we'll add Com, otherwise include Instrument.
'Write the formula to CAR.
car = "=TRIM(C2) & " & Chr(34) & ": COM" & Chr(34)
With ns.Range("H2:H" & jr)
.Formula = car
.Value = .Value
End With

car = ""


'Put in formula for Shrs Out in Col L. 'This formula looks at the file holding all the Ticker data and Shares Outstanding, located 'C:\AM\" & maStrPath & "'.
'That file has named ranges, Uniq - a column with the same calculation as Unique Identifier to match the Issuer.
'Range ShareTbl which is all the Ticker info, column 4 has the Shares Outstanding value.
car = "=IF(ISNA(MATCH(H2, " & maStrPath & "!Uniq,0))," & Chr(34) & "No Shrs" & Chr(34) & ", IF(INDEX(" & maStrPath & "!ShareTbl, MATCH(H2," & maStrPath & "!Uniq,0),4)=" & Chr(34) & Chr(34) & "," & Chr(34) & "No Value" & Chr(34) & ",INDEX(" & maStrPath & "!ShareTbl,MATCH(H2," & maStrPath & "!Uniq,0),4)))"
With ns.Range("L2:L" & jr)
.Formula = car
.Value = .Value
End With
car = ""

'need to put in Col N a value Y for if any of Title of Class contain SHS or COM, this allows only Shares to be identified.
car = "=IF(OR(ISNUMBER(L2), AND(SUMPRODUCT(--ISNUMBER(SEARCH(" & maStrPath & "!Excluded,B2)))<1, SUMPRODUCT(--ISNUMBER(SEARCH(" & maStrPath & "!Included,B2)))>0)), " & Chr(34) & "Y" & Chr(34) & ", " & Chr(34) & Chr(34) & ")"
With ns.Range("N2:N" & jr)
.Formula = car
.Value = .Value
End With

car = ""



'Put in correct formula for Total Shares in Col I
car = "=SUMIFS($E$2:$E$" & jr & ", $H$2:$H$" & jr & ", H2, $G$2:$G$" & jr & ", " & Chr(34) & "<>Put" & Chr(34) & ", $G$2:$G$" & jr & ", " & Chr(34) & "<>Call" & Chr(34) & ")"
With ns.Range("I2:I" & jr)
.Formula = car
.Value = .Value
End With
car = ""


'Col J Total Puts
car = "=IF(SUMIFS($E$2:$E$" & jr & ",$H$2:$H$" & jr & ", H2, $G$2:$G$" & jr & ", " & Chr(34) & "PUT" & Chr(34) & ")=0, " & Chr(34) & Chr(34) & ",SUMIFS($E$2:$E$" & jr & ",$H$2:$H$" & jr & ", H2, $G$2:$G$" & jr & ", " & Chr(34) & "PUT" & Chr(34) & "))"
With ns.Range("J2:J" & jr)
.Formula = car
.Value = .Value
End With

car = ""

'Col K Final Total
car = "=IF(J2=" & Chr(34) & Chr(34) & ", I2, I2-J2)"
With ns.Range("K2:K" & jr)
.Formula = car
.Value = .Value
End With
car = ""


'Col M %Stake
car = "=IF(OR(L2=" & Chr(34) & "No Shrs" & Chr(34) & ", L2 = " & Chr(34) & "No Value" & Chr(34) & ")," & Chr(34) & Chr(34) & ", ROUND((I2/L2)*100, 2))"
With ns.Range("M2:M" & jr)
.Formula = car
.Value = .Value
End With
car = ""

'Need to close the CUSIP file
Workbooks(maStrPath).Close False

'Now format the Final Total to show negatives in red
ns.Range("K2:K" & jr).NumberFormat = "#,##0_ ;[Red]-#,##0"

'Now format shares columns as numbers with separator
ns.Range("I2:J" & jr).NumberFormat = "#,##0_ ;[Red]-#,##0"
ns.Range("L2:L" & jr).NumberFormat = "#,##0_ ;[Red]-#,##0"




End Sub


Private Sub FnSort()


Dim lr As Long
Dim i As Long
Dim wb As Workbook, ws As Worksheet
Dim car As String
Dim sls As String

Dim PB As Workbook, ps As Worksheet
Dim ub As Workbook, us As Worksheet
Dim p As Long, pr As Long, u As Long, ur As Long

Dim newb As Workbook


Const sSortOrder As String = "Exit,New,Delta"

Set PB = Workbooks("AMDeltas.xlsm")
Set ps = PB.Worksheets("Filings")
ps.DisplayPageBreaks = False

'Find out how many rows we have in AMDeltas.xlsm
pr = ps.Cells(ps.Rows.Count, 1).End(xlUp).Row

'Sort descending by Year (R) Quarter (S) descending by Filing Date (O) ascending by Activist (P) custom sort using sSortOrder for Status (Q) and descending for Change (M) and ascending for Issuer (A)
ps.Sort.SortFields.Clear


ps.Sort.SortFields.Add Key:=ps.Range("R1" & pr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("S1" & pr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("O1" & pr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("P1" & pr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("Q1" & pr), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:=sSortOrder, _
DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("M1" & pr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
ps.Sort.SortFields.Add Key:=ps.Range("A1" & pr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal


With ps.Sort
.SetRange ps.Range("A1:X" & pr)
.Header = xlYes
.MatchCase = False
.Orientation = xlTopToBottom
.SortMethod = xlPinYin
.Apply
End With
ps.Sort.SortFields.Clear



End Sub

Function ColLetter(z)
    ColLetter = Split(Cells(1, z).Address, "$")(1)
End Function

Private Sub UsedRange_Example_Column()
    Dim LastColumn As Long
    
Dim drp As String
Dim car As String
Dim sls As String
Dim strPath As String
Dim wb As Workbook
Dim ws As Worksheet
Dim i As Long
Dim CIK As String
Dim chat As String


Dim x As Long
Dim z As Long
Dim y As Long

Dim kr As Long
Dim k As Long


Set wb = Workbooks("CIK.xlsm")
Set ws = wb.Sheets("Activists")

'Find last row in CIK
lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


'Get last column which has text in it this will give us the extent of the data captured.
y = ws.Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious).Column

'Get last column number in Header row which has text in it
x = ws.Range("A1").End(xlToRight).Column

'find out if we need to write anymore column headers by subtracting x from y.
If x = y Then
'We don't, they are the sa,e

Else

'Pass y into z, we'll need it to give us the number of times we need to loop through and add headers.
z = y

'Subtract x from z to tell us how many new columns of data we have
'Divide z by 4 as we want to know how many times we copy in the range of 4 cells to make the new header
z = z - x
z = z / 4

'Now loop through, copy the last 4 columns headers to the next 4 cells. Then recalculate x each time - the last column header.

For kr = 1 To z

ws.Range(Cells(1, x - 3), Cells(1, x)).Copy ws.Cells(1, x + 1)

'Recalculate x
x = ws.Range("A1").End(xlToRight).Column
'Debug.Print x

Next kr

'All column headers are done now
End If
    
End Sub


Private Sub Frzr()

Dim mb As Workbook, ms As Worksheet, gs As Worksheet


Set mb = Workbooks("Tester.xlsx")
Set ms = mb.Sheets(1)
Set gs = mb.Sheets(2)


'With ActiveWindow
'
'    .SplitColumn = 0
'    .SplitRow = 1
'    .FreezePanes = True
'End With

With ms

With Application.Windows(mb.Name)
Application.Goto ms.Range("A2")
'    .SplitColumn = 0
'    .SplitRow = 1
    .FreezePanes = True
End With

End With

With gs

With Application.Windows(mb.Name)
Application.Goto gs.Range("A2")
'    .SplitColumn = 0
'    .SplitRow = 1
    .FreezePanes = True
End With

End With


'Application.Goto ms.Range("A2")
'ActiveWindow.FreezePanes = True
'Application.Goto ms.Range("A1")
'
'Application.Goto gs.Range("A2")
'ActiveWindow.FreezePanes = True
'Application.Goto gs.Range("A1")



End Sub

Private Sub xr()
'Dim Refo As Reference
Dim wb As Workbook
Set wb = Workbooks("CIK.xlsm")

For Each Reference In wb.VBProject.References
    Debug.Print Reference.Description; " -- "; Reference.FullPath
Next

End Sub


Private Sub Savr()

Dim wb As Workbook
Dim numame As String
Dim trModuleName As String

Dim strFolder As String
Dim strTempFile As String

Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
    
       


Set wb = Workbooks("CIK.xlsm")
Set nb = Workbooks.Open("C:\AM\ActivistCollector.xlsm")

'Remove the ActCollect module from ActivistCollector.xlsm
Set VBProj = nb.VBProject
Set VBComp = VBProj.VBComponents("ActCollect")
VBProj.VBComponents.Remove VBComp


strFolder = wb.Path

If Len(strFolder) = 0 Then strFolder = CurDir
strFolder = strFolder & "\"
strTempFile = strFolder & "~tmpexport.bas"
On Error Resume Next
strModuleName = "ActCollect"
wb.VBProject.VBComponents(strModuleName).Export strTempFile
nb.VBProject.VBComponents.Import strTempFile
Kill strTempFile
    On Error GoTo 0

nb.Close True
wb.Close True

End Sub

Private Sub CusipHR()


Dim wb As Workbook
Dim ws As Worksheet

Dim i As String

Dim mabi As String
Dim maCIK As String
Dim maStrPath As String
Dim mabb As Workbook

Set wb = Workbooks("CIK.xlsm")
Set ws = wb.Sheets(1)

i = 2

'"=IF(ROUNDUP(MONTH(J2)/3, 0) = 1," & Chr(34) & "Q4" & Chr(34) & ", " & Chr(34) & "Q" & Chr(34) & " & ROUNDUP(MONTH(J2)/3, 0)-1)"
'wb.Sheets(2).Range("J4").Value = "=IF(J3 = " & Chr(34) & "Q4" & Chr(34) & ", YEAR(J2)-1, YEAR(J2))"
'mabi = ws.Range("N" & i).Value
'mabi = wb.Sheets(2).Range("J3").Value 'Q
'maCIK

'mabi = Month(ws.Range("N" & i).Value)
'Debug.Print "Mabi month = " & mabi
mabi = IIf(Application.WorksheetFunction.RoundUp(Month(ws.Range("N" & i).Value) / 3, 0) = 1, "Q4", "Q" & Application.WorksheetFunction.RoundUp(Month(ws.Range("N" & i).Value) / 3, 0) - 1)
maCIK = IIf(mabi = "Q4", Year(ws.Range("N" & i).Value) - 1, Year(ws.Range("N" & i).Value))

mabi = mabi & maCIK
maStrPath = "C:\AM\" & mabi & "CUSIPS.xlsx"


'Open the CUSIPS file
Workbooks.Open (maStrPath)

'Set its filename into maStrPath
maStrPath = mabi & "CUSIPS.xlsx"
Set mabb = Workbooks(maStrPath)



End Sub


Private Sub CusipHRA()

Dim wb As Workbook
Dim ws As Worksheet

Dim i As String

Dim mabi As String
Dim maCIK As String

Set wb = Workbooks("CIK.xlsm")
Set ws = wb.Sheets(1)

i = 10

Dim maStrPath As String
Dim mabb As Workbook
Dim mawbks As String
Dim maDText As String



maStrPath = ws.Range("G" & i).Value & ws.Range("F" & i).Value
Workbooks.Open (maStrPath)
'
''Set strPath as the filename, and set the master into BB
'
maStrPath = ws.Range("F" & i).Value
Set mabb = Workbooks("Master_Engine Capital_0001665590.xlsx")

mawbks = mabb.Worksheets.Count

For maz = 3 To mawbks

If mabb.Worksheets(maz).Name Like "*HRA*" Then
'Do nothing, it's an HRA, go to the next

Else

'This is the previous 13F-HR, get the last 8 chars and manipulate it to make a date field in format mm/dd/yyyy
maDText = Mid(Right(mabb.Worksheets(maz).Name, 8), 5, 2) & "/"
maDText = maDText & Right(mabb.Worksheets(maz).Name, 2) & "/"
maDText = maDText & Left(Right(mabb.Worksheets(maz).Name, 8), 4)

'Close the Master file
mabb.Close False

'put the Q into mabi and the Y into maCIK
mabi = IIf(Application.WorksheetFunction.RoundUp(Month(maDText) / 3, 0) = 1, "Q4", "Q" & Application.WorksheetFunction.RoundUp(Month(maDText) / 3, 0) - 1)
maCIK = IIf(mabi = "Q4", Year(maDText) - 1, Year(maDText))

Exit For

End If

Next maz

'Write mabi and maCIK together to make a string e.g. Q32016 which will help form the filepath of the correct CUSIP file to open
mabi = mabi & maCIK
maStrPath = "C:\AM\" & mabi & "CUSIPS.xlsx"

'Open the CUSIPS file
Workbooks.Open (maStrPath)

'Set its filename into maStrPath
maStrPath = mabi & "CUSIPS.xlsx"
Set mabb = Workbooks(maStrPath)


End Sub
