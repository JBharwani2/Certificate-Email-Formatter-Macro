Sub CertificateEmailFormatter()
    '
    ' CertificateEmailFormatter Macro
    ' Created by Jeremy Bharwani on 6/7/21
    ' Updated by Jeremy Bharwani on 7/14/21
    ' (questions- email jcb926@gmail.com)
    '
    ' Iterates through selected worksheets within CHS workbook and reformats each sheet into its own workbook that is saved in the corresponding file under
    ' the 'Emails' folder. If the correctly dated folder does not already exist, a new folder is created.
    '
    ' Time: 1.25 seconds per sheet
    ' References: "mscorlib"
    '

    'VARIABLES -------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim fileDate As String
    Dim month As String
    Dim year As String
    Dim strFolderName As String
    Dim strFolderExists As String
    Dim message As String
    Dim count As Integer
    count = 0

    'DATE SETUP ------------------------------------------------------------------------------------------------------------------------------------------------------
    'Gets user input for the month and year of this batch of files
    fileDate = Left(ThisWorkbook.Name, 5)
    month = Left(fileDate, 2)
    year = Right(fileDate, 2)

    'Asks user if they are sure that they want to run the macro with the selected sheets
    CarryOn = MsgBox("You have selected " & ActiveWindow.SelectedSheets.count & " sheets for the month of " & month & "-" & year &
        ". Do you want to create email formatted files for these sheets?", vbYesNo, "Macro Run Confirmation")

    If CarryOn = vbYes Then

        'FOLDER SETUP ----------------------------------------------------------------------------------------------------------------------------------------------------
        'Checks if the necessary folder exists, if not it creates a new folder
        strFolderName = "{REMOVED FILEPATH}" & year & "\" & year & "-CHS\Email\" & month & year
        strFolderExists = Dir(strFolderName, vbDirectory)
        If strFolderExists = "" Then
            MkDir("{REMOVED FILEPATH}" & year & "\" & year & "-CHS\Email\" & month & year)
        End If

        'COPY & SAVE -----------------------------------------------------------------------------------------------------------------------------------------------------
        Application.ScreenUpdating = False 'aviods the visible opening and closing of the new workbooks
        For Each ws In ActiveWindow.SelectedSheets
            'Move sheet to an entirely new workbook
            Sheets(ws.Name).Select
            Sheets(ws.Name).Copy

            'Reformat equations to values
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False

            'Saves as a new file in the correct folder (S:\acct\JLS\2021\21-CHS\Email\0521)
            ActiveWorkbook.SaveAs Filename:=strFolderName & "\CPS " & ws.Name & " " & month & year, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            ActiveWindow.Close

            count = count + 1
        Next ws
        Application.ScreenUpdating = True

        'COMPLETION ------------------------------------------------------------------------------------------------------------------------------------------------------
        CompleteMsg = MsgBox(count & " sheets successfully formatted and saved to '" & strFolderName & "'", vbOKOnly, "Macro Run Confirmation")
    End If
End Sub
