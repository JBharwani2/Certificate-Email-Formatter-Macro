Imports Microsoft.Office.Interop.Excel
Imports System.Activities.Presentation.View
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Xml

Public Class CHSEmailFormat
    Sub CHSEmailFormat()
        '
        ' CHSEmailFormat Macro
        ' Created by Jeremy Bharwani on 6/7/21
        ' (questions- email jcb926@gmail.com)
        '
        ' Iterates through worksheets within CHS document and reformats each sheet into its own workbook that is saved in the corresponding file under the 'Emails' folder.
        ' References Used: "mscorlib"
        '

        'VARIABLES -------------------------------------------------------------------------------------------------------------------------------------------------------
        Dim ws As Worksheet
        Dim UserName As String
        Dim month As String
        Dim year As String
        Dim strFolderName As String
        Dim strFolderExists As String
        Dim count As Integer
        count = 0

        'DIRECTORY SETUP --------------------------------------------------------------------------------------------------------------------------------------------------
        'Gets user input for the month and year of this batch of files
        UserName = InputBox("Please input the month and year for these files in this format: '05/21'")
        month = Left(UserName, 2)
        year = Right(UserName, 2)

        'Checks if the necessary folder exists, if not it creates a new folder
        strFolderName = "S:\acct\DB\JB\20" + year + "\" + year + "-CHS\Email\" + month + year
        strFolderExists = Dir(strFolderName, vbDirectory)
        If strFolderExists = "" Then
            MkDir("S:\acct\DB\JB\20" + year + "\" + year + "-CHS\Email\" + month + year)
        End If

        'MAIN PROCESS -----------------------------------------------------------------------------------------------------------------------------------------------------
        For Each ws In ActiveWindow.SelectedSheets
            'Move sheet to an entirely new workbook
            Sheets(ws.Name).Select
            Sheets(ws.Name).Copy

            'Reformat equations to values
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False Application.CutCopyMode = False

            'Saves as a new file in the correct folder
            ActiveWorkbook.SaveAs Filename:=strFolderName + "\CPS " + ws.Name + " " + month + year, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        'S:\acct\JLS\2021\21-CHS\Email\0521       "S:\acct\JLS\20" &year& "\" &year& "-CHS\Email\" &month&year
            ActiveWindow.Close

            count = count + 1
        Next ws

        MsgBox Str(count) + " sheets successfully formatted and saved to " + strFolderName
End Sub

End Class
