Imports System.Globalization
Imports System.IO

Module Module1
    Private sDestPath As String, sFolderPrMth As String, checkFolder As String
    Sub Main()
        Call BaiDsoBackup()
        Call BbmiDsoBackup()

        'My.Computer.FileSystem.CopyFile(Environ("userprofile") & "\Desktop\addin_soresa.docx", "H:\me.docx")

        'Exit Sub

        'Dim objFso As Object, sDestPath As String, sFolderPrMth As String, checkFolder As String
        ''objFso = CreateObject("Scripting.FileSystemObject")

        ''sDestPath = "T:\Aristide_Lapa\MacrAris\Bai_Current_Month\BAI_Archive\BAI_Macro_" _
        ''& sdhLastDayInMonth(Now()) _
        ''& "_Backup_" & Format(Now, "yyyy-mm-dd hh-mm-ss")

        'sDestPath = Environ("UserProfile") & "\Desktop\Bai_Current_Month\BAI_Archive\BAI_Macro_" _
        '& sdhLastDayInMonth(Now()) _
        '& "_Backup_" & Format(Now, "yyyy-MM-dd_hh-mm-ss")

        'Console.WriteLine(sDestPath)
        ''sFolderPrMth = "T:\Aristide_Lapa\MacrAris\Bai_Current_Month\BAI_Macro"
        'sFolderPrMth = Environ("UserProfile") & "\Desktop\Bai_Current_Month\BAI_Macro\"
        'Console.WriteLine(sFolderPrMth)
        ''Console.ReadLine()

        'checkFolder = "BAI_Macro_" _
        '        & sdhLastDayInMonth(Now())
        ''checkFolder = Environ("UserProfile") & "\Desktop\Bai_Current_Month\BAI_Archive\" _
        ''            & checkFolder

        'Console.WriteLine(checkFolder)
        'Console.ReadLine()
        'Dim Cartella_scr As String = Environ("UserProfile") & "\Desktop\Bai_Current_Month\BAI_Archive\"

        'If VerificaFolder(checkFolder, Cartella_scr) = True Then
        '    Exit Sub
        '    ''    MsgBox "folder exists"
        'Else
        '    ' objFso.copyfolder(sFolderPrMth, sDestPath)
        '    My.Computer.FileSystem.CopyDirectory(sFolderPrMth, sDestPath, True)

        'End If
        ''Exit Sub

        'On Error Resume Next


        'For Each foundFile In My.Computer.FileSystem.GetFiles(
        '    sFolderPrMth)
        '    Select Case My.Computer.FileSystem.GetFileInfo(foundFile).Name.ToString.ToUpper
        '        Case "BAI_ANAGR.XLSX"
        '        Case "EXTRACT_BAI_INDICICLIENTI.XLS"
        '        Case "BAI_AG.XLSX"
        '        Case "FATT_MENSILE.XLS"
        '        Case "BAI_LISA_RAW.XLSX"
        '        Case "BAI_FATT_REVOLVING_YR.XLS"
        '        Case "MAX_PRIVATI_AGINGBAI.XLSX"
        '        Case Else : My.Computer.FileSystem.DeleteFile(foundFile)

        '    End Select
        'Next

        'For Each FoundDirectory In My.Computer.FileSystem.GetDirectories(sFolderPrMth)
        '    My.Computer.FileSystem.DeleteDirectory(FoundDirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
        'Next

        'Exit Sub
        'On Error Resume Next
        ''FSO.DeleteFolder(sFolderPrMth & "*.*", True)

        ''Dim myString As String = Format(Now, "yyyy-MM-dd hh-mm-ss tt")
        ''Console.WriteLine(myString)
        ''Console.ReadLine()
        ''Console.WriteLine(Now)
        ''Console.ReadLine()
        ''Dim date1 As Date = #5/8/2015#
        ''Console.WriteLine(date1.ToString("D", _
        ''                  CultureInfo.CreateSpecificCulture("it-it")))
        ''Console.ReadLine()

        ''Dim sLastDayMonth As String = sdhLastDayInMonth(Now()
        'Console.WriteLine(sdhLastDayInMonth(Now()))
        'Console.ReadLine()


    End Sub
    Sub BaiDsoBackup()
        sDestPath = Environ("UserProfile") & "\Desktop\BAI_Current_Month\BAI_Archive\BAI_Macro_" _
        & sdhLastDayInMonth(Now()) _
        & "_Backup_" & Format(Now, "yyyy-MM-dd_hh-mm-ss")

        sFolderPrMth = Environ("UserProfile") & "\Desktop\BAI_Current_Month\BAI_Macro\"

        checkFolder = "BAI_Macro_" _
                & sdhLastDayInMonth(Now())

        Dim Cartella_scr As String = Environ("UserProfile") & "\Desktop\BAI_Current_Month\BAI_Archive\"

        If VerificaFolder(checkFolder, Cartella_scr) Then '= True Then
            Exit Sub

        Else

            My.Computer.FileSystem.CopyDirectory(sFolderPrMth, sDestPath, True)

        End If

        On Error Resume Next
        For Each foundFile In My.Computer.FileSystem.GetFiles(
            sFolderPrMth)
            Select Case My.Computer.FileSystem.GetFileInfo(foundFile).Name.ToString.ToUpper
                Case "BAI_ANAGR.XLSX"
                Case "EXTRACT_BAI_INDICICLIENTI.XLS"
                Case "BAI_AG.XLSX"
                Case "FATT_MENSILE.XLS"
                Case "BAI_LISA_RAW.XLSX"
                Case "BAI_FATT_REVOLVING_YR.XLS"
                Case "MAX_PRIVATI_AGINGBAI.XLSX"
                Case Else : My.Computer.FileSystem.DeleteFile(foundFile)

            End Select
        Next

        For Each FoundDirectory In My.Computer.FileSystem.GetDirectories(sFolderPrMth)
            My.Computer.FileSystem.DeleteDirectory(FoundDirectory, _
                                                   FileIO.DeleteDirectoryOption.DeleteAllContents)
        Next

    End Sub
    Sub BbmiDsoBackup()

        sDestPath = Environ("UserProfile") & "\Desktop\BBMI_Current_Month\BBMI_Archive\BBMI_Macro_" _
        & sdhLastDayInMonth(Now()) _
        & "_Backup_" & Format(Now, "yyyy-MM-dd_hh-mm-ss")

        sFolderPrMth = Environ("UserProfile") & "\Desktop\BBMI_Current_Month\BBMI_Macro\"

        checkFolder = "BBMI_Macro_" _
                & sdhLastDayInMonth(Now())

        Dim Cartella_scr As String = Environ("UserProfile") & "\Desktop\BBMI_Current_Month\BBMI_Archive\"

        If VerificaFolder(checkFolder, Cartella_scr) Then '= True Then
            Exit Sub

        Else

            My.Computer.FileSystem.CopyDirectory(sFolderPrMth, sDestPath, True)

        End If

        On Error Resume Next

        For Each foundFile In My.Computer.FileSystem.GetFiles(
            sFolderPrMth)
            Select Case My.Computer.FileSystem.GetFileInfo(foundFile).Name.ToString.ToUpper
                Case "BBMI_ANAGR.XLSX"
                Case "EXTRACT_BBMI_INDICICLIENTI.XLS"
                Case "BBMI_AG.XLSX"
                Case "FATT_MENSILE.XLS"
                Case "BBMI_LISA_RAW.XLSX"
                Case "BBMI_FATT_REVOLVING_YR.XLS"
                    ' Case "MAX_PRIVATI_AGINGBBMI.XLSX"
                Case Else : My.Computer.FileSystem.DeleteFile(foundFile)

            End Select
        Next

        For Each FoundDirectory In My.Computer.FileSystem.GetDirectories(sFolderPrMth)
            My.Computer.FileSystem.DeleteDirectory(FoundDirectory, _
                                                   FileIO.DeleteDirectoryOption.DeleteAllContents)
        Next

    End Sub


    Function VerificaFolder(folderToCheck As String, srcFolder As String) As Boolean

        ' '' '''Source
        ' '' '''http://www.ozgrid.com/forum/showthread.php?t=69086
        ' '' '''http://www.mrexcel.com/forum/excel-questions/545447-visual-basic-applications-looping-through-directory-structure.html
        ' '' '''http://www.rondebruin.nl/win/s3/win026.htm

        ''Dim Cartella_scr As String = Environ("UserProfile") & "\Desktop\Bai_Current_Month\BAI_Archive\"

        '' 'WORKED ALREADY

        ' ''For Each fToCheck In My.Computer.FileSystem.GetDirectories(Cartella_scr)
        ' ''    If My.Computer.FileSystem.GetDirectoryInfo(fToCheck).ToString Like "*" & checkFolder & "*" Then
        ' ''        Return True

        ' ''    End If

        ' ''Next

        For Each foundDirectory As String In
               My.Computer.FileSystem.GetDirectories(srcFolder, _
                   FileIO.SearchOption.SearchTopLevelOnly, "*" & folderToCheck & "*")
        Next
        Return True
    End Function

    Function sdhLastDayInMonth(dtmDate As Date) As String
        '' Return the last day in the specified month.
        'If dtmDate = Now() Then
        '    ' Did the caller pass in a date? If not, use
        '    ' the current date.
        'dtmDate = Date
        'End If
        'dtmDate = Now()
        Dim dhLastDayInMonth As Date = DateSerial(Year(dtmDate), _
         Month(dtmDate), 0)
        sdhLastDayInMonth = Format(dhLastDayInMonth, "yyyy-MM")

    End Function

End Module