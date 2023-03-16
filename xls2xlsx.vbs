Option Explicit

' 将xls文件批量转为xlsx格式
' 操作：支持拖放单xls文件/多xls文件/目录/混合
' 注意：本机需要安装Office2007+；转换xlsx后会自动删除旧xls文件
' 更新：https://github.com/playGitboy/xls2xlsx
Sub main()
	Dim objArgs,objArg
	Set objArgs = WScript.Arguments
	For Each objArg in objArgs
		ProcessFile(objArg)
	Next
End Sub

Function FileExists(FilePath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
        FileExists=CBool(1)
    Else
        FileExists=CBool(0)
    End If
End Function

Function ProcessFile(Path)
    Dim TotalFiles, Objshell
    TotalFiles = 0
    Set Objshell = CreateObject("scripting.filesystemobject")
    If Objshell.FolderExists(Path) Then
    Dim FolderPath, File, Files, SubFolder, Folder
        Set FolderPath = Objshell.GetFolder(Path)
        Set Files = FolderPath.Files
        For Each File In Files
            if convertOfficeFile(File.Path) Then
                TotalFiles = TotalFiles + 1
            End If
        Next
        Set SubFolder = FolderPath.SubFolders
        For Each Folder In SubFolder
            TotalFiles = TotalFiles + ProcessFile(Folder)
        Next
    Elseif convertOfficeFile(Path) Then
        TotalFiles = TotalFiles + 1
    End If
    ProcessFile = TotalFiles
End Function

Function convertOfficeFile(Path)
    If isXlsFile(Path) Then
        convertXlsToXlsx Path
        convertOfficeFile = true
    Else
        convertOfficeFile = false
    End If
End Function

Function convertXlsToXlsx(Path)
    Dim Objshell, ParentFolder, BaseName, XlsApp, Doc, XlsxPath
    Set Objshell = CreateObject("scripting.filesystemobject")
    ParentFolder = Objshell.GetParentFolderName(Path)
    BaseName = Objshell.GetBaseName(Path)
    XlsxPath = parentFolder & "\" & BaseName & ".xlsx"
    If not FileExists(XlsxPath) Then
        Set XlsApp = CreateObject("Excel.application")
        Set Doc = XlsApp.Workbooks.Open(Path)
        Doc.SaveAs XlsxPath,51
        Doc.close False
        XlsApp.Quit
        Objshell.DeleteFile(Path)
    End If
    Set Objshell = Nothing
End Function

Function isXlsFile(Path)
    Dim Objshell
    Set Objshell = CreateObject("scripting.filesystemobject")
    Dim Arrs, Arr
    Arrs = Array("xls","xlsx")
    Dim FileExtension
    isXlsFile = False
    FileExtension = Objshell.GetExtensionName(Path)
    For Each Arr In Arrs
        If InStr(UCase(FileExtension),UCase(Arr)) <> 0 Then
            isXlsFile = True
            Exit For
        End If
    Next
    Set Objshell = Nothing
End Function

Call main