' Text Scripting on VBA v1.2.0 (Safety-first fork)
' last update: 2026-02-14
' Original: http://rsh.csh.sh/text-scripting-vba/
'
' This workbook code is designed for Excel workbook/add-in usage.

Option Explicit


'----------------------------- Consts ---------------

' Library list file (placed in the same directory as this workbook/add-in)
Const FILENAME_LIBLIST As String = "libdef.txt"
Const FILENAME_EXPORT As String = "ThisWorkbook-sjis.cls"
Const FILENAME_OPERATION_LOG As String = "reload-log.txt"
Const NAME_RELEASE_MODE As String = "_TSVBA_RELEASE_MODE"

' Run loader automatically when workbook/add-in opens
Const ENABLE_WORKBOOK_OPEN As Boolean = False

' Shortcut key: Ctrl + r
Const SHORTKEY_RELOAD As String = "r"

' Safety settings (all safe options enabled by default)
Const SAFE_ENABLE_PREFLIGHT As Boolean = True
Const SAFE_MANAGED_ONLY As Boolean = True
Const SAFE_ENABLE_BACKUP As Boolean = True
Const SAFE_ENABLE_ROLLBACK As Boolean = True
Const SAFE_REQUIRE_FORM_FRX As Boolean = True
Const SAFE_WARN_SIGNATURE_RISK As Boolean = True
Const SAFE_ENABLE_TEXT_LOG As Boolean = True
Const SAFE_PROMPT_RELEASE_ON_CLOSE As Boolean = True


'----------------------------- Workbook_open() ---------------

Private Sub Workbook_Open()
  If ENABLE_WORKBOOK_OPEN = False Then Exit Sub
  If isReleaseMode() Then Exit Sub

  Call setShortKey
  Call reloadComponents
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  If SAFE_PROMPT_RELEASE_ON_CLOSE And (Not isReleaseMode()) Then
    If confirmReleaseModeOnClose(Cancel) Then
      If Cancel Then Exit Sub
    End If
  End If

  Call clearShortKey
End Sub


'----------------------------- public Subs/Functions ---------------

Private Sub runReloadComponents()
  If isReleaseMode() Then
    MsgBox "Release mode is ON. Reload is disabled."
    Exit Sub
  End If

  Dim msgResult As String
  Dim statusText As String
  msgResult = loadModule("." & Application.PathSeparator & FILENAME_LIBLIST)

  statusText = getOperationStatus(msgResult)
  If SAFE_ENABLE_TEXT_LOG Then
    Call appendOperationLog("reloadComponents", statusText, msgResult)
  End If

  If Len(msgResult) > 0 Then
    MsgBox msgResult
  End If
End Sub

Public Sub enableReleaseMode()
  Call setReleaseMode(True)
  Call clearShortKey
  MsgBox "Release mode enabled."
End Sub

Private Sub disableReleaseMode()
  Call setReleaseMode(False)
  Call setShortKey
  MsgBox "Release mode disabled."
End Sub

Public Sub reloadComponents()
  Call runReloadComponents
End Sub

Public Sub exportThisWorkbook()
  Call exportModule("ThisWorkbook", FILENAME_EXPORT)
End Sub

Public Sub exportComponents()
  If isReleaseMode() Then
    MsgBox "Release mode is ON. Export is disabled."
    Exit Sub
  End If

  Dim msgResult As String
  msgResult = exportManagedComponents("." & Application.PathSeparator & FILENAME_LIBLIST)

  If Len(msgResult) > 0 Then
    MsgBox msgResult
  End If
End Sub


'----------------------------- main Subs/Functions ---------------

Private Function loadModule(ByVal pathConf As String) As String
  Dim msg As String
  msg = ""

  Dim importPaths As Collection
  Set importPaths = New Collection

  Dim importNames As Collection
  Set importNames = New Collection

  pathConf = absPath(pathConf)
  If checkExistFile(pathConf) = False Then
    loadModule = "Error: library list file not found: " & pathConf
    Exit Function
  End If

  msg = collectImportPlan(pathConf, importPaths, importNames)
  If Len(msg) > 0 Then
    loadModule = msg
    Exit Function
  End If

  If importPaths.Count = 0 Then
    loadModule = "Error: no valid component entries in library list."
    Exit Function
  End If

  If SAFE_ENABLE_PREFLIGHT Then
    msg = preflightImportPlan(importPaths, importNames)
    If Len(msg) > 0 Then
      loadModule = msg
      Exit Function
    End If
  End If

  Dim backupDir As String
  backupDir = ""

  If SAFE_ENABLE_BACKUP Then
    backupDir = createBackupSnapshot(msg)
    If Len(msg) > 0 Then
      loadModule = msg
      Exit Function
    End If
  End If

  msg = applyImportPlan(importPaths, importNames)
  If Len(msg) > 0 Then
    If SAFE_ENABLE_ROLLBACK And SAFE_ENABLE_BACKUP And Len(backupDir) > 0 Then
      Dim msgRollback As String
      msgRollback = rollbackFromBackup(backupDir)
      loadModule = msg & vbCrLf & msgRollback
    Else
      loadModule = msg
    End If
    Exit Function
  End If

  If SAFE_WARN_SIGNATURE_RISK Then
    msg = msg & "Warning: editing VBProject may invalidate VBA digital signatures." & vbCrLf
  End If

  If SAFE_ENABLE_BACKUP And Len(backupDir) > 0 Then
    msg = msg & "Backup created: " & backupDir
  End If

  loadModule = msg
End Function

Private Function exportManagedComponents(ByVal pathConf As String) As String
  Dim msg As String
  msg = ""

  pathConf = absPath(pathConf)
  If Len(pathConf) = 0 Then
    exportManagedComponents = "Error: failed to resolve libdef path."
    Exit Function
  End If

  If checkExistFile(pathConf) = False Then
    exportManagedComponents = "Error: library list file not found: " & pathConf
    Exit Function
  End If

  Dim listPaths As Collection
  Set listPaths = New Collection

  Dim listNames As Collection
  Set listNames = New Collection

  Call collectLibdefEntries(pathConf, listPaths, listNames)

  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
    If isRemovableComponentType(component.Type) Then
      Dim ext As String
      ext = getExportExtension(component.Type)
      If Len(ext) = 0 Then GoTo NextComponent

      Dim exportPath As String
      exportPath = findPathByComponentName(listPaths, listNames, component.Name)

      If Len(exportPath) = 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox( _
          "Module '" & component.Name & "' is not listed in libdef.txt." & vbCrLf & _
          "Add it and export now?", _
          vbYesNoCancel + vbQuestion, _
          "Export Components")

        If resp = vbCancel Then
          exportManagedComponents = "Canceled."
          Exit Function
        End If

        If resp = vbYes Then
          exportPath = ThisWorkbook.Path & Application.PathSeparator & component.Name & "." & ext
          If appendLibdefEntry(pathConf, component.Name & "." & ext) Then
            listPaths.Add exportPath
            listNames.Add component.Name
          Else
            msg = msg & "Failed to append libdef entry for " & component.Name & vbCrLf
            GoTo NextComponent
          End If
        Else
          GoTo NextComponent
        End If
      End If

      If LCase$(getFileExtension(exportPath)) <> LCase$(ext) Then
        exportPath = replaceExtension(exportPath, ext)
      End If

      If ensureParentFolder(exportPath) = False Then
        msg = msg & "Failed to create folder for " & exportPath & vbCrLf
        GoTo NextComponent
      End If

      On Error Resume Next
      component.Export exportPath
      If Err.Number <> 0 Then
        msg = msg & "Failed to export " & component.Name & " to " & exportPath & ": " & Err.Description & vbCrLf
        Err.Clear
      End If
      On Error GoTo 0
    End If
NextComponent:
  Next component

  If Len(msg) = 0 Then
    msg = "Export completed."
  End If

  exportManagedComponents = msg
End Function

Private Function collectImportPlan(ByVal pathConf As String, ByRef importPaths As Collection, ByRef importNames As Collection) As String
  Dim listEntries As Variant
  listEntries = list2array(pathConf)

  If UBound(listEntries) = 0 Then
    collectImportPlan = ""
    Exit Function
  End If

  Dim i As Integer
  Dim msg As String
  msg = ""

  For i = 0 To UBound(listEntries) - 1
    Dim p As String
    p = absPath(Trim$(listEntries(i)))

    If Len(p) = 0 Then GoTo NextEntry

    Dim ext As String
    ext = LCase$(getFileExtension(p))

    If isSupportedExtension(ext) = False Then
      msg = msg & "Unsupported extension: " & p & vbCrLf
      GoTo NextEntry
    End If

    If checkExistFile(p) = False Then
      msg = msg & "Missing file: " & p & vbCrLf
      GoTo NextEntry
    End If

    If ext = "frm" And SAFE_REQUIRE_FORM_FRX Then
      Dim frxPath As String
      frxPath = replaceExtension(p, "frx")
      If checkExistFile(frxPath) = False Then
        msg = msg & "Missing paired .frx for form: " & frxPath & vbCrLf
        GoTo NextEntry
      End If
    End If

    importPaths.Add p
    importNames.Add getBaseFileName(p)

NextEntry:
  Next i

  collectImportPlan = msg
End Function

Private Sub collectLibdefEntries(ByVal pathConf As String, ByRef outPaths As Collection, ByRef outNames As Collection)
  Dim listEntries As Variant
  listEntries = list2array(pathConf)

  If UBound(listEntries) = 0 Then Exit Sub

  Dim i As Long
  For i = 0 To UBound(listEntries) - 1
    Dim p As String
    p = absPath(Trim$(listEntries(i)))
    If Len(p) = 0 Then GoTo NextEntry
    If isSupportedExtension(LCase$(getFileExtension(p))) = False Then GoTo NextEntry

    outPaths.Add p
    outNames.Add getBaseFileName(p)

NextEntry:
  Next i
End Sub

Private Function preflightImportPlan(ByRef importPaths As Collection, ByRef importNames As Collection) As String
  Dim msg As String
  msg = ""

  Dim i As Long
  For i = 1 To importNames.Count
    If countTextInCollection(importNames, importNames(i)) > 1 Then
      msg = msg & "Duplicate target component name in list: " & importNames(i) & vbCrLf
    End If
  Next i

  Dim hasAnyManaged As Boolean
  hasAnyManaged = (importPaths.Count > 0)
  If hasAnyManaged = False Then
    msg = msg & "No import targets detected." & vbCrLf
  End If

  preflightImportPlan = msg
End Function

Private Function applyImportPlan(ByRef importPaths As Collection, ByRef importNames As Collection) As String
  Dim msg As String
  msg = ""

  If SAFE_MANAGED_ONLY = False Then
    If clearModules() = False Then
      applyImportPlan = "Error: failed to clear VBA components."
      Exit Function
    End If
  End If

  Dim i As Long
  For i = 1 To importPaths.Count
    Dim targetName As String
    Dim declaredName As String
    targetName = importNames(i)
    declaredName = getDeclaredComponentName(importPaths(i))

    If SAFE_MANAGED_ONLY Then
      Call removeComponentByName(targetName)
      If Len(declaredName) > 0 Then
        If LCase$(declaredName) <> LCase$(targetName) Then
          Call removeComponentByName(declaredName)
        End If
      End If
    End If

    Dim imported As Object
    On Error Resume Next
    Set imported = ThisWorkbook.VBProject.VBComponents.Import(importPaths(i))
    If Err.Number <> 0 Then
      msg = "Error: failed to import " & importPaths(i) & ": " & Err.Description
      Err.Clear
      On Error GoTo 0
      Exit For
    End If

    imported.Name = targetName
    If Err.Number <> 0 Then
      msg = "Error: imported but failed to rename to " & targetName & ": " & Err.Description
      Err.Clear
      On Error GoTo 0
      Exit For
    End If
    On Error GoTo 0
  Next i

  applyImportPlan = msg
End Function

Private Function createBackupSnapshot(ByRef msg As String) As String
  Dim baseDir As String
  baseDir = ThisWorkbook.Path & Application.PathSeparator & "backup"

  If ensureFolder(baseDir) = False Then
    msg = "Error: failed to create backup folder: " & baseDir
    createBackupSnapshot = ""
    Exit Function
  End If

  Dim stamp As String
  stamp = formatTimestamp(Now)

  Dim backupDir As String
  backupDir = baseDir & Application.PathSeparator & stamp

  If ensureFolder(backupDir) = False Then
    msg = "Error: failed to create backup snapshot folder: " & backupDir
    createBackupSnapshot = ""
    Exit Function
  End If

  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
    If isRemovableComponentType(component.Type) Then
      Dim ext As String
      ext = getExportExtension(component.Type)

      If Len(ext) = 0 Then GoTo NextBackup

      On Error Resume Next
      component.Export backupDir & Application.PathSeparator & component.Name & "." & ext
      If Err.Number <> 0 Then
        msg = "Error: failed to export backup for " & component.Name & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        createBackupSnapshot = ""
        Exit Function
      End If
      On Error GoTo 0
    End If
NextBackup:
  Next component

  msg = ""
  createBackupSnapshot = backupDir
End Function

Private Function rollbackFromBackup(ByVal backupDir As String) As String
  Dim msg As String
  msg = "Rollback started." & vbCrLf

  If clearModules() = False Then
    rollbackFromBackup = msg & "Rollback failed: could not clear components before restore."
    Exit Function
  End If

  Dim msgRestore As String
  msgRestore = importFolderComponents(backupDir)

  If Len(msgRestore) > 0 Then
    rollbackFromBackup = msg & "Rollback failed: " & msgRestore
  Else
    rollbackFromBackup = msg & "Rollback completed from backup: " & backupDir
  End If
End Function

Private Function importFolderComponents(ByVal targetDir As String) As String
  Dim msg As String
  msg = ""

  msg = msg & importByPattern(targetDir, "*.bas")
  msg = msg & importByPattern(targetDir, "*.cls")
  msg = msg & importByPattern(targetDir, "*.frm")

  importFolderComponents = msg
End Function

Private Function importByPattern(ByVal targetDir As String, ByVal pattern As String) As String
  Dim msg As String
  msg = ""

  Dim p As String
  p = Dir(targetDir & Application.PathSeparator & pattern)

  Do While Len(p) > 0
    Dim fullPath As String
    fullPath = targetDir & Application.PathSeparator & p

    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Import fullPath
    If Err.Number <> 0 Then
      msg = msg & "Failed to restore " & fullPath & ": " & Err.Description & vbCrLf
      Err.Clear
    End If
    On Error GoTo 0

    p = Dir()
  Loop

  importByPattern = msg
End Function


'----------------------------- Utility ---------------

Private Sub exportModule(ByVal nameModule As String, ByVal nameFile As String)
  Dim component As Object

  For Each component In ThisWorkbook.VBProject.VBComponents
    If component.Name = nameModule Then
      component.Export ThisWorkbook.Path & Application.PathSeparator & nameFile
      MsgBox "Exported " & nameModule & " to " & ThisWorkbook.Path & Application.PathSeparator & nameFile
      Exit Sub
    End If
  Next component

  MsgBox "Error: module not found: " & nameModule
End Sub

Private Function clearModules() As Boolean
  Dim i As Long

  With ThisWorkbook.VBProject
    For i = .VBComponents.Count To 1 Step -1
      If isRemovableComponentType(.VBComponents(i).Type) Then
        .VBComponents.Remove .VBComponents(i)
      End If
    Next i
  End With

  clearModules = (countRemovableComponents() = 0)
End Function

Private Function removeComponentByName(ByVal componentName As String) As Boolean
  Dim i As Long

  removeComponentByName = False
  With ThisWorkbook.VBProject
    For i = .VBComponents.Count To 1 Step -1
      If LCase$(.VBComponents(i).Name) = LCase$(componentName) Then
        If isRemovableComponentType(.VBComponents(i).Type) Then
          .VBComponents.Remove .VBComponents(i)
          removeComponentByName = True
          Exit Function
        End If
      End If
    Next i
  End With
End Function

Private Function isRemovableComponentType(ByVal componentType As Integer) As Boolean
  ' 1=StdModule, 2=ClassModule, 3=MSForm
  isRemovableComponentType = (componentType = 1 Or componentType = 2 Or componentType = 3)
End Function

Private Function countRemovableComponents() As Long
  Dim i As Long
  Dim cnt As Long

  cnt = 0
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.Count
      If isRemovableComponentType(.VBComponents(i).Type) Then
        cnt = cnt + 1
      End If
    Next i
  End With

  countRemovableComponents = cnt
End Function

Private Function getExportExtension(ByVal componentType As Integer) As String
  Select Case componentType
    Case 1
      getExportExtension = "bas"
    Case 2
      getExportExtension = "cls"
    Case 3
      getExportExtension = "frm"
    Case Else
      getExportExtension = ""
  End Select
End Function

Private Function getBaseFileName(ByVal pathFile As String) As String
  Dim fileName As String
  fileName = Mid$(pathFile, InStrRev(pathFile, Application.PathSeparator) + 1)

  Dim dotPos As Long
  dotPos = InStrRev(fileName, ".")

  If dotPos > 1 Then
    getBaseFileName = Left$(fileName, dotPos - 1)
  Else
    getBaseFileName = fileName
  End If
End Function

Private Function getDeclaredComponentName(ByVal pathFile As String) As String
  On Error GoTo ReadErr

  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp

  Do Until EOF(fp)
    Dim lineText As String
    Line Input #fp, lineText
    lineText = Trim$(lineText)
    If Len(lineText) > 0 Then
      If AscW(Left$(lineText, 1)) = 65279 Then
        lineText = Mid$(lineText, 2)
      End If
    End If

    If InStr(1, lineText, "Attribute VB_Name =", vbTextCompare) > 0 Then
      Dim q1 As Long
      Dim q2 As Long

      q1 = InStr(1, lineText, Chr$(34), vbBinaryCompare)
      If q1 > 0 Then
        q2 = InStr(q1 + 1, lineText, Chr$(34), vbBinaryCompare)
        If q2 > q1 Then
          getDeclaredComponentName = Mid$(lineText, q1 + 1, q2 - q1 - 1)
          Close #fp
          Exit Function
        End If
      End If
    End If
  Loop

  Close #fp
  getDeclaredComponentName = ""
  Exit Function

ReadErr:
  On Error Resume Next
  Close #fp
  getDeclaredComponentName = ""
End Function

Private Function getFileExtension(ByVal pathFile As String) As String
  Dim dotPos As Long
  dotPos = InStrRev(pathFile, ".")

  If dotPos > 0 And dotPos < Len(pathFile) Then
    getFileExtension = Mid$(pathFile, dotPos + 1)
  Else
    getFileExtension = ""
  End If
End Function

Private Function replaceExtension(ByVal pathFile As String, ByVal newExt As String) As String
  Dim dotPos As Long
  dotPos = InStrRev(pathFile, ".")

  If dotPos > 0 Then
    replaceExtension = Left$(pathFile, dotPos - 1) & "." & newExt
  Else
    replaceExtension = pathFile & "." & newExt
  End If
End Function

Private Function isSupportedExtension(ByVal ext As String) As Boolean
  isSupportedExtension = (ext = "bas" Or ext = "cls" Or ext = "frm")
End Function

Private Function countTextInCollection(ByRef items As Collection, ByVal keyText As String) As Long
  Dim i As Long
  Dim cnt As Long

  cnt = 0
  For i = 1 To items.Count
    If LCase$(items(i)) = LCase$(keyText) Then
      cnt = cnt + 1
    End If
  Next i

  countTextInCollection = cnt
End Function

Private Function ensureFolder(ByVal targetDir As String) As Boolean
  On Error GoTo MkDirErr

  If Len(Dir(targetDir, vbDirectory)) > 0 Then
    ensureFolder = True
    Exit Function
  End If

  MkDir targetDir
  ensureFolder = True
  Exit Function

MkDirErr:
  ensureFolder = False
End Function

Private Function ensureParentFolder(ByVal pathFile As String) As Boolean
  Dim parentDir As String
  parentDir = getParentDir(pathFile)
  If Len(parentDir) = 0 Then
    ensureParentFolder = False
    Exit Function
  End If

  ensureParentFolder = ensureFolderTree(parentDir)
End Function

Private Function ensureFolderTree(ByVal targetDir As String) As Boolean
  On Error GoTo MkTreeErr

  If Len(Dir(targetDir, vbDirectory)) > 0 Then
    ensureFolderTree = True
    Exit Function
  End If

  Dim parentDir As String
  parentDir = getParentDir(targetDir)
  If Len(parentDir) > 0 Then
    If Len(Dir(parentDir, vbDirectory)) = 0 Then
      If ensureFolderTree(parentDir) = False Then
        ensureFolderTree = False
        Exit Function
      End If
    End If
  End If

  MkDir targetDir
  ensureFolderTree = True
  Exit Function

MkTreeErr:
  ensureFolderTree = (Len(Dir(targetDir, vbDirectory)) > 0)
End Function

Private Function getParentDir(ByVal pathText As String) As String
  Dim pos As Long
  pos = InStrRev(pathText, Application.PathSeparator)
  If pos <= 1 Then
    getParentDir = ""
  Else
    getParentDir = Left$(pathText, pos - 1)
  End If
End Function

Private Function findPathByComponentName(ByRef paths As Collection, ByRef names As Collection, ByVal componentName As String) As String
  Dim i As Long
  For i = 1 To names.Count
    If LCase$(names(i)) = LCase$(componentName) Then
      findPathByComponentName = paths(i)
      Exit Function
    End If
  Next i

  findPathByComponentName = ""
End Function

Private Function appendLibdefEntry(ByVal pathConf As String, ByVal fileNameWithExt As String) As Boolean
  On Error GoTo AppendErr

  Dim needLeadingNewLine As Boolean
  needLeadingNewLine = fileNeedsLeadingNewLine(pathConf)

  Dim fp As Integer
  fp = FreeFile
  Open pathConf For Append As #fp
  If needLeadingNewLine Then
    Print #fp, ""
  End If
  Print #fp, "./" & fileNameWithExt
  Close #fp

  appendLibdefEntry = True
  Exit Function

AppendErr:
  On Error Resume Next
  Close #fp
  appendLibdefEntry = False
End Function

Private Function fileNeedsLeadingNewLine(ByVal pathFile As String) As Boolean
  On Error GoTo ReadErr

  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Binary Access Read As #fp

  If LOF(fp) = 0 Then
    Close #fp
    fileNeedsLeadingNewLine = False
    Exit Function
  End If

  Dim lastChar As String * 1
  Get #fp, LOF(fp), lastChar
  Close #fp

  fileNeedsLeadingNewLine = Not (lastChar = vbCr Or lastChar = vbLf)
  Exit Function

ReadErr:
  On Error Resume Next
  Close #fp
  fileNeedsLeadingNewLine = False
End Function

Private Function formatTimestamp(ByVal dt As Date) As String
  formatTimestamp = Format$(dt, "yyyymmdd_HHNNSS")
End Function

Private Function absPath(ByVal pathFile As String) As String
  Dim nameOS As String
  nameOS = Application.OperatingSystem

  pathFile = Replace(pathFile, Chr(92), Application.PathSeparator)
  pathFile = Replace(pathFile, ":", Application.PathSeparator)
  pathFile = Replace(pathFile, "/", Application.PathSeparator)

  Select Case Left(pathFile, 1)
    Case "."
      Select Case Left(pathFile, 2)
        Case ".."
          absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
          Exit Function
        Case Else
          absPath = ThisWorkbook.Path & Mid(pathFile, 2, Len(pathFile) - 1)
          Exit Function
      End Select

    Case Application.PathSeparator
      If Left(pathFile, 2) = Chr(92) & Chr(92) Then
        absPath = pathFile
        Exit Function
      Else
        absPath = pathFile
        Exit Function
      End If
  End Select

  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & Application.PathSeparator Then
    absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
    Exit Function
  End If

  If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
    absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
    Exit Function
  End If

  absPath = ""
End Function

Private Function checkExistFile(ByVal pathFile As String) As Boolean
  On Error GoTo ErrCheck

  If Dir(pathFile) = "" Then
    checkExistFile = False
  Else
    checkExistFile = True
  End If
  Exit Function

ErrCheck:
  checkExistFile = False
End Function

Private Function list2array(ByVal pathFile As String) As Variant
  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp

  Dim arrayOutput() As String
  Dim countLine As Integer
  countLine = 0
  ReDim Preserve arrayOutput(countLine)

  Do Until EOF(fp)
    Dim strLine As String
    Line Input #fp, strLine
    strLine = Replace(strLine, vbCr, "")
    strLine = Trim$(strLine)

    If Len(strLine) > 0 Then
      If Left$(strLine, 1) <> "'" Then
        countLine = countLine + 1
        ReDim Preserve arrayOutput(countLine)
        arrayOutput(countLine - 1) = strLine
      End If
    End If
  Loop

  Close #fp
  list2array = arrayOutput
End Function

Private Sub setShortKey()
  If Application.OperatingSystem Like "Windows *" Then
    Application.MacroOptions Macro:="ThisWorkbook.reloadComponents", ShortcutKey:=SHORTKEY_RELOAD
  Else
    Application.OnKey "^" & SHORTKEY_RELOAD, "ThisWorkbook.reloadComponents"
  End If
End Sub

Private Sub clearShortKey()
  If Application.OperatingSystem Like "Windows *" Then
    Application.MacroOptions Macro:="ThisWorkbook.reloadComponents", ShortcutKey:=""
  Else
    Application.OnKey "^" & SHORTKEY_RELOAD, ""
  End If
End Sub

Private Function confirmReleaseModeOnClose(ByRef Cancel As Boolean) As Boolean
  Dim currentMode As Boolean
  currentMode = isReleaseMode()

  Dim prompt As String
  prompt = "Switch this workbook to RELEASE MODE before close?" & vbCrLf & _
           "Yes: enable release mode" & vbCrLf & _
           "No: keep current mode (" & modeText(currentMode) & ")" & vbCrLf & _
           "Cancel: abort close"

  Dim result As VbMsgBoxResult
  result = MsgBox(prompt, vbYesNoCancel + vbQuestion, "Release Mode")

  Select Case result
    Case vbYes
      setReleaseMode True
      confirmReleaseModeOnClose = True
    Case vbNo
      confirmReleaseModeOnClose = True
    Case vbCancel
      Cancel = True
      confirmReleaseModeOnClose = True
    Case Else
      confirmReleaseModeOnClose = False
  End Select
End Function

Private Function modeText(ByVal modeValue As Boolean) As String
  If modeValue Then
    modeText = "ON"
  Else
    modeText = "OFF"
  End If
End Function

Private Function isReleaseMode() As Boolean
  On Error GoTo NoFlag

  Dim nm As Name
  Set nm = ThisWorkbook.Names(NAME_RELEASE_MODE)
  isReleaseMode = CBool(Evaluate(nm.RefersTo))
  Exit Function

NoFlag:
  isReleaseMode = False
End Function

Private Sub setReleaseMode(ByVal enabled As Boolean)
  Dim refersExpr As String
  If enabled Then
    refersExpr = "=TRUE"
  Else
    refersExpr = "=FALSE"
  End If

  On Error GoTo AddNew
  ThisWorkbook.Names(NAME_RELEASE_MODE).RefersTo = refersExpr
  If enabled Then
    Call clearShortKey
  End If
  Exit Sub

AddNew:
  Err.Clear
  ThisWorkbook.Names.Add Name:=NAME_RELEASE_MODE, RefersTo:=refersExpr, Visible:=False
  If enabled Then
    Call clearShortKey
  End If
End Sub

Private Function getOperationStatus(ByVal msgText As String) As String
  If InStr(1, msgText, "Error:", vbTextCompare) > 0 Then
    getOperationStatus = "FAILED"
  ElseIf Len(msgText) = 0 Then
    getOperationStatus = "SUCCESS"
  Else
    getOperationStatus = "SUCCESS_WITH_NOTICE"
  End If
End Function

Private Sub appendOperationLog(ByVal actionName As String, ByVal statusText As String, ByVal detailText As String)
  On Error GoTo LogErr

  Dim logPath As String
  logPath = ThisWorkbook.Path & Application.PathSeparator & FILENAME_OPERATION_LOG

  Dim fp As Integer
  fp = FreeFile
  Open logPath For Append As #fp

  Print #fp, "-----"
  Print #fp, "[" & Format$(Now, "yyyy-mm-dd HH:nn:ss") & "] " & actionName & " " & statusText
  Print #fp, "Workbook: " & ThisWorkbook.Name
  Print #fp, "Path: " & ThisWorkbook.Path
  Print #fp, "Component snapshot:"
  Print #fp, buildComponentSnapshot()
  If Len(detailText) > 0 Then
    Print #fp, "Details:"
    Print #fp, sanitizeLogText(detailText)
  End If

  Close #fp
  Exit Sub

LogErr:
  On Error Resume Next
  Close #fp
End Sub

Private Function buildComponentSnapshot() As String
  Dim snapshot As String
  snapshot = ""

  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
    If isRemovableComponentType(component.Type) Then
      snapshot = snapshot & " - " & component.Name & " (" & componentTypeName(component.Type) & ")" & vbCrLf
    End If
  Next component

  If Len(snapshot) = 0 Then
    snapshot = " - (none)" & vbCrLf
  End If

  buildComponentSnapshot = snapshot
End Function

Private Function componentTypeName(ByVal componentType As Integer) As String
  Select Case componentType
    Case 1
      componentTypeName = "StdModule"
    Case 2
      componentTypeName = "ClassModule"
    Case 3
      componentTypeName = "UserForm"
    Case Else
      componentTypeName = "Other"
  End Select
End Function

Private Function sanitizeLogText(ByVal rawText As String) As String
  Dim t As String
  t = Replace(rawText, vbCrLf, vbLf)
  t = Replace(t, vbCr, vbLf)
  t = Replace(t, vbLf, vbCrLf)
  sanitizeLogText = t
End Function
