VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TheFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
'=============================================================================
'#Const USE_FILE_SYSTEM_OBJECT = 1
'   if defined, it uses FSO implementation instead of my VB implementation.
'


Option Explicit

Private Const PATH_SEPARATOR_CHAR = "\"

' Join the path only if the 3rd param is true
' Don't use this. INstead, use IIF( )
'Private Function joinPathIf(ByVal path1 As String, ByVal path2 As String, ByVal joinIt As Boolean) As String
'    If joinIt Then
'        joinPathIf = joinPath(path1, path2)
'    Else
'        joinPathIf = path2
'    End If
'End Function

' Returns Collection of files / folders.
' TODO: return TheList instead?  Or create a function that converts collection to list quickly in
' ie TheLIst::copy(Collection)
' includeFile = files are included in the "ls"
' includeDir = dir are included in the ls. BY default, it is NOT included.
' returnFullPath = print full path, not just the folder or file name
' recursive = -r
' classify = -F, add "/" at the end of folder. Doesn't add "*" after exec, though.

Public Function ls(ByVal folderName As String, Optional extension, _
    Optional recursive = True, Optional includeFile = True, Optional includeDir = False, _
    Optional returnFullPath = True, Optional ByVal isCurrentlyRecursing = False, Optional classify = True) _
    As Collection
    
    Dim i
    Static fileList As Collection
    'Dim tmpList As Collection
    Dim dirList As New Collection   'holds dir
    
    Dim path As String
    Dim currentFile As String 'tmp name for files/folders
    Dim filename As String
    
    If Not isCurrentlyRecursing Then
        'clearCollection fileList  'start new search
        Set fileList = New Collection
    End If
        'Debug.Print joinPath(folderName, filter)
    filename = Dir(joinPath(folderName, "*.*"), vbNormal + vbDirectory + vbHidden + vbSystem + vbReadOnly + vbArchive)
    
    Do Until filename = ""
        currentFile = joinPath(folderName, filename)
        If GetAttr(currentFile) And vbDirectory Then
        '------------ DIRECTORY -----------
            If filename = "." Or filename = ".." Then
                ' -A (--Almost-all) flag is assumed.
                'ignore . and ..
            Else
                
                currentFile = IIf(returnFullPath, joinPath(folderName, filename), filename)
                currentFile = IIf(classify, addTrailingPathSeparator(currentFile), currentFile)
                dirList.add currentFile
            End If
        Else
        '------------- FILE ------------
            'Debug.Print joinPath(folderName, filename)
            If includeFile Then
                currentFile = IIf(returnFullPath, joinPath(folderName, filename), filename)
                If IsMissing(extension) Then
                    fileList.add currentFile
                ElseIf hasExtension(currentFile, extension) Then
                    fileList.add currentFile
                End If
            End If
        End If          'end File
        filename = Dir
    Loop
    '---------- RECURSE DIR IF NEEDED -------------
    For Each i In dirList
        If includeDir Then  '--directory (-d) ???
            'currentFile = joinPathIf(folderName, filename, returnFullPath)
            'fileList.Add addTrailingPathSeparator(currentFile)
            fileList.add i
        Else
            'TODO: fileList.Add getLastDirName(i) 'strip prefix dir
        End If
        If recursive Then   '--recursive (-R)
           ls i, extension, recursive, includeFile, includeDir, _
                returnFullPath, True, classify
                
        End If 'recursive
    Next
    Set ls = fileList
End Function
    

' loads app. It can also open a document.  Loads asynchronously so it won't wait till app finishes.
' launchApp('d:\abc.txt')  'opens abc.txt using notepad
Public Sub launchApp(path As String)
    Dim result
    'Execute a file
    On Error GoTo errorHandler
    result = Shell("rundll32.exe url.dll,FileProtocolHandler " & (path))
    Exit Sub
errorHandler:
    MsgBox Err.DESCRIPTION, vbExclamation, "Error Loading " & path
End Sub


'=============================================================================
'FileSystemObject Implementation
'=============================================================================
#If USE_FILE_SYSTEM_OBJECT Then



#Else
'=============================================================================
' non-FileSystemObject Implementation
'=============================================================================
Public Function stripFileExtension(ByVal filename As String) As String
    Dim i
    For i = Len(filename) To 1 Step -1
        If Mid(filename, i, 1) = "." Then
            Exit For
        ElseIf isPathSeparatorChar(Mid(filename, i, 1)) Then
            i = 1
            Exit For
        End If
    Next
    If i <= 1 Then
        stripFileExtension = filename ' no extension found
    Else
        stripFileExtension = Left(filename, i - 1)
    End If
End Function

Public Function replaceFileExtension(ByVal filename As String, _
    ByVal newExtension As String) As String
    Dim newName As String
    newName = stripFileExtension(filename)
    If Left(newExtension, 1) <> "." Then    'prepend "." to Ext.
        newExtension = "." & newExtension
    End If
    replaceFileExtension = newName & newExtension
End Function

' ex: hasExtension("myFile.txt",".txt") => true
Public Function hasExtension(ByVal filename As String, ByVal extension As String)
    Dim fileNameExtension As String  'extracted ext of file
    If Left(extension, 1) <> "." Then   'prepend "." to ext
        extension = "." & extension
    End If
    extension = LCase(extension)
    fileNameExtension = getExtension(filename)
    fileNameExtension = LCase(fileNameExtension)
    If extension = fileNameExtension Then
        hasExtension = True
    Else
        hasExtension = False
    End If
End Function

Public Function getExtension(ByVal filename)
    Dim i
    i = InStrRev(filename, ".")
    If i = 0 Then
        getExtension = ""
    Else
        getExtension = Right(filename, Len(filename) - i)
        getExtension = "." + getExtension
        If InStr(getExtension, PATH_SEPARATOR_CHAR) Then
            getExtension = ""
            Exit Function
        End If
    End If
End Function
Public Function isPathSeparatorChar(ByVal c) As Boolean
    Select Case c
    'Case ":"
    '    isPathSeparatorChar = True
    Case "/"
        isPathSeparatorChar = True
    Case "\"
        isPathSeparatorChar = True
    Case Else
        isPathSeparatorChar = False
    End Select
    
End Function

Public Function addTrailingPathSeparator(ByVal path As String) As String
    If Not isPathSeparatorChar(Right$(path, 1)) Then
        path = path & PATH_SEPARATOR_CHAR
    End If
    addTrailingPathSeparator = path
End Function

Public Function removeInitialPathSeparator(ByVal path As String) As String
    If isPathSeparatorChar(Left$(path, 1)) Then
        path = Right$(path, Len(path) - 1)
    End If
    removeInitialPathSeparator = path
End Function

Public Function joinPath(ByVal path1 As String, ByVal path2 As String) As String
    path1 = addTrailingPathSeparator(path1)
    path2 = removeInitialPathSeparator(path2)
    joinPath = path1 & path2
End Function

#End If     ' NON-FileSystemObject Implementation

'=============================================================================
'=============================================================================

'=============================================================================
' OLD_CODE
' These codes were later replaced by the FileSystemObject Scripting Library
'   implementation.
'=============================================================================


    




