' =========================================================================
' COKER'S OUTLOOK MINER (2026)
' Architected by Steven James Coker. Dynamic text extraction engine and 
' automated file routing developed in collaboration with Google's AI.
'
' LICENSE (DONATIONWARE):
' This macro is free for non-commercial use. You may freely copy and
' distribute it. For commercial use inquiries, please contact the author.
'
' SUPPORT THE WORK:
' If this tool helps streamline your research, please consider a donation:
'   - PayPal:   paypal.com/paypalme/SJCoker
'   - GoFundMe: gofundme.com/f/genetic-genealogy
' =========================================================================

' =========================================================================
' WINDOWS API DECLARATIONS FOR INI MANAGEMENT (MUST BE AT VERY TOP OF MODULE)
' NOTE: If you are on a modern 64-bit system, the lines under #Else may appear 
' RED in your VBA editor. This is completely normal and expected! The compiler 
' will automatically ignore those red fallback lines when you run the macro.
' =========================================================================
#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

' =========================================================================
' MACRO: ATTACHMENTS_EMAIL (UNIVERSAL SYNC PIPELINE)
' =========================================================================
Sub Attachments_Email()
    Dim olNs As Outlook.NameSpace
    Dim olStartFolder As Outlook.MAPIFolder
    Dim fso As Object, txtStream As Object
    Dim sRootPath As String, sOldRoot As String, sOldRouteIni As String
    Dim sBaseFolder As String, sAttRoot As String, sTextRoot As String
    Dim sSyncIni As String, sRouteIni As String
    Dim sModeInput As String, sDepthLabel As String, sThresholdInput As String
    Dim sLastSync As String, sProfileName As String, sRouteChoice As String
    Dim sRouteName As String, sRouteDest As String, sSummary As String
    Dim dtCutoff As Date
    Dim iMode As Integer, maxDepth As Integer
    Dim dThresholdMB As Double
    Dim lThresholdBytes As Long
    Dim sortResponse As VbMsgBoxResult
    Dim bNewestFirst As Boolean, bMoreRoutes As Boolean
    Dim emailCount As Long, attCount As Long

    emailCount = 0: attCount = 0
    lThresholdBytes = 0
    On Error GoTo ErrorHandler

    Set olNs = Application.GetNamespace("MAPI")
    Set olStartFolder = olNs.PickFolder
    If olStartFolder Is Nothing Then Exit Sub
    
    ' 1. ROOT PATH PROMPT (With Registry Memory)
    sOldRoot = GetSetting("ExtractsMacro", "Settings", "RootPath", "S:\")
    sRootPath = InputBox("Enter the ROOT drive/folder for your STANDARD extracts." & vbCrLf & _
                         "We will use/create a static 'Extracts' folder inside it.", _
                         "Enter Root Location", sOldRoot)
    If StrPtr(sRootPath) = 0 Then Exit Sub ' INSTANT CANCEL TRIGGER
    sRootPath = Trim(sRootPath)
    If sRootPath = "" Then Exit Sub
    If Right(sRootPath, 1) <> "\" Then sRootPath = sRootPath & "\"
    
    ' Save the confirmed root path back to memory
    SaveSetting "ExtractsMacro", "Settings", "RootPath", sRootPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    sBaseFolder = sRootPath & "Extracts\"
    If Not fso.FolderExists(sBaseFolder) Then fso.CreateFolder sBaseFolder
    
    sSyncIni = sBaseFolder & "Last_Sync.ini"
    sRouteIni = sBaseFolder & "Custom_Routing.ini"

    ' --- MIGRATION LOGIC FOR NEW ROOT PATHS ---
    If sOldRoot <> "" And UCase(sOldRoot) <> UCase(sRootPath) Then
        sOldRouteIni = sOldRoot & "Extracts\Custom_Routing.ini"
        If fso.FileExists(sOldRouteIni) And Not fso.FileExists(sRouteIni) Then
            If MsgBox("You changed your Root Path from:" & vbCrLf & sOldRoot & vbCrLf & vbCrLf & _
                      "Do you want to copy your custom routing rules to this new location?", _
                      vbYesNo + vbQuestion, "Import Routing Rules?") = vbYes Then
                fso.CopyFile sOldRouteIni, sRouteIni
            End If
        End If
    End If
    ' ------------------------------------------

    ' 2. TRAFFIC CONTROLLER CONFIGURATION MENU
    sRouteChoice = InputBox("CUSTOM FOLDER ROUTING" & vbCrLf & _
                            "How would you like to handle custom routing locations?" & vbCrLf & vbCrLf & _
                            "1 = Guided Setup (Prompt me step-by-step)" & vbCrLf & _
                            "2 = Manual Setup (Open INI in Notepad)" & vbCrLf & _
                            "3 = Skip (Proceed with current settings)", "Traffic Controller Setup", "3")
    If StrPtr(sRouteChoice) = 0 Then Exit Sub
    
    Select Case Trim(sRouteChoice)
        Case "1"
            bMoreRoutes = True
            Do While bMoreRoutes
                sRouteName = InputBox("Current Custom Routes:" & vbCrLf & GetActiveRoutes(sRouteIni, fso) & vbCrLf & _
                                      "Type the exact name of the Outlook folder (e.g., !CGMS):", "Add Routing Rule")
                If StrPtr(sRouteName) = 0 Then Exit Sub
                If Trim(sRouteName) <> "" Then
                    sRouteDest = InputBox("Paste the full destination path for '" & sRouteName & "':" & vbCrLf & _
                                          "(e.g., D:\Dropbox\CGMS\)", "Set Destination")
                    If StrPtr(sRouteDest) = 0 Then Exit Sub
                    If Trim(sRouteDest) <> "" Then WriteINI "Traffic_Controller", Trim(sRouteName), Trim(sRouteDest), sRouteIni
                End If
                If MsgBox("Route Saved!" & vbCrLf & vbCrLf & "Do you want to add another custom location?", vbYesNo + vbQuestion, "Add Another?") = vbNo Then bMoreRoutes = False
            Loop
        Case "2"
            If Not fso.FileExists(sRouteIni) Then
                Set txtStream = fso.CreateTextFile(sRouteIni, True)
                txtStream.WriteLine "[Traffic_Controller]"
                txtStream.WriteLine "; ========================================================================="
                txtStream.WriteLine "; CUSTOM FOLDER ROUTING RULES"
                txtStream.WriteLine "; ========================================================================="
                txtStream.WriteLine "; To route an Outlook folder to a specific location, type the folder name,"
                txtStream.WriteLine "; an equals sign, and the full destination path."
                txtStream.WriteLine ";"
                txtStream.WriteLine "; Format Example:   FolderName=Drive:\Folder\Path\"
                txtStream.WriteLine ";"
                txtStream.WriteLine "; Note: The macro will automatically fix missing ending backslashes."
                txtStream.WriteLine "; ========================================================================="
                txtStream.Close
            End If
            Shell "notepad.exe """ & sRouteIni & """", vbNormalFocus
            MsgBox "Please make your edits in Notepad." & vbCrLf & vbCrLf & _
                   "Save the file, close Notepad, and then click OK here to continue.", vbInformation, "Manual INI Edit"
        Case "3"
            MsgBox "Current Active Routes:" & vbCrLf & vbCrLf & GetActiveRoutes(sRouteIni, fso), vbInformation, "Current Settings"
    End Select

    ' 3. EXACT DEPTH MENU PROMPT (Optimized Width)
    sModeInput = InputBox("Select Text Extraction Depth (Folder Level):" & vbCrLf & vbCrLf & _
                          "1 = Depth 1: Root Mailbox Only (1 Master Sequence)" & vbCrLf & _
                          "2 = Depth 2: Top-Level Folders (Root\Dir1)" & vbCrLf & _
                          "3 = Depth 3: Sub-Folders (Root\Dir1\Dir2)" & vbCrLf & _
                          "4 = Depth 4: 3rd-Level (Root\Dir1\Dir2\Dir3)" & vbCrLf & _
                          "5 = Depth 5: 4th-Level (Root\Dir1\...\Dir4)" & vbCrLf & _
                          "6 = Depth All: Granular (Each folder gets a txt file)", "Select Mode", "2")
    If StrPtr(sModeInput) = 0 Then Exit Sub
    iMode = Val(sModeInput)
    If iMode < 1 Or iMode > 6 Then Exit Sub

    Select Case iMode
        Case 1: maxDepth = 1: sDepthLabel = "1"
        Case 2: maxDepth = 2: sDepthLabel = "2"
        Case 3: maxDepth = 3: sDepthLabel = "3"
        Case 4: maxDepth = 4: sDepthLabel = "4"
        Case 5: maxDepth = 5: sDepthLabel = "5"
        Case 6: maxDepth = 999: sDepthLabel = "All"
    End Select

    ' 4. UNIVERSAL CHUNK DIAL (Floating Point Enabled)
    sThresholdInput = InputBox("UNIVERSAL CHUNKING LIMIT (IN MEGABYTES):" & vbCrLf & vbCrLf & _
                               "Enter the max file size for text chunks (e.g., 2.8)." & vbCrLf & _
                               "Leave blank or enter 0 to disable chunking.", _
                               "Set MB Chunk Limit", "3")
    If StrPtr(sThresholdInput) = 0 Then Exit Sub
    dThresholdMB = Val(sThresholdInput)
    If dThresholdMB > 0 Then
        lThresholdBytes = CLng(dThresholdMB * 1048576)
        ' Safe folder naming (replaces decimal point with a dash)
        sDepthLabel = sDepthLabel & "_" & Replace(CStr(dThresholdMB), ".", "-") & "MB" 
    Else
        lThresholdBytes = 0
        sDepthLabel = sDepthLabel & "_0MB"
    End If

    ' 5. PROFILE-SPECIFIC INI READ & DATE CUTOFF
    sProfileName = "Depth_" & sDepthLabel
    sLastSync = ReadINI(sProfileName, "LastSync", "2000-01-01 00:00:00", sSyncIni)

    sLastSync = InputBox("INCREMENTAL SYNC DATE FOR: " & sProfileName & vbCrLf & vbCrLf & _
                         "The date below was auto-populated from your last run of this specific setup." & vbCrLf & _
                         "Attachments OLDER than this date will be skipped." & vbCrLf & _
                         "NOTE: Text files ALWAYS do a 100% full rewrite." & vbCrLf & vbCrLf & _
                         "OVERRIDE: Leave blank or type '0' to force extract ALL attachments.", _
                         "Set Cutoff Date", sLastSync)
    If StrPtr(sLastSync) = 0 Then Exit Sub
    
    sLastSync = Trim(sLastSync)
    If sLastSync = "" Or sLastSync = "0" Then
        dtCutoff = CDate("1900-01-01")
    ElseIf Not IsDate(sLastSync) Then
        MsgBox "Invalid date format. Cancelling.", vbCritical
        Exit Sub
    Else
        dtCutoff = CDate(sLastSync)
    End If

    sortResponse = MsgBox("Sort emails Newest to Oldest?", vbYesNo + vbQuestion, "Sort Order")
    bNewestFirst = (sortResponse = vbYes)

    ' 6. THE PRE-FLIGHT SUMMARY (With Attribution Footer)
    sSummary = "READY TO EXECUTE: Pre-Flight Summary" & vbCrLf & vbCrLf & _
               "Please review your settings before starting the extraction:" & vbCrLf & vbCrLf & _
               "Root Destination: " & sRootPath & vbCrLf & _
               "Extraction Depth: " & iMode & vbCrLf & _
               "Chunking Limit: " & IIf(dThresholdMB > 0, dThresholdMB & " MB", "Disabled") & vbCrLf & _
               "Incremental Cutoff: " & IIf(Year(dtCutoff) <= 1900, "Full Rewrite (No Cutoff)", Format(dtCutoff, "yyyy-mm-dd hh:mm:ss")) & vbCrLf & _
               "Sort Order: " & IIf(bNewestFirst, "Newest to Oldest", "Oldest to Newest") & vbCrLf & vbCrLf & _
               "Active Routing Rules:" & vbCrLf & GetActiveRoutes(sRouteIni, fso) & vbCrLf & _
               "System Files Location:" & vbCrLf & _
               "Sync Dates: " & sSyncIni & vbCrLf & _
               "Routing Rules: " & sRouteIni & vbCrLf & vbCrLf & _
               "Click OK to execute the extraction, or Cancel to safely abort." & vbCrLf & vbCrLf & _
               "------------------------------------------------" & vbCrLf & _
               "Attachments & Email Extractor (2026)" & vbCrLf & _
               "By Steven James Coker & Google's AI" & vbCrLf & _
               "Donationware: See code header for support links."
               
    If MsgBox(sSummary, vbOKCancel + vbInformation, "Pre-Flight Summary") = vbCancel Then Exit Sub

    ' --- DIRECTORY ARCHITECTURE BUILD ---
    sAttRoot = sBaseFolder & "Attachments\"
    If Not fso.FolderExists(sAttRoot) Then fso.CreateFolder sAttRoot
    
    If Not fso.FolderExists(sBaseFolder & "Emails\") Then fso.CreateFolder sBaseFolder & "Emails\"
    sTextRoot = sBaseFolder & "Emails\Depth_" & sDepthLabel & "\"
    If Not fso.FolderExists(sTextRoot) Then fso.CreateFolder sTextRoot

    ' --- LAUNCH DYNAMIC ENGINE ---
    Process_Folder_Tree olStartFolder, sTextRoot, sAttRoot, bNewestFirst, fso, emailCount, attCount, dtCutoff, 1, maxDepth, lThresholdBytes, sRouteIni, , 1, "", "", "", False

    ' Write Timestamp to Machine INI
    WriteINI sProfileName, "LastSync", Format(Now, "yyyy-mm-dd hh:mm:ss"), sSyncIni

    MsgBox "Success! Emails: " & emailCount & vbCrLf & "Attachments: " & attCount & vbCrLf & _
           "Text Saved To: Emails\Depth_" & sDepthLabel, vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' =========================================================================
' INI FILE MANAGEMENT & UTILITIES
' =========================================================================
Function ReadINI(sSection As String, sKey As String, sDefault As String, sIniFile As String) As String
    Dim sRet As String * 255, lLen As Long
    lLen = GetPrivateProfileString(sSection, sKey, sDefault, sRet, 255, sIniFile)
    ReadINI = Left(sRet, lLen)
End Function

Sub WriteINI(sSection As String, sKey As String, sValue As String, sIniFile As String)
    WritePrivateProfileString sSection, sKey, sValue, sIniFile
End Sub

Function GetActiveRoutes(sIniPath As String, fso As Object) As String
    Dim txtStream As Object, sLine As String, sResult As String
    sResult = ""
    If fso.FileExists(sIniPath) Then
        Set txtStream = fso.OpenTextFile(sIniPath, 1)
        Do While Not txtStream.AtEndOfStream
            sLine = Trim(txtStream.ReadLine)
            If sLine <> "" And Left(sLine, 1) <> ";" And Left(sLine, 1) <> "[" Then
                sResult = sResult & "- " & Replace(sLine, "=", " -> ") & vbCrLf
            End If
        Loop
        txtStream.Close
    End If
    If sResult = "" Then sResult = "- None" & vbCrLf
    GetActiveRoutes = sResult
End Function

Function GetCustomLocation(sFolderName As String, sRouteIni As String) As String
    Dim sResult As String
    sResult = Trim(ReadINI("Traffic_Controller", sFolderName, "", sRouteIni))
    ' Backslash Fail-Safe Auto-Corrector
    If sResult <> "" And Right(sResult, 1) <> "\" Then sResult = sResult & "\"
    GetCustomLocation = sResult
End Function

' =========================================================================
' DYNAMIC DEPTH ENGINE WITH PRIVACY ISOLATION & SCOPE PROTECTION
' =========================================================================
Private Sub Process_Folder_Tree(olFolder As Outlook.MAPIFolder, sTextRoot As String, sCurrentAttPath As String, _
                                bSort As Boolean, fso As Object, ByRef eCount As Long, ByRef aCount As Long, _
                                dtCutoff As Date, currentDepth As Integer, ByVal maxDepth As Integer, _
                                ByVal lThresholdBytes As Long, ByVal sRouteIni As String, _
                                Optional ByRef parentStream As Object = Nothing, _
                                Optional ByRef chunkIndex As Integer = 1, _
                                Optional ByVal baseLogName As String = "", _
                                Optional ByVal activeLogFolder As String = "", _
                                Optional ByVal sRelativePath As String = "", _
                                Optional ByVal bIsPrivateScope As Boolean = False)
    
    If Not HasItems(olFolder) Then Exit Sub

    Dim sAttPath As String, utfStream As Object, olSub As Outlook.MAPIFolder
    Dim sFolderName As String, sCustom As String
    Dim bOwnStream As Boolean, bInitPrivate As Boolean
    Dim localChunkIndex As Integer

    sFolderName = SanitizeFileName(olFolder.Name)
    sCustom = GetCustomLocation(olFolder.Name, sRouteIni)
    
    If currentDepth > 1 Then sRelativePath = sRelativePath & sFolderName & "\"
    
    If sCustom <> "" Then
        sAttPath = sCustom
    ElseIf currentDepth = 1 Then
        sAttPath = sCurrentAttPath
    Else
        sAttPath = sCurrentAttPath & sFolderName & "\"
    End If
    
    bInitPrivate = False
    If Left(sFolderName, 1) = "_" And Not bIsPrivateScope Then
        bInitPrivate = True
        bIsPrivateScope = True 
    End If

    If currentDepth <= maxDepth Or bInitPrivate Then
        bOwnStream = True
        localChunkIndex = 1 
        
        If currentDepth = 1 And maxDepth > 1 Then 
            baseLogName = "Root_Items"
            activeLogFolder = sTextRoot & "Root_Items\"
        Else 
            baseLogName = sFolderName
            activeLogFolder = sTextRoot & sRelativePath
        End If
        
        If Not fso.FolderExists(activeLogFolder) Then BuildFolderTree activeLogFolder, fso
        
        Set utfStream = CreateObject("ADODB.Stream")
        utfStream.Type = 2: utfStream.Charset = "utf-8": utfStream.Open
        utfStream.WriteText "Export Date: " & Now & vbCrLf & "Outlook Path: " & olFolder.FolderPath & vbCrLf & "------------------" & vbCrLf
        
        If olFolder.Items.Count > 0 Then
            utfStream.WriteText vbCrLf & "===== FOLDER PATH: " & olFolder.FolderPath & " =====" & vbCrLf
            Process_Items_Only olFolder, utfStream, sAttPath, eCount, aCount, bSort, fso, dtCutoff, lThresholdBytes, localChunkIndex, baseLogName, activeLogFolder
        End If

        For Each olSub In olFolder.Folders
            If Not IsJunkFolder(olSub.Name) Then 
                Process_Folder_Tree olSub, sTextRoot, sAttPath, bSort, fso, eCount, aCount, dtCutoff, currentDepth + 1, maxDepth, lThresholdBytes, sRouteIni, utfStream, localChunkIndex, baseLogName, activeLogFolder, sRelativePath, bIsPrivateScope
            End If
        Next
        
        Dim finalLogName As String
        If lThresholdBytes > 0 Then
            finalLogName = GetUniqueLogFileName(activeLogFolder, baseLogName & "_Part" & Format(localChunkIndex, "000") & "-" & Format(Date, "yyyy-mm-dd"), fso)
        Else
            finalLogName = GetUniqueLogFileName(activeLogFolder, baseLogName & "-" & Format(Date, "yyyy-mm-dd"), fso)
        End If
        
        If Not utfStream Is Nothing Then
            If utfStream.State = 1 Then 
                utfStream.SaveToFile finalLogName, 2
                utfStream.Close
            End If
        End If

    Else
        bOwnStream = False
        Set utfStream = parentStream
        
        If olFolder.Items.Count > 0 Then
            utfStream.WriteText vbCrLf & "===== FOLDER PATH: " & olFolder.FolderPath & " =====" & vbCrLf
            Process_Items_Only olFolder, utfStream, sAttPath, eCount, aCount, bSort, fso, dtCutoff, lThresholdBytes, chunkIndex, baseLogName, activeLogFolder
        End If

        For Each olSub In olFolder.Folders
            If Not IsJunkFolder(olSub.Name) Then 
                Process_Folder_Tree olSub, sTextRoot, sAttPath, bSort, fso, eCount, aCount, dtCutoff, currentDepth + 1, maxDepth, lThresholdBytes, sRouteIni, utfStream, chunkIndex, baseLogName, activeLogFolder, sRelativePath, bIsPrivateScope
            End If
        Next
        
        Set parentStream = utfStream 
    End If
End Sub

Private Sub Process_Items_Only(olFolder As Outlook.MAPIFolder, ByRef streamObj As Object, sAttPath As String, ByRef eCount As Long, ByRef aCount As Long, bSort As Boolean, fso As Object, dtCutoff As Date, lThresholdBytes As Long, ByRef currentChunkIndex As Integer, ByVal baseLogName As String, ByVal activeLogFolder As String)
    Dim olItem As Object, colItems As Outlook.Items, olAtt As Outlook.Attachment
    Dim sDatePrefix As String, sSafeName As String, itemDate As Date
    Dim sSubject As String, sFrom As String, sSent As String, sTo As String

    Set colItems = olFolder.Items
    On Error Resume Next
    colItems.Sort "[SentOn]", bSort
    On Error GoTo 0
    
    For Each olItem In colItems
        On Error Resume Next
        sFrom = "Unknown": sSent = "Unknown": sTo = "": sSubject = "Unknown"
        itemDate = 0
        
        sFrom = olItem.SenderEmailAddress: If sFrom = "" Then sFrom = olItem.SenderName
        sSent = CStr(olItem.SentOn): sTo = olItem.To: sSubject = olItem.Subject
        
        itemDate = olItem.SentOn
        If Year(itemDate) < 1900 Or Year(itemDate) > 2100 Then itemDate = olItem.LastModificationTime
        If Year(itemDate) < 1900 Or Year(itemDate) > 2100 Then itemDate = Date
        
        sDatePrefix = Format(itemDate, "yyyy-mm-dd_hhnn")
        
        If sSent = "1/1/4501" Or sSent = "12/30/1899" Then sSent = "None"
        
        If Err.Number = 0 Then
            eCount = eCount + 1
            
            If lThresholdBytes > 0 Then
                If streamObj.Size >= lThresholdBytes Then
                    Dim sSavePath As String
                    sSavePath = GetUniqueLogFileName(activeLogFolder, baseLogName & "_Part" & Format(currentChunkIndex, "000") & "-" & Format(Date, "yyyy-mm-dd"), fso)
                    streamObj.SaveToFile sSavePath, 2
                    streamObj.Close
                    
                    currentChunkIndex = currentChunkIndex + 1
                    
                    Set streamObj = CreateObject("ADODB.Stream")
                    streamObj.Type = 2: streamObj.Charset = "utf-8": streamObj.Open
                    streamObj.WriteText "Export Date: " & Now & vbCrLf & "Continued from Part " & Format(currentChunkIndex - 1, "000") & vbCrLf & "------------------" & vbCrLf
                End If
            End If
            
            streamObj.WriteText "From: " & sFrom & vbCrLf & "Sent: " & sSent & vbCrLf & "Subject: " & sSubject & vbCrLf
            
            If itemDate >= dtCutoff Then
                If olItem.Attachments.Count > 0 Then
                     For Each olAtt In olItem.Attachments
                        If olAtt.Type = 1 Then 
                            sSafeName = sDatePrefix & "_" & SanitizeFileName(sSubject) & "_" & SanitizeFileName(olAtt.FileName)
                            If Len(sSafeName) > 120 Then sSafeName = Left(sSafeName, 110) & Right(olAtt.FileName, 6)
                            If Not fso.FolderExists(sAttPath) Then BuildFolderTree sAttPath, fso
                            
                            ' === GHOST ATTACHMENT FIX ===
                            If Not fso.FileExists(sAttPath & sSafeName) Then
                                Err.Clear
                                olAtt.SaveAsFile sAttPath & sSafeName
                                If Err.Number = 0 Then
                                    streamObj.WriteText "   -> Saved Attachment: " & sSafeName & vbCrLf
                                    aCount = aCount + 1
                                Else
                                    streamObj.WriteText "   -> FAILED TO SAVE (Skipped): " & sSafeName & " [Error: " & Err.Description & "]" & vbCrLf
                                End If
                            Else
                                streamObj.WriteText "   -> Skipped (Exists): " & sSafeName & vbCrLf
                            End If
                            ' ==============================
                        End If
                        Err.Clear
                    Next olAtt
                End If
            End If
            streamObj.WriteText vbCrLf & "--------------------------------------------------" & vbCrLf
            streamObj.WriteText olItem.Body & vbCrLf
            streamObj.WriteText vbCrLf & "--------------------------------------------------" & vbCrLf
        Else
            Err.Clear
        End If
        On Error GoTo 0
        DoEvents
    Next olItem
End Sub

' =========================================================================
' UTILITIES, FILTERS, & SCOUTS
' =========================================================================
Function HasItems(fld As Outlook.MAPIFolder) As Boolean
    Dim subFld As Outlook.MAPIFolder
    On Error Resume Next
    If fld.Items.Count > 0 Then
        HasItems = True
        Exit Function
    End If
    For Each subFld In fld.Folders
        If Not IsJunkFolder(subFld.Name) Then
            If HasItems(subFld) Then
                HasItems = True
                Exit Function
            End If
        End If
    Next
    HasItems = False
End Function

Function IsJunkFolder(sName As String) As Boolean
    Dim sLower As String: sLower = LCase(Trim(sName))
    Select Case sLower
        Case "junk e-mail", "junk email", "spam", "deleted items", "trash", "sync issues", "conflicts"
            IsJunkFolder = True
        Case Else
            IsJunkFolder = False
    End Select
End Function

Function SanitizeFileName(sName As String) As String
    Dim invalidChars As Variant, i As Integer, tempName As String
    tempName = sName
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    For i = LBound(invalidChars) To UBound(invalidChars)
        tempName = Replace(tempName, invalidChars(i), "-")
    Next i
    SanitizeFileName = Trim(tempName)
End Function

Function GetUniqueLogFileName(sBasePath As String, sFolderName As String, fso As Object) As String
    Dim sTestPath As String, iCounter As Integer
    sTestPath = sBasePath & SanitizeFileName(sFolderName) & ".txt"
    iCounter = 1
    Do While fso.FileExists(sTestPath)
        sTestPath = sBasePath & SanitizeFileName(sFolderName) & "_" & iCounter & ".txt"
        iCounter = iCounter + 1
    Loop
    GetUniqueLogFileName = sTestPath
End Function

Sub BuildFolderTree(ByVal sPath As String, fso As Object)
    Dim sParent As String
    If sPath = "" Then Exit Sub
    If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
    If Not fso.FolderExists(sPath) Then
        sParent = fso.GetParentFolderName(sPath)
        If Not fso.FolderExists(sParent) And sParent <> "" Then BuildFolderTree sParent, fso
        fso.CreateFolder sPath
    End If

End Sub
