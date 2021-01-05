Attribute VB_Name = "FolderSync"
Dim rootFolderPath As String
Dim rootFolder As Outlook.folder
Dim duplicateRootFolderPath As String

Public Sub Start()
  Dim folderSource As Outlook.MAPIFolder
  Dim folderCompareTo As Outlook.MAPIFolder
  Dim EditSubfoldersOnly As Boolean

  'Select start folder
  Set folderSource = Application.Session.PickFolder
  Set folderCompareTo = Application.Session.PickFolder
  
  If Not folderSource Is Nothing And Not folderCompareTo Is Nothing Then
  
    Debug.Print "Started at " & Now
    
    CompareFolders folderSource, folderCompareTo
  
  End If
  
  Debug.Print "Finished at " & Now
  
End Sub

Private Sub DoFolderActions(folder As Outlook.MAPIFolder)

  Dim duplicateTargetFolderPath As String
  Dim duplicateTagertFolder As Outlook.folder
  
  duplicateTargetFolderPath = Replace(folder.FolderPath, rootFolderPath, duplicateRootFolderPath)
  CreateFolder (duplicateTargetFolderPath)
  Set duplicateTagertFolder = GetFolder(duplicateTargetFolderPath)
  RemoveDuplicateItems folder, duplicateTagertFolder
    
End Sub

Function CalculateItemKey(objItem As Object) As String
    
    If (objItem Is Nothing) Then
        CalculateItemKey = ""
        Exit Function
    End If

    Select Case True
       'Check email subject, body and sent time
       Case TypeOf objItem Is Outlook.MailItem
         Dim currentMailItem As Outlook.MailItem
         Set currentMailItem = objItem
         strKey = "MailItem" & currentMailItem.Subject & "," & Left(currentMailItem.Body, 250) & "," & currentMailItem.To & "," & currentMailItem.CC & "," & currentMailItem.BCC & "," & currentMailItem.SenderEmailAddress & "," & currentMailItem.SentOn
       'Check appointment subject, start time, duration, location and body
       Case TypeOf objItem Is Outlook.MeetingItem
        strKey = "MeetingItem" & objItem.Subject & "," & objItem.Body & "," & objItem.SentOn
       Case TypeOf objItem Is Outlook.ReportItem
        strKey = "ReportItem" & objItem.Subject & "," & objItem.Body
       Case TypeOf objItem Is Outlook.AppointmentItem
         strKey = "AppointmentItem" & objItem.Subject & "," & objItem.Start & "," & objItem.Duration & "," & objItem.Location & "," & objItem.Body
       'Check contact full name and email address
       Case TypeOf objItem Is Outlook.ContactItem
         strKey = "ContactItem" & objItem.FullName & "," & objItem.Email1Address & "," & objItem.Email2Address & "," & objItem.Email3Address
       'Check task subject, start date, due date and body
       Case TypeOf objItem Is Outlook.TaskItem
         strKey = "TaskItem" & objItem.Subject & "," & objItem.StartDate & "," & objItem.DueDate & "," & objItem.Body
    End Select

    If strKey = "" Then
        Debug.Print "Error: Found an unrecognized item type"
        CalculateItemKey = ""
        Exit Function
    End If
    
    strKey = Replace(strKey, ", ", Chr(32))

    CalculateItemKey = strKey
            
End Function


Function CompareFolders(folderLeft As Outlook.folder, folderRight As Outlook.folder)
    Dim leftDictionary As Object
    Dim i As Long
    Dim totalDuplicatesDetected As Long
    Dim objItem As Object
    Dim strKey As String
    
    Set leftDictionary = CreateObject("scripting.dictionary")
    
    If (folderLeft Is Nothing Or folderRight Is Nothing) Then
        Exit Function
    End If
    
    Dim leftFolderItems As Outlook.Items
    Set leftFolderItems = folderLeft.Items
    Set rightFolderItems = folderRight.Items
    
    If (folderLeft.DefaultItemType = olMailItem And folderRight.DefaultItemType = olMailItem) Then
        leftFolderItems.Sort "[ReceivedTime][Subject]", True
        rightFolderItems.Sort "[ReceivedTime][Subject]", True
    End If
    
    Debug.Print Now & " | Reading left folder: " & folderLeft.FolderPath
    Debug.Print Now & " | Items to process: " & leftFolderItems.Count
    
    For i = leftFolderItems.Count To 1 Step -1
        Set objItem = leftFolderItems.item(i)
        strKey = CalculateItemKey(objItem)
        
        If i Mod 1000 = 0 Then
            Debug.Print Now & " | Items to process: " & i
        End If

        If Not strKey = "" Then
            If leftDictionary.Exists(strKey) = False Then
                leftDictionary.Add strKey, objItem
            End If
        Else
            Debug.Print "Error: Found an unrecognized item type"
        End If
        
        DoEvents
    Next i
    
    Debug.Print Now & " | Reading right folder: " & folderRight.FolderPath
    Debug.Print Now & " | Items to process: " & rightFolderItems.Count
    
    For i = rightFolderItems.Count To 1 Step -1
        Set objItem = rightFolderItems.item(i)
        strKey = CalculateItemKey(objItem)
        
        If i Mod 1000 = 0 Then
            Debug.Print Now & " | Items to process: " & i
        End If

        If Not strKey = "" Then
        
          'Remove the duplicate items
          If leftDictionary.Exists(strKey) = False Then
          
            Dim copyTarget As Outlook.folder
            Set copyTarget = GetFolder(folderLeft.FolderPath & "\## MISSING ##")
            
            If copyTarget Is Nothing Then
                Set copyTarget = folderLeft.Folders.Add("## MISSING ##")
            End If
          
            Dim missingItem As Object
            Set missingItem = objItem.Copy
            missingItem.Move copyTarget
            totalDuplicatesDetected = totalDuplicatesDetected + 1
          End If
          
        Else
            Debug.Print "Error: Found an unrecognized item type"
        End If
        
        DoEvents
    Next i
    
    Debug.Print "Found " & totalDuplicatesDetected & " missing item(s)"
    
End Function


Function GetFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    On Error GoTo 0
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = Nothing
            
            On Error Resume Next
            Set TestFolder = SubFolders.item(FoldersArray(i))
            On Error GoTo 0
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
                Exit Function
            End If
        Next
    End If
     
   'Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function
 
GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function
End Function


Function CreateFolder(ByVal FolderPath As String) As Outlook.folder
    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer
 
    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If
    
    'Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = Nothing
            
            On Error Resume Next
            Set TestFolder = SubFolders.item(FoldersArray(i))
            If TestFolder Is Nothing Then
                SubFolders.Add (FoldersArray(i))
                Set TestFolder = SubFolders.item(FoldersArray(i))
            End If
        Next
    End If
     
    Exit Function
 
GetFolder_Error:
    Exit Function
End Function

