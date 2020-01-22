' ------------------------------------------------
' ExtractCode.vbs
' Derived from original code from Chris Hemedinger (cjdinger)
' https://github.com/cjdinger/sas-eg-automation/blob/master/vbscript/ExtractCode.vbs
'
' This program extracts code from EGP file quite differently from original code:
' Uses SAS Enterprise Guide automation to read an EGP file
' and export all SAS programs to ONE SAS file
'
' This code is customized to work with an EGP file where the user has mannualy
' entered numbers as prefixes to the labels of all the projects nodes (defined
' in node.Name) so that, for example when the nodes are added to a EG Ordered
' List, a proper execution order is found by simply sorting the nodes by the
' label node.Name.  That ordering is detected by this code; its extracted and
' sorted, but only nodes with numbers are output to the new file.
'
' This script uses the Scripting.FileSystemObject to save the code,
' _which_WILL_NOT_ include
' other "wrapper" code around each program, including macro
' variables and definitions that are used by SAS Enterprise Guide
' An exception to this is that the _CLIENTTASKLABEL is included to recognize
' the name of the node (perhaps slightly altered) and is given without quotes.
'
' USAGE:
'   Add code to and Excel Visual Basic Application editor
'   Check boxes for object files in Tool-->References
'   Change variables: EGPfile, programsFolder
'   Change variable egVersion when SAS upgrades its SASEGObjectModel version
'
' Required References:
'  * SAS: Integrated Object Model (IOM)
'  * Microsoft Scripting Runtime

Option Explicit
Dim Application ' As SASEGuide.Application

Sub EGPConvert_v8()
  Dim Project
  
  ' Change if running a different version of EG
  Dim egVersion
  egVersion = "SASEGObjectModel.Application.8.1"
  
  ' enumeration of project item types
  Const egLog = 0
  Const egCode = 1
  Const egData = 2
  Const egQuery = 3
  Const egContainer = 4
  Const egDocBuilder = 5
  Const egNote = 6
  Const egResult = 7
  Const egTask = 8
  Const egTaskCode = 9
  Const egProjectParameter = 10
  Const egOutputData = 11
  Const egStoredProcess = 12
  Const egStoredProcessParameter = 13
  Const egPublishAction = 14
  Const egCube = 15
  Const egReport = 18
  Const egReportSnapshot = 19
  Const egOrderedList = 20
  Const egSchedule = 21
  Const egLink = 22
  Const egFile = 23
  Const egIntrNetApp = 24
  Const egInformationMap = 25
  Dim EGPfile As String
  Dim programsFolder As String
  
  EGPfile = "C:\Users\me\Desktop\ExtractCode\fact_sheet.egp"
  programsFolder = "C:\Users\me\Desktop\ExtractCode\"
   
  ' Create a new SAS Enterprise Guide automation session
  On Error Resume Next
  Set Application = CreateObject(egVersion)
  If Err.Number <> 0 Then
    MsgBox "ERROR: Need help with 'Set Application = WScript.CreateObject(egVersion)'"
    End
  End If
   MsgBox Application.Name & ", Version: " & Application.Version & vbCrLf & "Opening project: " & EGPfile
  
  ' Open the EGP file with the Application
  Set Project = Application.Open(EGPfile, "")
  If Err.Number <> 0 Then
    MsgBox "ERROR: Unable to open " & EGPfile & " as a project file"
      End
  End If
  
 
  
  Dim item
  Dim flow
  Dim Unsorted
  Dim Items_Sorted
  Dim outputFilename
  Dim TxtOutput
  Dim i
  Dim nodeNumber
  
  MkDir programsFolder & Project.Name
    outputFilename = programsFolder & Project.Name & "\" & Project.Name & ".sas"
    MsgBox "saving code to " & vbCrLf & outputFilename
    'TxtOutput will be saved in outputFilename
    TxtOutput = ""
  
  Set Unsorted = CreateObject("System.Collections.ArrayList")
  
  ' Navigate the process flows in the Project object
  
  For Each flow In Project.ContainerCollection
    ' ProcessFlow is ContainerType of 0
    If flow.ContainerType = 0 Then
        MsgBox "Process Flow: " & flow.Name
        nodeNumber = 0
    ' Navigate the items in each process flow
    For Each item In flow.Items
      MsgBox "=Unsorted<--" & item.Name & " Type: " & item.Type
      Unsorted.Add item
    Next
    End If
  Next
    Set Items_Sorted = SortArrayList_ByName(Unsorted)
    For i = 0 To Items_Sorted.Count - 1
      Set item = Items_Sorted(i)
      'MsgBox "  " & item.Name & ", item.Type=" & Str(item.Type)

      'Only Process if item.name begins with a number
        Dim item_Name_begins_with_number As Boolean
          item_Name_begins_with_number = InStr(1, "0123456789", Left(item.Name, 1), 0) > 0
       
        If item_Name_begins_with_number Then
          Select Case item.Type
  
          Case egCode
            'MsgBox "  " & item.Name
          TxtOutput = TxtOutput & _
            "%LET _CLIENTTASKLABEL=" & strClean(item.Name) & ";" & vbCrLf & vbCrLf & item.text & vbCrLf & vbCrLf
                  nodeNumber = nodeNumber + 1
          
          Case egTask, egQuery
          Dim item_TaskCode_Is_Nothing As Boolean
            item_TaskCode_Is_Nothing = item.TaskCode Is Nothing
            MsgBox "  " & item.Name & ", Task/Query" & vbCrLf & _
                   "item_TaskCode_Is_Nothing is" & Str(item_TaskCode_Is_Nothing)
            
          If (Not item.TaskCode Is Nothing) Then
            TxtOutput = TxtOutput & _
              "%LET _CLIENTTASKLABEL=" & strClean(item.Name) & ";" & vbCrLf & vbCrLf & item.TaskCode.text & vbCrLf & vbCrLf
          End If
                  nodeNumber = nodeNumber + 1
  
          Case egData
            'MsgBox "  " & item.Name & ", Data: " & item.fileName
          'tableOfContents = tableOfContents & " (Data set)"
          Dim task
          For Each task In item.Tasks
                        MsgBox "    " & task.Name & ", sub-task: "
            If (Not task.TaskCode Is Nothing) Then
            TxtOutput = TxtOutput & _
              "%LET _CLIENTTASKLABEL=" & strClean(task.Name) & ";" & vbCrLf & vbCrLf & item.TaskCode.text & vbCrLf & vbCrLf
            End If
          Next
                    nodeNumber = nodeNumber + 1

  
          End Select
      End If
          Next
          
  MsgBox "nodeNumber=" & Str(nodeNumber)
    If nodeNumber > 0 Then _
      saveTextToFile outputFilename, TxtOutput
  
  
  ' Close the project
  Project.Close
  ' Quit/close the Application object
  Application.Quit
   
End Sub

' --- Helper functions ----------------

' function to get the current working directory
Function getWorkingDirectory()
  Dim objShell
  Set objShell = CreateObject("Wscript.Shell")
  getWorkingDirectory = objShell.CurrentDirectory
End Function

' function to create a new subfolder, if it doesn't yet exist
Function createFolder(folderName)
  Dim objFSO
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  If objFSO.FolderExists(folderName) = False Then
  objFSO.createFolder folderName
  End If
End Function

' Save a block of text (code or log) to text file
Function saveTextToFile(fileName, text)
  MsgBox "running: saveTextToFile"
  Dim objFS As FileSystemObject
  Dim objOutFile
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objOutFile = objFS.CreateTextFile(fileName, True)
  objOutFile.Write (text)
  objOutFile.Close
  IF_Saving_Error_THEN_Report Err.Number, fileName
End Function

Function IF_Saving_Error_THEN_Report(ByRef Err_Number, fileName)
  If Err.Number <> 0 Then
    MsgBox "     ERROR: There was an error while saving " & fileName
    Err.Clear
  End If
End Function

'Sort an array of object pointers that have the Name Property
Function SortArrayList_ByName(ByVal array_)
    Dim i, j, temp
   ' MSGbox array_(1).Name
    For i = (array_.Count - 1) To 1 Step -1
      'MSGbox "i-"&i
        For j = 0 To i - 1
    'MSGbox "j-"&j
            If array_(j).Name > array_(j + 1).Name Then
                Set temp = array_(j + 1)
                Set array_(j + 1) = array_(j)
                Set array_(j) = temp
                'MSGbox array_(j).Name
            End If
        Next
  'MSGbox array_(i).Name
    Next
    Set SortArrayList_ByName = array_
End Function

'Clean characters from string...
Function strClean(inString)
  Dim outString
  Dim thisChar
  Dim i
  outString = ""
  For i = 1 To Len(inString)
    thisChar = Mid(inString, i, 1)
    Select Case thisChar
      Case ";", "'", """", "&", "%", "*" 'strings that might cause problems when executing SAS code
        thisChar = " "
      Case "<", ">", "|", "/", "\"       'string, that might cause problems whilst saving files
        thisChar = " "
    End Select
    outString = outString + thisChar
   Next
   strClean = outString
End Function

Function Checkerror(fnName)
    Checkerror = False
    
    Dim strmsg      ' As String
    Dim errNum      ' As Long
    
    If Err.Number <> 0 Then
        strmsg = "Error #" & Hex(Err.Number) & vbCrLf & "In Function " & fnName & vbCrLf & Err.Description
        'MsgBox strmsg  'Uncomment this line if you want to be notified via MessageBox of Errors in the script.
        Checkerror = True
    End If
         
End Function







