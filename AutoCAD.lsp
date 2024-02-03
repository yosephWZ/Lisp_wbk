; // @#$!@#$@!#$Progress### << ///Jan_2024..??>> @ @#$@%#!@%!Saving??@!#$@!#$@!#$!@#$@!#$
; // !!!!! don't stay at one part for more than 1 weekds...!!!!
; // !!!!! proceed as you go through by commenting .. for furhter refirinement..
; // //..//...//..//...//..//..//..//...//
; // C:\Users\User\AppData\Local\Programs\Python\Python311
; // //..//...//..//...//..//..//..//...//



; CHAPTER 1 The VBA Integrated Development Environment (AutoLispIDE)
Private Sub TextBox1_Change()
  If Len(TextBox1.Text) > 0 Then
    CommandButton1.Enabled = True
    
    Else
      CommandButton1.Enabled = False
  End If
End Sub

Public Sub Start()
  Application.ActiveDocument.SetVariable "OSMODE", 35
End Sub

(defun-q Startup ()
  (command "-vbarun" "Start")
)
(setq S::STARTUP (append S::STARTUP Startup))


; CHAPTER 2 Introduction to Visual Basic Programming 
Public Sub MyMacro()
Dim Count As Integer
Dim NumericString As String

  Count = 100
  NumericString = "555"
  
  MsgBox "Integer: " & Count & vbCrLf & _
         "String: " & NumericString
         
  Count = NumericString
  
  MsgBox "Integer: " & Count
End Sub

Public Sub IterateLayers()
Dim Layer As AcadLayer
  For Each Layer In ThisDrawing.Layers
    Debug.Print Layer.Name
  Next Layer
End Sub

Select Case UCase(ColorName)
  Case "RED"
    Layer.Color = acRed
    
  Case "YELLOW"
    Layer.Color = acYellow
    
  Case "GREEN"
    Layer.Color = acGreen
    
  Case "CYAN"
    Layer.Color = acCyan
    
  Case "BLUE"
    Layer.Color = acBlue
    
  Case "MAGENTA"
    Layer.Color = acMagenta
    
  Case "WHITE"
    Layer.Color = acWhite
    
  Case Else
    If CInt(ColorName) > 0 And CInt(ColorName) < 256 Then
      Layer.Color = CInt(ColorName)
      
      Else
        MsgBox UCase(ColorName) & " is an invalid color name", _
               vbCritical, "Invalid Color Selected"
    End If
End Select

Select Case Ucase(ColorName)
  Case RED: Layer.Color = acRed
  Case BLUE: Layer.Color = acBlue
End Select

Do While Index - 1 < Length
  Character = Asc(Mid(Name, Index, 1))
  Select Case Character
    Case 36, 45, 48 To 57, 65 To 90, 95
      IsOK = True
    Case Else
      IsOK = False
      Exit Sub
  End Select
  Index = Index + 1
Loop

Do While Not Recordset.EOF
  Debug.Print Recordset.Fields("layername").Value
  Recordset.MoveNext
Loop

Do Until Recordset.EOF
  Debug.Print Recordset.Fields("layername").Value
  Recordset.MoveNext
Loop

Public Sub DisplayLayers()
  Dim Layer As AcadLayer
     
  For Each Layer In ThisDrawing.Layers
    Debug.Print Layer.Name
  Next Layer
End Sub

Public Sub WhatColor()
Dim Layer As AcadLayer
Dim Answer As String

  For Each Layer In ThisDrawing.Layers
    Answer = InputBox("Enter color name: ")
  
    'user pressed cancel
    If Answer = "" Then Exit Sub
  
    Select Case UCase(Answer)
      Case "RED"
        Layer.Color = acRed
    
      Case "YELLOW"
        Layer.Color = acYellow
    
      Case "GREEN"
        Layer.Color = acGreen
    
      Case "CYAN"
        Layer.Color = acCyan
    
      Case "BLUE"
        Layer.Color = acBlue
    
      Case "MAGENTA"
        Layer.Color = acMagenta
      
      Case "WHITE"
        Layer.Color = acWhite
    
      Case Else
        If CInt(Answer) > 0 And CInt(Answer) < 256 Then
          Layer.Color = CInt(Answer)
      
          Else
            MsgBox UCase(Answer) & " is an invalid color name", _
                 vbCritical, "Invalid Color Selected"
        End If
    End Select
  Next Layer
End Sub

Dim mylayer As AcadLayer
mylayer = ThisDrawing.ActiveLayer

With mylayer
  .Color = acBlue
  .Linetype = "continuous"
  .Lineweight = acLnWtByLwDefault
  .Freeze = False
  .LayerOn = True
  .Lock = False
End With

 
 ; CHAPTER 3 Application Elements 
Private Sub UserForm_Activate()
  UserForm1.Width = 200
  UserForm1.Height = 150
End Sub

Private Sub cmdPick_Click()
Dim Point As Variant
  On Error Resume Next
  'hide the UserForm
  UserForm1.Hide
  'ask user to select a point
  Point = ThisDrawing.Utility.GetPoint(, "Select a point")
  If Err Then Exit Sub
  'assign values to appropriate textbox
  txtX = Point(0): txtY = Point(1): txtZ = Point(2)
  'redisplay the UserForm
  UserForm1.Show
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub UserForm_Activate()
  Label1.Caption = "# of Blocks = " & ThisDrawing.Blocks.Count
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, _
                       ByVal Shift As Integer)
  If KeyCode = 13 Then
    MsgBox "You entered: " & TextBox1.Text
  End If
End Sub

Private Sub UserForm_Activate()
  With ComboBox1
    .AddItem "Item 1"
    .AddItem "Item 2"
    .AddItem "Item 3"
  End With
End Sub

Private Sub ComboBox1_Click()
  MsgBox "You choose: " & ComboBox1.List(ComboBox1.ListIndex)
End Sub

Private Sub UserForm_Activate()
  With ListBox1
    .AddItem "Item 1"
    .AddItem "Item 2"
    .AddItem "Item 3"
  End With
End Sub

Private Sub ListBox1_Click()
  MsgBox "You clicked on: " & ListBox1.List(ListBox1.ListIndex)
End Sub

Private Sub CheckBox1_Click()
  If CheckBox1.Value Then
    MsgBox "Checked"
    
    Else
      MsgBox "Unchecked"
  End If
End Sub

Private Sub OptionButton1_Click()
  MsgBox "OptionButton1"
End Sub

Private Sub OptionButton2_Click()
  MsgBox "OptionButton2"
End Sub

Private Sub ToggleButton1_Click()
  Select Case ToggleButton1.Value
    Case False
      MsgBox "ToggleButton1 is Off"
    Case True
      MsgBox "ToggleButton1 is On"
  End Select
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

; CHAPTER 4 AutoCAD Events 
Public WithEvents objApp As AcadApplication

Option Explicit
Public objApp As New clsApplicationEvents

Public Sub InitializeEvents()
  Set objApp.objApp = ThisDrawing.Application
End Sub

Public Sub App_StartMacro()
  InitializeEvents
End Sub

Private Sub objApp_SysVarChanged(ByVal SysvarName As String, _
                                 ByVal newVal As Variant)
  MsgBox "The System Variable: " & SysvarName & " has changed to " & newVal
End Sub

Private Sub objApp_NewDrawing()
  ThisDrawing.SetVariable "SAVETIME", 30
  MsgBox "The autosave interval is currently set to 30 mins"
End Sub

Option Explicit
Public objCurrentLayer As AcadLayer
Public objPreviousLayer As AcadLayer

Private Sub AcadDocument_BeginCommand(ByVal CommandName As String)
  Set objPreviousLayer = ThisDrawing.ActiveLayer
  Select Case CommandName
    Case "LINE"
      If Not ThisDrawing.ActiveLayer.Name = "OBJECTS" Then
        Set objCurrentLayer = ThisDrawing.Layers.Add("OBJECTS")
        ThisDrawing.ActiveLayer = objCurrentLayer
      End If
  End Select
End Sub

Private Sub AcadDocument_EndCommand(ByVal CommandName As String)
  Select Case CommandName
    Case "LINE"
        ThisDrawing.ActiveLayer = objPreviousLayer
    End Select
  Set objCurrentLayer = Nothing
  Set objPreviousLayer = Nothing
End Sub

Public WithEvents objLine As AcadLine
Dim objLine As New clsObjectEvent
Public Sub InitializeEvent()
Dim dblStart(2) As Double
Dim dblEnd(2) As Double
  dblEnd(0) = 1: dblEnd(1) = 1: dblEnd(2) = 0
  Set objLine.objLine = ThisDrawing.ModelSpace.AddLine(dblStart, dblEnd)
End Sub

Private Sub objLine_Modified(ByVal pObject As AutoCAD.IAcadObject)
Dim varStartPoint As Variant
Dim varEndPoint As Variant

  varStartPoint = pObject.StartPoint
  varEndPoint = pObject.EndPoint
  MsgBox "New line runs from (" & varStartPoint(0) & ", " & _
    varStartPoint(1) & ", " & varStartPoint(2) & " ) to (" & _
    varEndPoint(0) & ", " & varEndPoint(1) & ", " & varEndPoint(2) & ")."
    
End Sub

; CHAPTER 5 User Preferences
Dim PrefProfiles As AcadPreferencesProfiles

Set PrefProfiles = ThisDrawing.Application.Preferences.Profiles

strSetPaths = ThisDrawing.Application.Preferences.Files.SupportPath

Dim strNewPath As String, strSetPath As String
strNewPath = ;c:\cadfiles\apress\dvb
strSetPath = ThisDrawing.Application.Preferences.Files.SupportPath

If Len(strSetPath & strNewPath) < 256 Then
  strNewPath = strSetPath & strNewPath
  ThisDrawing.Application.Preferences.Files.SupportPath = strNewPath
End If

intCursorSize = ThisDrawing.Application.Preferences.Display.CursorSize
'The value of the CursorSize property is a positive integer that represents the percentage of 'the cursor size to the screen size. The default value is 5, the minimum is 1, and the 'maximum is 100. All other values will generate an error. 
'The following example prompts the user to enter a value for the cursor size, and if the 'returned value lies between 1 and 100, the size is changed. Otherwise, the user is informed 'that he or she has entered an invalid value.
Dim intCursorSize As Integer

intCursorSize = ThisDrawing.Utility.GetInteger(vbCrLf & _
    "Enter number for size of cursor proportional to screen size" & vbCrLf)

If intCursorSize < 1 Or intCursorSize > 100 Then
    MsgBox "Cursor Size value must be between 1 and 100"
  Else
    ThisDrawing.Application.Preferences.Display.CursorSize = intCursorSize
End If

If ThisDrawing.Application.Preferences.OpenSave.AutoSaveInterval <> 15 Then
    ThisDrawing.Application.Preferences.OpenSave.AutoSaveInterval = 15
    MsgBox "The autosave interval has been changed to 15 minutes."
End If

strTemplatePath = ThisDrawing.Application.Preferences.Files.TemplateDwgPath

ThisDrawing.Application.Preferences.Files.TemplateDwgPath = _
"C:\Program " &  "Files\AutoCAD 2006\Templates"

Dim strPC3Path As String
strPC3Path = c:\cadfiles\plotconfigs

ThisDrawing.Application.Preferences.Files.PrinterConfigPath = strPC3Path

Public Sub SaveAsType()
Dim iSaveAsType As Integer
  iSaveAsType = ThisDrawing.Application.Preferences.OpenSave.SaveAsType
  Select Case iSaveAsType
    Case acR12_dxf
      MsgBox "Current save as format is R12_DXF", vbInformation
    Case ac2000_dwg
      MsgBox "Current save as format is 2000_DWG", vbInformation
    Case ac2000_dxf
      MsgBox "Current save as format is 2000_DXF", vbInformation
    Case ac2000_Template
      MsgBox "Current save as format is 2000_Template", vbInformation
    Case ac2004_dwg, acNative
      MsgBox "Current save as format is 2004_DWG", vbInformation
    Case ac2004_dxf
      MsgBox "Current save as format is 2004_DXF", vbInformation
    Case ac2004_Template
      MsgBox "Current save as format is 2004_Template", vbInformation
    Case acUnknown
      MsgBox "Current save as format is Unknown or Read-Only", vbInformation
  End Select
End Sub

ThisDrawing.Application.Preferences.OpenSave.SaveAsType = ac2000_dwg

ThisDrawing.Application.Preferences.System.EnableStartupDialog = False

Dim strActiveProfile as String

strActiveProfile = ThisDrawing.Application.Preferences.Profiles.ActiveProfile

ThisDrawing.Application.Preferences.Profiles.ExportProfile _
    strActiveProfile, "C:\MYPROFILE.ARG"
Dim strMyProfile As String

'name of profile
strMyProfile = "My Personal Profile"

ThisDrawing.Application.Preferences.Profiles.ImportProfile _
    strMyProfile, "C:\MYPROFILE.ARG", True

ThisDrawing.Application.Preferences.Profiles.ActiveProfile = strMyProfile

Dim strActiveProfile As String
Dim strMyProfile As String
strMyProfile = "My Personal Profile"
strActiveProfile = ThisDrawing.Application.Preferences.Profiles.ActiveProfile

With ThisDrawing.Application.Preferences.Profiles
  .RenameProfile strActiveProfile " MyBackupProfile"
  .ImportProfile strMyProfile, "C:\MYPROFILE.ARG ",True
  .ActiveProfile = strMyProfile
  .DeleteProfile "MyBackupProfile"
End With

; CHAPTER 6 Controlling Layers and Linetypes 
Dim objLayers As AcadLayers
Set objLayers = ThisDrawing.Layers

Dim objLayer As AcadLayer

Set objLayer = objLayers.Item(2)
Set objLayer = objLayers.Item("My Layer")

Public Sub ListLayers()
    Dim objLayer As AcadLayer

    For Each objLayer In ThisDrawing.Layers
        Debug.Print objLayer.Name
    Next
End Sub

Public Sub ListLayersManually()
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    Dim intI As Integer

    Set objLayers = ThisDrawing.Layers

    For intI = 0 To objLayers.Count - 1
        Set objLayer = objLayers(intI)
        Debug.Print objLayer.name
    Next
End Sub

Public Sub ListLayersBackwards()
    Dim objLayers As AcadLayers
    Dim objLayer As AcadLayer
    Dim intI As Integer

   Set objLayers = ThisDrawing.Layers

   For intI = objLayers.Count - 1 To 0 Step -1
       Set objLayer = objLayers(intI)
       Debug.Print objLayer.name
   Next
End Sub

Public Sub CheckForLayerByIteration()
    Dim objLayer As AcadLayer
    Dim strLayerName As String

    strLayername = InputBox("Enter a Layer name to search for: ")
    If "" = strLayername Then Exit Sub    ' exit if no name entered
 
    For Each objLayer In ThisDrawing.Layers    ' iterate layers 
        If 0 = StrComp(objLayer.name, strLayername, vbTextCompare) Then
            MsgBox "Layer '" & strLayername & "' exists"
            Exit Sub                           ' exit after finding layer
        End If
    Next objLayer
    MsgBox "Layer '" & strLayername & "' does not exist"
End Sub

Public Sub CheckForLayerByException()
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    strLayerName = InputBox("Enter a Layer name to search for: ")
    If "" = strLayerName Then Exit Sub      ' exit if no name entered
    
    On Error Resume Next               ' handle exceptions inline
    Set objLayer = ThisDrawing.Layers(strLayerName)
        
    If objLayer Is Nothing Then        ' check if obj has been set
        MsgBox "Layer '" & strLayerName & "' does not exist"
    Else
        MsgBox "Layer '" & objLayer.Name & "' exists"
    End If
End Sub

Public Sub AddLayer()
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    strLayerName = InputBox("Name of Layer to add: ")
    If "" = strLayerName Then Exit Sub      ' exit if no name entered
    
    On Error Resume Next               ' handle exceptions inline
    'check to see if layer already exists
    Set objLayer = ThisDrawing.Layers(strLayerName)
        
    If objLayer Is Nothing Then
        Set objLayer = ThisDrawing.Layers.Add(strLayerName)
        If objLayer Is Nothing Then ' check if obj has been set
            MsgBox "Unable to Add '" & strLayerName & "'"
        Else
            MsgBox "Added Layer '" & objLayer.Name & "'"
        End If
    Else
        MsgBox "Layer already existed"
    End If
End Sub

ThisDrawing.ActiveLayer = ThisDrawing.Layers("Walls")

Public Sub ChangeEntityLayer()
    On Error Resume Next                   ' handle exceptions inline
    Dim objEntity As AcadEntity
    Dim varPick As Variant
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    ThisDrawing.Utility.GetEntity objEntity, varPick, "Select an entity"
    If objEntity Is Nothing Then
        MsgBox "No entity was selected"
        Exit Sub  ' exit if no entity picked
    End If
    
    strLayerName = InputBox("Enter a new Layer name: ")
    If "" = strLayerName Then Exit Sub          ' exit if no name entered
    
    Set objLayer = ThisDrawing.Layers(strLayerName)
    If objLayer Is Nothing Then
        MsgBox "Layer was not recognized"
        Exit Sub   ' exit if layer not found
    End If
    objEntity.Layer = strLayerName              ' else change entity layer
End Sub

If ThisDrawing.ActiveLayer.Name = "Walls" Then ...

Public Function IsLayerActive(strLayerName As String) As Boolean
    IsLayerActive = False          'assume failure
    If 0 = StrComp(ThisDrawing.ActiveLayer.Name, strLayerName, _
                   vbTextCompare) Then
        IsLayerActive = True
    End If
End Function

Public Sub LayerActive()
    Dim strLayerName As String
    strLayerName = InputBox("Name of the Layer to check: ")
    
    If IsLayerActive(strLayerName) Then
        MsgBox "'" & strLayerName & "' is active"
    Else
        MsgBox "'" & strLayerName & "' is not active"
    End If
End Sub

Public Sub ShowOnlyLayer()
    On Error Resume Next                   ' handle exceptions inline
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    strLayerName = InputBox("Enter a Layer name to show: ")
    If "" = strLayerName Then Exit Sub          ' exit if no name entered
  
    For Each objLayer In ThisDrawing.Layers
        objLayer.LayerOn = False           ' turn off all the layers
    Next objLayer
    
    Set objLayer = ThisDrawing.Layers(strLayerName)
    If objLayer Is Nothing Then
        MsgBox "Layer does not exist"
        Exit Sub   ' exit if layer not found
    End If
    
    objLayer.LayerOn = True                ' turn on the desired layer
End Sub

Public Sub RenameLayer()
    On Error Resume Next                   ' handle exceptions inline
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    strLayerName = InputBox("Original Layer name: ")
    If "" = strLayerName Then Exit Sub          ' exit if no old name
    
    Set objLayer = ThisDrawing.Layers(strLayerName)
    If objLayer Is Nothing Then            ' exit if not found
        MsgBox "Layer '" & strLayerName & "' not found"
        Exit Sub
    End If
    
    strLayerName = InputBox("New Layer name: ")
    If "" = strLayerName Then Exit Sub          ' exit if no new name
    
    objLayer.Name = strLayerName                ' try and change name
    If Err Then                            ' check if it worked
        MsgBox "Unable to rename layer: " & vbCr & Err.Description
    Else
        MsgBox "Layer renamed to '" & strLayerName & "'"
    End If
End Sub

Public Sub DeleteLayer()
    On Error Resume Next                   ' handle exceptions inline
    Dim strLayerName As String
    Dim objLayer As AcadLayer
    
    strLayerName = InputBox("Layer name to delete: ")
    If "" = strLayerName Then Exit Sub          ' exit if no old name

    Set objLayer = ThisDrawing.Layers(strLayerName)
    If objLayer Is Nothing Then            ' exit if not found
        MsgBox "Layer '" & strLayerName & "' not found"
        Exit Sub
    End If

    objLayer.Delete                        ' try to delete it
    If Err Then                            ' check if it worked
        MsgBox "Unable to delete layer: " & vbCr & Err.Description
    Else
        MsgBox "Layer '" & strLayerName & "' deleted"
    End If
End Sub

Dim objLayer As AcadLayer
Dim strLayerHandle As String

Set objLayer = ThisDrawing.Layers("0")
strLayerHandle = objLayer.Handle

Public Sub Layer0Linetype()
Dim objLayer As AcadLayer
Dim strLayerLinetype As String

    Set objLayer = ThisDrawing.Layers("0")
    objLayer.Linetype = "Continuous"
    strLayerLinetype = objLayer.Linetype

End Sub

Public Sub GetLwt()
    Dim objLayer As AcadLayer
    Dim lwtLweight As Integer
    Set objLayer = ThisDrawing.ActiveLayer
    lwtLweight = objLayer.Lineweight
    Debug.Print "Lineweight is " & lwtLweight
End Sub

Public Sub CheckForLinetypeByIteration()
    Dim objLinetype As AcadLineType
    Dim strLinetypeName As String

    strLinetypeName = InputBox("Enter a Linetype name to search for: ")
    If "" = strLinetypeName Then Exit Sub         ' exit if no name entered
 
    For Each objLinetype In ThisDrawing.Linetypes
        If 0 = StrComp(objLinetype.Name, strLinetypeName, vbTextCompare) Then
            MsgBox "Linetype '" & strLinetypeName & "' exists"
            Exit Sub                           ' exit after finding linetype
        End If
    Next objLinetype

    MsgBox "Linetype '" & strLinetypeName & "' does not exist"
End Sub

Public Sub CheckForLinetypeByException()
    Dim strLinetypeName As String
    Dim objLinetype As AcadLineType
    
    strLinetypeName = InputBox("Enter a Linetype name to search for: ")
    If "" = strLinetypeName Then Exit Sub      ' exit if no name entered
    
    On Error Resume Next               ' handle exceptions inline
    Set objLinetype = ThisDrawing.Linetypes(strLinetypeName)
        
    If objLinetype Is Nothing Then     ' check if obj has been set
        MsgBox "Linetype '" & strLinetypeName & "' does not exist"
    Else
        MsgBox "Linetype '" & objLinetype.Name & "' exists"
    End If
End Sub

Public Sub LoadLinetype()
    Dim strLinetypeName As String
    Dim objLinetype As AcadLineType
    
    strLinetypeName = InputBox("Enter a Linetype name" & _
                               " to load from ACAD.LIN: ")
    If "" = strLinetypeName Then Exit Sub      ' exit if no name entered
    
    On Error Resume Next               ' handle exceptions inline
    ThisDrawing.Linetypes.Load strLinetypeName, "acad.lin"
        
    If Err Then                        ' check if err was thrown
        MsgBox "Error loading '" & strLinetypeName & "'" & vbCr & _
                Err.Description
    Else
        MsgBox "Loaded Linetype '" & strLinetypeName & "'"
    End If
End Sub

Public Sub ChangeEntityLinetype()
    On Error Resume Next                      ' handle exceptions inline
    Dim objEntity As AcadEntity
    Dim varPick As Variant
    Dim strLinetypeName As String
    Dim objLinetype As AcadLineType
    
    ThisDrawing.Utility.GetEntity objEntity, varPick, "Select an entity"
    If objEntity Is Nothing Then Exit Sub     ' exit if no entity picked
    
    strLinetypeName = InputBox("Enter a new Linetype name: ")
    If "" = strLinetypeName Then Exit Sub      ' exit if no name entered
    
    Set objLinetype = ThisDrawing.Linetypes(strLinetypeName)
    If objLinetype Is Nothing Then
        MsgBox "Linetype is not loaded"
        Exit Sub   ' exit if linetype not found
    End If
    
    objEntity.Linetype = strLinetypeName        ' else change entity layer
End Sub

Public Sub RenameLinetype()
    On Error Resume Next                   ' handle exceptions inline
    On Error Resume Next                   ' handle exceptions inline
    Dim strLinetypeName As String
    Dim objLinetype As AcadLineType
    
    strLinetypeName = InputBox("Original Linetype name: ")
    If "" = strLinetypeName Then Exit Sub          ' exit if no old name
    
    Set objLinetype = ThisDrawing.Linetypes(strLinetypeName)
    If objLinetype Is Nothing Then         ' exit if not found
        MsgBox "Linetype '" & strLinetypeName & "' not found"
        Exit Sub
    End If
    
    strLinetypeName = InputBox("New Linetype name: ")
    If "" = strLinetypeName Then Exit Sub          ' exit if no new name
    
    objLinetype.Name = strLinetypeName             ' try and change name
    If Err Then                            ' check if it worked
        MsgBox "Unable to rename Linetype: " & vbCr & Err.Description
    Else
        MsgBox "Linetype renamed to '" & strLinetypeName & "'"
    End If
End Sub

Public Sub DeleteLinetype()
    On Error Resume Next                   ' handle exceptions inline
    Dim strLinetypeName As String
    Dim objLinetype As AcadLineType
    
    strLinetypeName = InputBox("Linetype name to delete: ")
    If "" = strLinetypeName Then Exit Sub          ' exit if no old name

    Set objLinetype = ThisDrawing.Linetypes(strLinetypeName)
    If objLinetype Is Nothing Then            ' exit if not found
        MsgBox "Linetype '" & strLinetypeName & "' not found"
        Exit Sub
    End If

    objLinetype.Delete                        ' try to delete it
    If Err Then                            ' check if it worked
        MsgBox "Unable to delete linetype: " & vbCr & Err.Description
    Else
        MsgBox "Linetype '" & strLinetypeName & "' deleted"
    End If
End Sub

Public Sub DescribeLinetype()
    On Error Resume Next                         ' handle exceptions inline
    Dim strLinetypeName As String
    Dim strLinetypeDescription As String
    Dim objLinetype As AcadLineType
    
    strLinetypeName = InputBox("Enter the Linetype name: ")
    If "" = strLinetypeName Then Exit Sub              ' exit if no old name
    
    Set objLinetype = ThisDrawing.Linetypes(strLinetypeName)
    If objLinetype Is Nothing Then               ' exit if not found
        MsgBox "Linetype '" & strLinetypeName & "' not found"
        Exit Sub
    End If
    
    strLinetypeDescription = InputBox("Enter the Linetype description: ")
    If "" = strLinetypeDescription Then Exit Sub       ' exit if no new name
    
    objLinetype.Description = strLinetypeDescription    ' try and change name
    If Err Then                                  ' check if it worked
        MsgBox "Unable to alter Linetype: " & vbCr & Err.Description
    Else
        MsgBox "Linetype '" & strLinetypeName & "' description changed"
    End If

; CHAPTER 7 User Interaction and the Utility Object 
Private Sub cmdGetReal_Click()
Dim dblInput As Double
    Me.Hide
    dblInput = ThisDrawing.Utility.GetReal("Enter a real value: ")
    Me.Show
End Sub

Public Sub TestUserInput()
Dim strInput As String
    With ThisDrawing.Utility
        .InitializeUserInput 1, "Line Arc Circle laSt"
        strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
        .Prompt "You selected '" & strInput & "'"
    End With
End Sub

Private Sub CommandButton1_Click()
Dim strInput As String
    Me.Hide
    With ThisDrawing.Utility
        .InitializeUserInput 1, "Line Arc Circle laSt"
        strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
        MsgBox "You selected '" & strInput & "'"
    End With
    Me.Show
End Sub

Public Sub TestPrompt()
    ThisDrawing.Utility.Prompt vbCrLf & "This is a simple message"
End Sub

Public Sub TestUserInput()
Dim strInput As String
    With ThisDrawing.Utility
        .InitializeUserInput 1, "Line Arc Circle laSt"
        strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
        .Prompt vbCr & "You selected '" & strInput & "'"
    End With
End Sub

Public Sub TestGetKeyword()
Dim strInput As String
    
    With ThisDrawing.Utility
        .InitializeUserInput 0, "Line Arc Circle"
        strInput = .GetKeyword(vbCr & "Command [Line/Arc/Circle]: ")
    End With
    
    Select Case strInput

        Case "Line": ThisDrawing.SendCommand "_Line" & vbCr
        Case "Arc": ThisDrawing.SendCommand "_Arc" & vbCr
        Case "Circle": ThisDrawing.SendCommand "_Circle" & vbCr
        Case Else:  MsgBox "You pressed Enter."

    End Select
End Sub

Public Sub TestGetString()
Dim strInput As String
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter a string: ")
        .Prompt vbCr & "You entered '" & strInput & "' "
    End With
End Sub

Public Sub TestGetInteger()
Dim intInput As Integer
    
    With ThisDrawing.Utility
        intInput = .GetInteger(vbCr & "Enter an integer: ")
        .Prompt vbCr & "You entered " & intInput
    End With
End Sub

Public Sub TestGetReal()
Dim dblInput As Double
    
    With ThisDrawing.Utility
        dblInput = .GetReal(vbCrLf & "Enter an real: ")
        .Prompt vbCr & "You entered " & dblInput
    End With
End Sub

Public Sub TestGetPoint()
Dim varPick As Variant
    
    With ThisDrawing.Utility
        varPick = .GetPoint(, vbCr & "Pick a point: ")
        .Prompt vbCr & varPick(0) & "," & varPick(1)
    End With
End Sub

Public Sub TestGetCorner()
Dim varBase As Variant
Dim varPick As Variant
    
    With ThisDrawing.Utility
        varBase = .GetPoint(, vbCr & "Pick the first corner: ")
        .Prompt vbCrLf & varBase(0) & "," & varBase(1)
        varPick = .GetCorner(varBase, vbLf & "Pick the second: ")
        .Prompt vbCr & varPick(0) & "," & varPick(1)
    End With
End Sub

Public Sub TestGetDistance()
Dim dblInput As Double
Dim dblBase(2) As Double
    
    dblBase(0) = 0:  dblBase(1) = 0:  dblBase(2) = 0
    
    With ThisDrawing.Utility
        dblInput = .GetDistance(dblBase, vbCr & "Enter a distance: ")
        .Prompt vbCr & "You entered " & dblInput
    End With
End Sub

Public Sub TestGetAngle()
Dim dblInput As Double
    
    ThisDrawing.SetVariable "DIMAUNIT", acDegrees
    
    With ThisDrawing.Utility
        dblInput = .GetAngle(, vbCr & "Enter an angle: ")
        .Prompt vbCr & "Angle in radians: " & dblInput
    End With
End Sub

Public Sub TestGetInput()
Dim intInput As Integer
Dim strInput As String
    
On Error Resume Next                    ' handle exceptions inline
    With ThisDrawing.Utility
        strInput = .GetInput()
        .InitializeUserInput 0, "Line Arc Circle"
        intInput = .GetInteger(vbCr & "Integer or [Line/Arc/Circle]: ")
    
        If Err.Description Like "*error*" Then
            .Prompt vbCr & "Input Cancelled"
        ElseIf Err.Description Like "*keyword*" Then
            strInput = .GetInput()
            Select Case strInput
                Case "Line": ThisDrawing.SendCommand "_Line" & vbCr
                Case "Arc": ThisDrawing.SendCommand "_Arc" & vbCr
                Case "Circle": ThisDrawing.SendCommand "_Circle" & vbCr
                Case Else: .Prompt vbCr & "Null Input entered"
            End Select
        Else
            .Prompt vbCr & "You entered " & intInput
        End If
    End With

End Sub

Public Sub TestGetInputBug()
On Error Resume Next                    ' handle exceptions inline
    
    With ThisDrawing.Utility
        
        '' first keyword input
        .InitializeUserInput 1, "Alpha Beta Ship"
        .GetInteger vbCr & "Option [Alpha/Beta/Ship]: "
        MsgBox "You entered: " & .GetInput()
            
        '' second keyword input - hit Enter here
        .InitializeUserInput 0, "Bug May Slip"
        .GetInteger vbCr & "Hit enter [Bug/May/Slip]: "
        MsgBox "GetInput still returns: " & .GetInput()
    End With
End Sub

Public Sub TestGetInputWorkaround()
Dim strBeforeKeyword As String
Dim strKeyword As String
    
On Error Resume Next                    ' handle exceptions inline
    With ThisDrawing.Utility
        
        '' first keyword input
        .InitializeUserInput 1, "This Bug Stuff"
        .GetInteger vbCrLf & "Option [This/Bug/Stuff]: "
        MsgBox "You entered: " & .GetInput()
            
        '' get lingering keyword
        strBeforeKeyword = .GetInput()
            
        '' second keyword input - press Enter
        .InitializeUserInput 0, "Make Life Rough"
        .GetInteger vbCrLf & "Hit enter [Make/Life/Rough]: "
        strKeyword = .GetInput()
        
        '' if input = lingering it might be null input
        If strKeyword = strBeforeKeyword Then
            MsgBox "Looks like null input: " & strKeyword
        Else
            MsgBox "This time you entered: " & strKeyword
        End If
    End With
End Sub

Public Sub TestGetEntity()
Dim objEnt As AcadEntity
Dim varPick As Variant

On Error Resume Next
    With ThisDrawing.Utility
        .GetEntity objEnt, varPick, vbCr & "Pick an entity: "
        If objEnt Is Nothing Then 'check if object was picked.
            .Prompt vbCrLf & "You did not pick as entity"
            Exit Sub
        End If
        .Prompt vbCr & "You picked a " & objEnt.ObjectName
        .Prompt vbCrLf & "At " & varPick(0) & "," & varPick(1)
    End With
End Sub

Public Sub TestGetSubEntity()
Dim objEnt As AcadEntity
Dim varPick As Variant
Dim varMatrix As Variant
Dim varParents As Variant
Dim intI As Integer
Dim intJ As Integer
Dim varID As Variant
    
    With ThisDrawing.Utility
        
        '' get the subentity from the user
        .GetSubEntity objEnt, varPick, varMatrix, varParents, _
            vbCr & "Pick an entity: "
        
        '' print some information about the entity
        .Prompt vbCr & "You picked a " & objEnt.ObjectName
        .Prompt vbCrLf & "At " & varPick(0) & "," & varPick(1)
        
        '' dump the varMatrix
        If Not IsEmpty(varMatrix) Then
            .Prompt vbLf & "MCS to WCS Translation varMatrix:"
            
            '' format varMatrix row
            For intI = 0 To 3
                .Prompt vbLf & "["
                
                '' format varMatrix column
                For intJ = 0 To 3
                    .Prompt "(" & varMatrix(intI, intJ) & ")"
                Next intJ
                
                .Prompt "]"
            Next intI
            .Prompt vbLf
        End If
        
        '' if it has a parent nest
        If Not IsEmpty(varParents) Then
            
            .Prompt vbLf & "Block nesting:"
            
            '' depth counter
            intI = -1
            
            '' traverse most to least deep (reverse order)
            For intJ = UBound(varParents) To LBound(varParents) Step -1
                
                '' increment depth
                intI = intI + 1
                
                '' indent output
                .Prompt vbLf & Space(intI * 2)
                
                '' parent object ID
                varID = varParents(intJ)
                
                '' parent entity
                Set objEnt = ThisDrawing.ObjectIdToObject(varID)
                                
                '' print info about parent
                .Prompt objEnt.ObjectName & " : " & objEnt.Name
            Next intJ
            .Prompt vbLf
        End If
        .Prompt vbCr
    End With
End Sub

Public Sub TestAngleToReal()
Dim strInput As String
Dim dblAngle As Double
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter an angle: ")
        dblAngle = .AngleToReal(strInput, acDegrees)
        .Prompt vbCr & "Radians: " & dblAngle
    End With
End Sub

Public Sub TestAngleToString()
Dim strInput As String
Dim strOutput As String
Dim dblAngle As Double
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter an angle: ")
        dblAngle = .AngleToReal(strInput, acDegrees)
        .Prompt vbCr & "Radians: " & dblAngle
        
        strOutput = .AngleToString(dblAngle, acDegrees, 4)
        .Prompt vbCrLf & "Degrees: " & strOutput
    End With
End Sub

Public Sub TestDistanceToReal()
Dim strInput As String
Dim dblDist As Double
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter a distance: ")
        dblDist = .DistanceToReal(strInput, acArchitectural)
        .Prompt vbCr & "Distance: " & dblDist
    End With
End Sub

Public Sub TestRealToString()
Dim strInput As String
Dim strOutput As String
Dim dblDist As Double
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter a distance: ")
        dblDist = .DistanceToReal(strInput, acArchitectural)
        .Prompt vbCr & "Double: " & dblDist
        
        strOutput = .RealToString(dblDist, acArchitectural, 4)
        .Prompt vbCrLf & "Distance: " & strOutput
    End With
End Sub

Public Sub TestAngleFromXAxis()
Dim varStart As Variant
Dim varEnd As Variant
Dim dblAngle As Double
    
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        varEnd = .GetPoint(varStart, vbCr & "Pick the end point: ")
        dblAngle = .AngleFromXAxis(varStart, varEnd)
        .Prompt vbCr & "The angle from the X-axis is " _
           & .AngleToString(dblAngle, acDegrees, 2) & " degrees"
    End With
End Sub

Public Sub TestPolarPoint()
Dim varpnt1 As Variant
Dim varpnt2 As Variant
Dim varpnt3 As Variant
Dim varpnt4 As Variant
Dim dblAngle As Double
Dim dblLength As Double
Dim dblHeight As Double
Dim dbl90Deg As Double
        
    '' get the point, length, height, and angle from user
    With ThisDrawing.Utility
        
        '' get point, length, height, and angle from user
        varpnt1 = .GetPoint(, vbCr & "Pick the start point: ")
        dblLength = .GetDistance(varpnt1, vbCr & "Enter the length: ")
        dblHeight = .GetDistance(varpnt1, vbCr & "Enter the height: ")
        dblAngle = .GetAngle(varpnt1, vbCr & "Enter the angle: ")
        
        '' calculate remaining rectangle points
        dbl90Deg = .AngleToReal("90d", acDegrees)
        varpnt2 = .PolarPoint(varpnt1, dblAngle, dblLength)
        varpnt3 = .PolarPoint(varpnt2, dblAngle + dbl90Deg, dblHeight)
        varpnt4 = .PolarPoint(varpnt3, dblAngle + (dbl90Deg * 2), dblLength)
    End With
    
    '' draw the rectangle
    With ThisDrawing
        .ModelSpace.AddLine varpnt1, varpnt2
        .ModelSpace.AddLine varpnt2, varpnt3
        .ModelSpace.AddLine varpnt3, varpnt4
        .ModelSpace.AddLine varpnt4, varpnt1
    End With
End Sub

Public Sub TestTranslateCoordinates()
Dim varpnt1 As Variant
Dim varpnt1Ucs As Variant
Dim varpnt2 As Variant
        
    '' get the point, length, height, and angle from user
    With ThisDrawing.Utility
        
        '' get start point
        varpnt1 = .GetPoint(, vbCr & "Pick the start point: ")
        
        '' convert to UCS for use in the base point rubberband line
        varpnt1Ucs = .TranslateCoordinates(varpnt1, acWorld, acUCS, False)

        '' get end point
        varpnt2 = .GetPoint(varpnt1Ucs, vbCr & "Pick the end point: ")
    End With
    
    '' draw the line
    With ThisDrawing
        .ModelSpace.AddLine varpnt1, varpnt2
    End With
End Sub

Public Sub TestIsURL()
Dim strInput As String
    
    With ThisDrawing.Utility
        strInput = .GetString(True, vbCr & "Enter a URL: ")
        If .IsURL(strInput) Then
            MsgBox "You entered a valid URL"
        Else
            MsgBox "That was not a URL"
        End If
    End With
End Sub

Public Sub TestLaunchBrowserDialog()
Dim strStartUrl As String
Dim strInput As String
Dim blnStatus As Boolean
    
    strStartUrl = InputBox("Enter a URL", , "http://www.apress.com")
    
    With ThisDrawing.Utility
        If .IsURL(strStartUrl) = False Then
            MsgBox "You did not enter a valid URL"
            Exit Sub
        End If
            
        blnStatus = .LaunchBrowserDialog(strInput, _
                                        "Select a URL", _
                                        "Select", _
                                        strStartUrl, _
                                        "ContractCADDgroup", _
                                        True)
        If Not blnStatus Then
            MsgBox "You cancelled without selecting anything"
            Exit Sub
        End If
        
        If strStartUrl = strInput Then
            MsgBox "You selected the original URL"
        Else
            MsgBox "You selected: " & strInput
        End If
    End With
End Sub

Public Sub TestGetRemoteFile()
Dim strUrl As String
Dim strLocalName As String
Dim blnStatus As Boolean
    
 strUrl = InputBox("Enter a URL of a drawing file")
    
    With ThisDrawing.Utility
        If .IsURL(strUrl) = False Then
            MsgBox "You did not enter a valid URL"
            Exit Sub
        End If
 
        .GetRemoteFile strUrl, strLocalName, True
        If Err Then
            MsgBox "Failed to download: " & strUrl & vbCr & Err.Description
        Else
            MsgBox "The file was downloaded to: " & strLocalName
        End If
    End With
End Sub 

Public Sub TestIsRemoteFile()
Dim strOutputURL As String
Dim strLocalName As String
         
    strLocalName = InputBox("Enter the file and path name to check")
    If strLocalName = "" Then Exit Sub
    With ThisDrawing.Utility
        '' check if the local file is from a URL
        If .IsRemoteFile(strLocalName, strOutputURL) Then
            MsgBox "This file was downloaded from: " & strOutputURL
        Else
            MsgBox "This file was not downloaded from a URL"
        End If
    End With
End Sub

; CHAPTER 8 Drawing Objects 
If ThisDrawing.ActiveSpace = acModelSpace Then
    MsgBox "The active space is model space"
Else
    MsgBox "The active space is paper space"
End If

Public Sub ToggleSpace()
    With ThisDrawing
        If .ActiveSpace = acModelSpace Then
            .ActiveSpace = acPaperSpace
        Else
            .ActiveSpace = acModelSpace
        End If
    End With
End Sub

ThisDrawing.ActiveSpace = (ThisDrawing.ActiveSpace + 1) Mod 2

Public Sub TestAddArc()
Dim varCenter As Variant
Dim dblRadius As Double
Dim dblStart As Double
Dim dblEnd As Double
Dim objEnt As AcadArc
      
    On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varCenter = .GetPoint(, vbCr & "Pick the center point: ")
        dblRadius = .GetDistance(varCenter, vbCr & "Enter the radius: ")
        dblStart = .GetAngle(varCenter, vbCr & "Enter the start angle: ")
        dblEnd = .GetAngle(varCenter, vbCr & "Enter the end angle: ")
    End With
      
    '' draw the arc
If ThisDrawing.ActiveSpace = acModelSpace Then
    Set objEnt = ThisDrawing.ModelSpace.AddArc(varCenter, dblRadius, _
                                               dblStart, dblEnd)
Else
Set objEnt = ThisDrawing.PaperSpace.AddArc(varCenter, dblRadius, dblStart, dblEnd)
End If
    objEnt.Update
End Sub

Public Sub TestAddCircle()
Dim varCenter As Variant
Dim dblRadius As Double
Dim objEnt As AcadCircle
      
On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varCenter = .GetPoint(, vbCr & "Pick the centerpoint: ")
        dblRadius = .GetDistance(varCenter, vbCr & "Enter the radius: ")
    End With
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddCircle(varCenter, dblRadius)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddCircle(varCenter, dblRadius)
End If
    objEnt.Update
End Sub

Public Sub TestAddEllipse()
Dim dblCenter(0 To 2) As Double
Dim dblMajor(0 To 2) As Double
Dim dblRatio As Double
Dim dblStart As Double
Dim dblEnd As Double
Dim objEnt As AcadEllipse

On Error Resume Next
      
    '' setup the ellipse parameters
    dblCenter(0) = 0: dblCenter(1) = 0: dblCenter(2) = 0
    dblMajor(0) = 10: dblMajor(1) = 0: dblMajor(2) = 0
    dblRatio = 0.5
     
    '' draw the ellipse
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddEllipse(dblCenter, dblMajor, dblRatio)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddEllipse(dblCenter, dblMajor, dblRatio)
End If
    objEnt.Update
    
    '' get angular input from user
    With ThisDrawing.Utility
        dblStart = .GetAngle(dblCenter, vbCr & "Enter the start angle: ")
        dblEnd = .GetAngle(dblCenter, vbCr & "Enter the end angle: ")
    End With
      
    '' convert the ellipse into elliptical arc
With objEnt
    .StartAngle = dblStart
    .EndAngle = dblEnd
    .Update
End With
End Sub

Public Sub TestAddLine()
Dim varStart As Variant
Dim varEnd As Variant
Dim objEnt As AcadLine

On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        varEnd = .GetPoint(varStart, vbCr & "Pick the end point: ")
    End With
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddLine(varStart, varEnd)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddLine(varStart, varEnd)
End If
    objEnt.Update
End Sub

Public Sub TestAddLWPolyline()
Dim objEnt As AcadLWPolyline
Dim dblVertices() As Double

    '' setup initial points
    ReDim dblVertices(11)
    dblVertices(0) = 0#: dblVertices(1) = 0#
    dblVertices(2) = 10#: dblVertices(3) = 0#
    dblVertices(4) = 10#: dblVertices(5) = 10#
    dblVertices(6) = 5#: dblVertices(7) = 5#
    dblVertices(8) = 2#: dblVertices(9) = 2#
    dblVertices(10) = 0#: dblVertices(11) = 10#
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddLightWeightPolyline(dblVertices)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddLightWeightPolyline(dblVertices)
End If
    objEnt.Closed = True
    objEnt.Update
End Sub

Public Sub TestAddVertex()
    On Error Resume Next
    Dim objEnt As AcadLWPolyline
    Dim dblNew(0 To 1) As Double
    Dim lngLastVertex As Long
    Dim varPick As Variant
    Dim varWCS As Variant
    
    With ThisDrawing.Utility
        
        '' get entity from user
        .GetEntity objEnt, varPick, vbCr & "Pick a polyline <exit>: "
        
        '' exit if no pick
        If objEnt Is Nothing Then Exit Sub
        
        '' exit if not a lwpolyline
        If objEnt.ObjectName <> "AcDbPolyline" Then
            MsgBox "You did not pick a polyline"
            Exit Sub
        End If
        
        '' copy last vertex of pline into pickpoint to begin loop
        ReDim varPick(2)
        varPick(0) = objEnt.Coordinates(UBound(objEnt.Coordinates) - 1)
        varPick(1) = objEnt.Coordinates(UBound(objEnt.Coordinates))
        varPick(2) = 0
        
        '' append vertexes in a loop
        Do
            '' translate picked point to UCS for basepoint below
            varWCS = .TranslateCoordinates(varPick, acWorld, acUCS, True)
            
            '' get user point for new vertex, use last pick as basepoint
            varPick = .GetPoint(varWCS, vbCr & "Pick another point <exit>: ")
            
            '' exit loop if no point picked
            If Err Then Exit Do
            
            '' copy picked point X and Y into new 2d point
            dblNew(0) = varPick(0):  dblNew(1) = varPick(1)
            
            '' get last vertex offset.  it is one half the array size
            lngLastVertex = (UBound(objEnt.Coordinates) + 1) / 2
            
            '' add new vertex to pline at last offset
            objEnt.AddVertex lngLastVertex, dblNew
        Loop
    End With
    objEnt.Update
End Sub

Public Sub TestAddMLine()
Dim objEnt As AcadMLine
Dim dblVertices(17) As Double

    '' setup initial points
    dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
    dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
    dblVertices(6) = 10: dblVertices(7) = 10: dblVertices(8) = 0
    dblVertices(9) = 5: dblVertices(10) = 10: dblVertices(11) = 0
    dblVertices(12) = 5: dblVertices(13) = 5: dblVertices(14) = 0
    dblVertices(15) = 0: dblVertices(16) = 5: dblVertices(17) = 0

    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddMLine(dblVertices)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddMLine(dblVertices)
End If
    objEnt.Update
End Sub

Public Sub TestAddPolyline()
Dim objEnt As AcadPolyline
Dim dblVertices(17) As Double

    '' setup initial points
    dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
    dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
    dblVertices(6) = 7: dblVertices(7) = 10: dblVertices(8) = 0
    dblVertices(9) = 5: dblVertices(10) = 7: dblVertices(11) = 0
    dblVertices(12) = 6: dblVertices(13) = 2: dblVertices(14) = 0
    dblVertices(15) = 0: dblVertices(16) = 4: dblVertices(17) = 0

    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddPolyline(dblVertices)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddPolyline(dblVertices)
End If
    objEnt.Type = acFitCurvePoly
    objEnt.Closed = True
    objEnt.Update
End Sub

Public Sub TestPolylineType()
Dim objEnt As AcadPolyline
Dim varPick As Variant
Dim strType As String
Dim intType As Integer

On Error Resume Next
    
    With ThisDrawing.Utility
        .GetEntity objEnt, varPick, vbCr & "Pick a polyline: "
        If Err Then
            MsgBox "That is not a Polyline"
            Exit Sub
        End If
    
        .InitializeUserInput 1, "Simple Fit Quad Cubic"
        strType = .GetKeyword(vbCr & "Change type [Simple/Fit/Quad/Cubic]: ")
                
        Select Case strType
            Case "Simple": intType = acSimplePoly
            Case "Fit": intType = acFitCurvePoly
            Case "Quad": intType = acQuadSplinePoly
            Case "Cubic": intType = acCubicSplinePoly
        End Select
    End With

    objEnt.Type = intType
    objEnt.Closed = True
    objEnt.Update
End Sub

Public Sub TestAddBulge()
Dim objEnt As AcadPolyline
Dim dblVertices(17) As Double

    '' setup initial points
    dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
    dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
    dblVertices(6) = 7: dblVertices(7) = 10: dblVertices(8) = 0
    dblVertices(9) = 5: dblVertices(10) = 7: dblVertices(11) = 0
    dblVertices(12) = 6: dblVertices(13) = 2: dblVertices(14) = 0
    dblVertices(15) = 0: dblVertices(16) = 4: dblVertices(17) = 0

    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddPolyline(dblVertices)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddPolyline(dblVertices)
End If
    objEnt.Type = acSimplePoly
    'add bulge to the fourth segment
    objEnt.SetBulge 3, 0.5
    objEnt.Update
End Sub

Public Sub TestAddRay()
Dim varStart As Variant
Dim varEnd As Variant
Dim objEnt As AcadRay

On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        varEnd = .GetPoint(varStart, vbCr & "Indicate a direction: ")
    End With
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddRay(varStart, varEnd)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddRay(varStart, varEnd)
End If
    objEnt.Update
End Sub

Public Sub TestAddSpline()
Dim objEnt As AcadSpline
Dim dblBegin(0 To 2) As Double
Dim dblEnd(0 To 2) As Double
Dim dblPoints(14) As Double
    
    '' set tangencies
    dblBegin(0) = 1.5:   dblBegin(1) = 0#:   dblBegin(2) = 0
    dblEnd(0) = 1.5:  dblEnd(1) = 0#:  dblEnd(2) = 0
    
    '' set the fit dblPoints
    dblPoints(0) = 0: dblPoints(1) = 0: dblPoints(2) = 0
    dblPoints(3) = 3: dblPoints(4) = 5: dblPoints(5) = 0
    dblPoints(6) = 5: dblPoints(7) = 0: dblPoints(8) = 0
    dblPoints(9) = 7: dblPoints(10) = -5: dblPoints(11) = 0
    dblPoints(12) = 10: dblPoints(13) = 0: dblPoints(14) = 0
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddSpline(dblPoints, dblBegin, dblEnd)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddSpline(dblPoints, dblBegin, dblEnd)
End If
    objEnt.Update
End Sub

Public Sub TestAddXline()
Dim varStart As Variant
Dim varEnd As Variant
Dim objEnt As AcadXline

On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        varEnd = .GetPoint(varStart, vbCr & "Indicate an angle: ")
    End With
      
    '' draw the entity
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddXline(varStart, varEnd)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddXline(varStart, varEnd)
End If
    objEnt.Update
End Sub

Public Sub TestAddHatch()
Dim varCenter As Variant
Dim dblRadius As Double
Dim dblAngle As Double
Dim objEnt As AcadHatch
Dim varOuter() As AcadEntity
Dim varInner() As AcadEntity

On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varCenter = .GetPoint(, vbCr & "Pick the center point: ")
        dblRadius = .GetDistance(varCenter, vbCr & "Indicate the radius: ")
        dblAngle = .AngleToReal("180", acDegrees)
    End With
    
    '' draw the entities    With ThisDrawing.ModelSpace
    
        '' draw the Outer loop (circle)
        ReDim varOuter(0)
        Set varOuter(0) = .AddCircle(varCenter, dblRadius)
        
        '' draw then Inner loop (semicircle)
        ReDim varInner(1)
        Set varInner(0) = .AddArc(varCenter, dblRadius * 0.5, 0, dblAngle)
        Set varInner(1) = .AddLine(varInner(0).StartPoint, _
                                   varInner(0).EndPoint)
        
        '' create the Hatch object
        Set objEnt = .AddHatch(acHatchPatternTypePreDefined, "ANSI31", True)
    
        '' append boundaries to the hatch
        objEnt.AppendOuterLoop varOuter
        objEnt.AppendInnerLoop varInner
        
        '' evaluate and display hatched boundaries
        objEnt.Evaluate
        objEnt.Update
    End With
End Sub

Public Sub TestAddMText()
Dim varStart As Variant
Dim dblWidth As Double
Dim strText As String
Dim objEnt As AcadMText
    
On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        dblWidth = .GetDistance(varStart, vbCr & "Indicate the width: ")
        strText = .GetString(True, vbCr & "Enter the text: ")
    End With
    
    '' add font and size formatting
    strText = "\Fromand.shx;\H0.5;" & strText
        
    '' create the mtext

    Set objEnt = ThisDrawing.ModelSpace.AddMText(varStart, dblWidth, strText)
    objEnt.Update
End Sub

Public Sub TestAddPoint()
Dim objEnt As AcadPoint
Dim varPick As Variant
Dim strType As String
Dim intType As Integer
Dim dblSize As Double
    
On Error Resume Next
    
    With ThisDrawing.Utility
        '' get the pdmode center type
        .InitializeUserInput 1, "Dot None Cross X Tick"
        strType = .GetKeyword(vbCr & "Center type [Dot/None/Cross/X/Tick]: ")
        If Err Then Exit Sub
        
        Select Case strType
            Case "Dot": intType = 0
            Case "None": intType = 1
            Case "Cross": intType = 2
            Case "X": intType = 3
            Case "Tick": intType = 4
        End Select
        
        '' get the pdmode surrounding type
        .InitializeUserInput 1, "Circle Square Both"
        strType = .GetKeyword(vbCr & "Outer type [Circle/Square/Both]: ")
        If Err Then Exit Sub
        
        Select Case strType
            Case "Circle": intType = intType + 32
            Case "Square": intType = intType + 64
            Case "Both": intType = intType + 96
        End Select
        
        '' get the pdsize
        .InitializeUserInput 1, ""
        dblSize = .GetDistance(, vbCr & "Enter a point size: ")
        If Err Then Exit Sub
        
        '' set the system varibles
        With ThisDrawing
        .SetVariable "PDMODE", intType
        .SetVariable "PDSIZE", dblSize
        End With
        '' now add points in a loop
        Do
            '' get user point for new vertex, use last pick as basepoint
            varPick = .GetPoint(, vbCr & "Pick a point <exit>: ")
            
            '' exit loop if no point picked
            If Err Then Exit Do
            
            '' add new vertex to pline at last offset

            ThisDrawing.ModelSpace.AddPoint varPick
        Loop
    End With
End Sub

Public Sub TestAddRegion()
Dim varCenter As Variant
Dim varMove As Variant
Dim dblRadius As Double
Dim dblAngle As Double
Dim varRegions As Variant
Dim objEnts() As AcadEntity
    
    On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varCenter = .GetPoint(, vbCr & "Pick the center point: ")
        dblRadius = .GetDistance(varCenter, vbCr & "Indicate the radius: ")
        dblAngle = .AngleToReal("180", acDegrees)
    End With
    
    '' draw the entities
    With ThisDrawing.ModelSpace
    
        '' draw the outer region (circle)
        ReDim objEnts(2)
        Set objEnts(0) = .AddCircle(varCenter, dblRadius)
        
        '' draw the inner region (semicircle)
        Set objEnts(1) = .AddArc(varCenter, dblRadius * 0.5, 0, dblAngle)
        Set objEnts(2) = .AddLine(objEnts(1).StartPoint, objEnts(1).EndPoint)
        
        '' create the regions
        varRegions = .AddRegion(objEnts)
    End With
    
    '' get new position from user
        varMove = ThisDrawing.Utility.GetPoint(varCenter, vbCr & _
                                               "Pick a new location: ")
    
    '' subtract the inner region from the outer
    varRegions(1).Boolean acSubtraction, varRegions(0)
    
    '' move the composite region to a new location
    varRegions(1).Move varCenter, varMove
End Sub

Public Sub TestAddSolid()
Dim varP1 As Variant
Dim varP2 As Variant
Dim varP3 As Variant
Dim varP4 As Variant
Dim objEnt As AcadSolid
      
On Error Resume Next
    '' ensure that solid fill is enabled
    ThisDrawing.SetVariable "FILLMODE", 1
    
    '' get input from user
    With ThisDrawing.Utility
        varP1 = .GetPoint(, vbCr & "Pick the start point: ")
        varP2 = .GetPoint(varP1, vbCr & "Pick the second point: ")
        varP3 = .GetPoint(varP1, vbCr & "Pick a point opposite the start: ")
        varP4 = .GetPoint(varP3, vbCr & "Pick the last point: ")
    End With
      
    '' draw the entity

    Set objEnt = ThisDrawing.ModelSpace.AddSolid(varP1, varP2, varP3, varP4)
    objEnt.Update
End Sub

Public Sub TestAddText()
Dim varStart As Variant
Dim dblHeight As Double
Dim strText As String
Dim objEnt As AcadText
    
On Error Resume Next
    '' get input from user
    With ThisDrawing.Utility
        varStart = .GetPoint(, vbCr & "Pick the start point: ")
        dblHeight = .GetDistance(varStart, vbCr & "Indicate the height: ")
        strText = .GetString(True, vbCr & "Enter the text: ")
    End With
    
    '' create the text
If ThisDrawing.ActiveSpace = acModelSpace Then
  Set objEnt = ThisDrawing.ModelSpace.AddText(strText, varStart, dblHeight)
Else
  Set objEnt = ThisDrawing.PaperSpace.AddText(strText, varStart, dblHeight)
End If
    objEnt.Update
End Sub

; CHAPTER 9 Creating 3-D Objects 
Public Sub SetViewpoint(Optional Zoom As Boolean = False, _
                    Optional X As Double = 1, _
                    Optional Y As Double = -2, _
                    Optional Z As Double = 1)
  
    Dim dblDirection(2) As Double
    dblDirection(0) = X:  dblDirection(1) = Y:  dblDirection(2) = Z
    
    With ThisDrawing
        .Preferences.ContourLinesPerSurface = 10    ' set surface countours
        .ActiveViewport.Direction = dblDirection    ' assign new direction
        .ActiveViewport = .ActiveViewport           ' force a viewport update
        If Zoom Then .Application.ZoomAll           ' zoomall if requested
    End With
End Sub

Public Sub TestAddBox()
      Dim varPick As Variant
      Dim dblLength As Double
      Dim dblWidth As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint Zoom:=True
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick a corner point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblLength = .GetDistance(varPick, vbCr & "Enter the X length: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblWidth = .GetDistance(varPick, vbCr & "Enter the Y width: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0) + (dblLength / 2)
    dblCenter(1) = varPick(1) + (dblWidth / 2)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddBox(dblCenter, dblLength, _
                                               dblWidth, dblHeight)
    objEnt.Update
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddCone()
      Dim varPick As Variant
      Dim dblRadius As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick the base center point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0)
    dblCenter(1) = varPick(1)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddCone(dblCenter, dblRadius, _
                                                dblHeight)
    objEnt.Update
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddCylinder()
      Dim varPick As Variant
      Dim dblRadius As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick the base center point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0)
    dblCenter(1) = varPick(1)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddCylinder(dblCenter, dblRadius, dblHeight)
    objEnt.Update
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddSphere()
      Dim varPick As Variant
      Dim dblRadius As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick the center point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
    End With
        
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddSphere(varPick, dblRadius)
    objEnt.Update

    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddTorus()
      Dim pntPick As Variant
      Dim pntRadius As Variant
      Dim dblRadius As Double
      Dim dblTube As Double
      Dim objEnt As Acad3DSolid
      Dim intI As Integer
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        pntPick = .GetPoint(, vbCr & "Pick the center point: ")
        .InitializeUserInput 1
        pntRadius = .GetPoint(pntPick, vbCr & "Pick a radius point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblTube = .GetDistance(pntRadius, vbCr & "Enter the tube radius: ")
    End With
        
    '' calculate radius from points
    For intI = 0 To 2
        dblRadius = dblRadius + (pntPick(intI) - pntRadius(intI)) ^ 2
    Next
    dblRadius = Sqr(dblRadius)
        
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddTorus(pntPick, dblRadius, dblTube)
    objEnt.Update
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddWedge()
      Dim varPick As Variant
      Dim dblLength As Double
      Dim dblWidth As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick a base corner point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblLength = .GetDistance(varPick, vbCr & "Enter the base X length: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblWidth = .GetDistance(varPick, vbCr & "Enter the base Y width: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & "Enter the base Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0) + (dblLength / 2)
    dblCenter(1) = varPick(1) + (dblWidth / 2)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddWedge(dblCenter, dblLength, _
                                                 dblWidth, dblHeight)
    objEnt.Update
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddEllipticalCone()
      Dim varPick As Variant
      Dim dblXAxis As Double
      Dim dblYAxis As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick a base center point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblXAxis = .GetDistance(varPick, vbCr & "Enter the X eccentricity: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblYAxis = .GetDistance(varPick, vbCr & "Enter the Y eccentricity: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & "Enter the cone Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0)
    dblCenter(1) = varPick(1)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddEllipticalCone(dblCenter, _
                                        dblXAxis, dblYAxis, dblHeight)
    objEnt.Update
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddEllipticalCylinder()
      Dim varPick As Variant
      Dim dblXAxis As Double
      Dim dblYAxis As Double
      Dim dblHeight As Double
      Dim dblCenter(2) As Double
      Dim objEnt As Acad3DSolid
      
    '' set the default viewpoint
    SetViewpoint
    
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varPick = .GetPoint(, vbCr & "Pick a base center point: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblXAxis = .GetDistance(varPick, vbCr & "Enter the X eccentricity: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblYAxis = .GetDistance(varPick, vbCr & "Enter the Y eccentricity: ")
        .InitializeUserInput 1 + 2 + 4, ""
        dblHeight = .GetDistance(varPick, vbCr & _
                                          "Enter the cylinder Z height: ")
    End With
        
    '' calculate center point from input
    dblCenter(0) = varPick(0)
    dblCenter(1) = varPick(1)
    dblCenter(2) = varPick(2) + (dblHeight / 2)
          
    '' draw the entity
    Set objEnt = ThisDrawing.ModelSpace.AddEllipticalCylinder(dblCenter, _
                                              dblXAxis, dblYAxis, dblHeight)
    objEnt.Update

    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddExtrudedSolid()
      Dim varCenter As Variant
      Dim dblRadius As Double
      Dim dblHeight As Double
      Dim dblTaper As Double
      Dim strInput As String
      Dim varRegions As Variant
      Dim objEnts() As AcadEntity
      Dim objEnt As Acad3DSolid
      Dim varItem As Variant

On Error GoTo Done

    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varCenter = .GetPoint(, vbCr & "Pick the center point: ")
        .InitializeUserInput 1 + 2 + 4
        dblRadius = .GetDistance(varCenter, vbCr & "Indicate the radius: ")
        .InitializeUserInput 1 + 2 + 4
        dblHeight = .GetDistance(varCenter, vbCr & _
                                 "Enter the extrusion height: ")
        
        '' get the taper type
        .InitializeUserInput 1, "Expand Contract None"
        strInput = .GetKeyword(vbCr & _
                               "Extrusion taper [Expand/Contract/None]: ")
        
        '' if none, taper = 0
        If strInput = "None" Then
            dblTaper = 0
            
        '' otherwise, get the taper angle
        Else
            .InitializeUserInput 1 + 2 + 4
            dblTaper = .GetReal("Enter the taper angle ( in degrees): ")
            dblTaper = .AngleToReal(CStr(dblTaper), acDegrees)
            
            
            '' if expanding, negate the angle
            If strInput = "Expand" Then dblTaper = -dblTaper
        End If
    End With
    
    '' draw the entities
    With ThisDrawing.ModelSpace
    
        '' draw the outer region (circle)
        ReDim objEnts(0)
        Set objEnts(0) = .AddCircle(varCenter, dblRadius)
        
        '' create the region
        varRegions = .AddRegion(objEnts)
        '' extrude the solid
        Set objEnt = .AddExtrudedSolid(varRegions(0), dblHeight, dblTaper)
    
        '' update the extruded solid
        objEnt.Update
    End With

Done:
    If Err Then MsgBox Err.Description
    
    '' delete the temporary geometry
    For Each varItem In objEnts
        varItem.Delete
    Next
    For Each varItem In varRegions
        varItem.Delete
    Next
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddExtrudedSolidAlongPath()
      Dim objPath As AcadSpline
      Dim varPick As Variant
      Dim intI As Integer
      Dim dblCenter(2) As Double
      Dim dblRadius As Double
      Dim objCircle As AcadCircle
      Dim objEnts() As AcadEntity
      Dim objShape As Acad3DSolid
      Dim varRegions As Variant
      Dim varItem As Variant

    '' set default viewpoint
    SetViewpoint
    
    '' pick path and calculate shape points
    With ThisDrawing.Utility
        
        '' pick the path
On Error Resume Next
        .GetEntity objPath, varPick, "Pick a Spline for the path"
        If Err Then
            MsgBox "You did not pick a spline"
            Exit Sub
        End If
        objPath.Color = acGreen
        For intI = 0 To 2
            dblCenter(intI) = objPath.FitPoints(intI)
        Next
        .InitializeUserInput 1 + 2 + 4
        dblRadius = .GetDistance(dblCenter, vbCr & "Indicate the radius: ")
    
    End With

    '' draw the circular region, then extrude along path
    With ThisDrawing.ModelSpace
        
        '' draw the outer region (circle)
        ReDim objEnts(0)
        Set objCircle = .AddCircle(dblCenter, dblRadius)
        objCircle.Normal = objPath.StartTangent
        Set objEnts(0) = objCircle
        '' create the region
        varRegions = .AddRegion(objEnts)

        Set objShape = .AddExtrudedSolidAlongPath(varRegions(0), objPath)
        objShape.Color = acRed
    End With

    '' delete the temporary geometry
    For Each varItem In objEnts:  varItem.Delete:  Next
    For Each varItem In varRegions:  varItem.Delete:  Next

    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestAddRevolvedSolid()
      Dim objShape As AcadLWPolyline
      Dim varPick As Variant
      Dim objEnt As AcadEntity
      Dim varPnt1 As Variant
      Dim dblOrigin(2) As Double
      Dim varVec As Variant
      Dim dblAngle As Double
      Dim objEnts() As AcadEntity
      Dim varRegions As Variant
      Dim varItem As Variant
    
    '' set default viewpoint
    SetViewpoint
    
    '' draw the shape and get rotation from user
    With ThisDrawing.Utility
        
        '' pick a shape
On Error Resume Next
        .GetEntity objShape, varPick, "pick a polyline shape"
        If Err Then
            MsgBox "You did not pick the correct type of shape"
            Exit Sub
        End If
On Error GoTo Done

        objShape.Closed = True
        
        '' add pline to region input array
        ReDim objEnts(0)
        Set objEnts(0) = objShape
        
        '' get the axis points
        .InitializeUserInput 1
        varPnt1 = .GetPoint(, vbLf & "Pick an origin of revolution: ")
        .InitializeUserInput 1
        varVec = .GetPoint(dblOrigin, vbLf & _
            "Indicate the axis of revolution: ")
        
        '' get the angle to revolve
        .InitializeUserInput 1
        dblAngle = .GetAngle(, vbLf & "Angle to revolve: ")
    End With

    '' make the region, then revolve it into a solid
    With ThisDrawing.ModelSpace
        
        '' make region from closed pline
        varRegions = .AddRegion(objEnts)
        
        '' revolve solid about axis
        Set objEnt = .AddRevolvedSolid(varRegions(0), varPnt1, varVec, _
                dblAngle)
        objEnt.Color = acRed
    End With

Done:
   If Err Then MsgBox Err.Description
    '' delete the temporary geometry
    For Each varItem In objEnts:  varItem.Delete:  Next
    If Not IsEmpty(varRegions) Then
        For Each varItem In varRegions:  varItem.Delete:  Next
    End If
    
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestBoolean()
      Dim objFirst As Acad3DSolid
      Dim objSecond As Acad3DSolid
      Dim varPick As Variant
      Dim strOp As String
        
On Error Resume Next
    With ThisDrawing.Utility
        
        '' get first solid from user
        .GetEntity objFirst, varPick, vbCr & "Pick a solid to edit: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
        
        '' highlight entity
        objFirst.Highlight True
        objFirst.Update

        '' get second solid from user
        .GetEntity objSecond, varPick, vbCr & "Pick a solid to combine: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
                
        '' exit if they're the same
        If objFirst Is objSecond Then
            MsgBox "You must pick 2 different solids"
            Exit Sub
        End If
                
        '' highlight entity
        objSecond.Highlight True
        objSecond.Update
        
        '' get boolean operation
        .InitializeUserInput 1, "Intersect Subtract Union"
        strOp = .GetKeyword(vbCr & _
                            "Boolean operation [Intersect/Subtract/Union]: ")
        
        '' combine the solids
        Select Case strOp
        Case "Intersect": objFirst.Boolean acIntersection, objSecond
        Case "Subtract": objFirst.Boolean acSubtraction, objSecond
        Case "Union": objFirst.Boolean acUnion, objSecond
        End Select
        
        '' highlight entity
        objFirst.Highlight False
        objFirst.Update
    End With
    
    '' shade the view, and start the interactive orbit command
    ThisDrawing.SendCommand "_shade" & vbCr & "_orbit" & vbCr
End Sub

Public Sub TestInterference()
      Dim objFirst As Acad3DSolid
      Dim objSecond As Acad3DSolid
      Dim objNew As Acad3DSolid
      Dim varPick As Variant
      Dim varNewPnt As Variant
    
On Error Resume Next
    '' set default viewpoint
    SetViewpoint
        
    With ThisDrawing.Utility
        
        '' get first solid from user
        .GetEntity objFirst, varPick, vbCr & "Pick the first solid: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
        
        '' highlight entity
        objFirst.Highlight True
        objFirst.Update

        '' get second solid from user
        .GetEntity objSecond, varPick, vbCr & "Pick the second solid: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
                
        '' exit if they're the same
        If objFirst Is objSecond Then
            MsgBox "You must pick 2 different solids"
            Exit Sub
        End If
                
        '' highlight entity
        objSecond.Highlight True
        objSecond.Update
        
        '' combine the solids
        Set objNew = objFirst.CheckInterference(objSecond, True)
        If objNew Is Nothing Then
            MsgBox "Those solids don't intersect"
        Else
            '' highlight new solid
            objNew.Highlight True
            objNew.Color = acWhite
            objNew.Update
                        
            '' move new solid
            .InitializeUserInput 1
            varNewPnt = .GetPoint(varPick, vbCr & "Pick a new location: ")
            objNew.Move varPick, varNewPnt
        End If
        
        '' dehighlight entities
        objFirst.Highlight False
        objFirst.Update
        objSecond.Highlight False
        objSecond.Update
    End With
    
    '' shade the view, and start the interactive orbit command
    ThisDrawing.SendCommand "_shade" & vbCr & "_orbit" & vbCr
End Sub

Public Sub TestSliceSolid()
      Dim objFirst As Acad3DSolid
      Dim objSecond As Acad3DSolid
      Dim objNew As Acad3DSolid
      Dim varPick As Variant
      Dim varPnt1 As Variant
      Dim varPnt2 As Variant
      Dim varPnt3 As Variant
      Dim strOp As String
      Dim blnOp As Boolean
    
        
On Error Resume Next
    With ThisDrawing.Utility
        
        '' get first solid from user
        .GetEntity objFirst, varPick, vbCr & "Pick a solid to slice: "
        If Err Then
            MsgBox "That is not a 3DSolid"
            Exit Sub
        End If
        
        '' highlight entity
        objFirst.Highlight True
        objFirst.Update

        .InitializeUserInput 1
        varPnt1 = .GetPoint(varPick, vbCr & "Pick first slice point: ")
        .InitializeUserInput 1
        varPnt2 = .GetPoint(varPnt1, vbCr & "Pick second slice point: ")
        .InitializeUserInput 1
        varPnt3 = .GetPoint(varPnt2, vbCr & "Pick last slice point: ")
       
        '' section the solid
        Set objNew = objFirst.SliceSolid(varPnt1, varPnt2, varPnt3, True)
        If objNew Is Nothing Then
            MsgBox "Couldn't slice using those points"
        Else
            '' highlight new solid
            objNew.Highlight False
            objNew.Color = objNew.Color + 1
            objNew.Update
            
            '' move section region to new location
            .InitializeUserInput 1
            varPnt2 = .GetPoint(varPnt1, vbCr & "Pick a new location: ")
            objNew.Move varPnt1, varPnt2
        End If
    End With
    
    '' shade the view
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestSectionSolid()
      Dim objFirst As Acad3DSolid
      Dim objSecond As Acad3DSolid
      Dim objNew As AcadRegion
      Dim varPick As Variant
      Dim varPnt1 As Variant
      Dim varPnt2 As Variant
      Dim varPnt3 As Variant
    
On Error Resume Next

    With ThisDrawing.Utility
        '' get first solid from user
        .GetEntity objFirst, varPick, vbCr & "Pick a solid to section: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
        
        '' highlight entity
        objFirst.Highlight True
        objFirst.Update

        .InitializeUserInput 1
        varPnt1 = .GetPoint(varPick, vbCr & "Pick first section point: ")
        .InitializeUserInput 1
        varPnt2 = .GetPoint(varPnt1, vbCr & "Pick second section point: ")
        .InitializeUserInput 1
        varPnt3 = .GetPoint(varPnt2, vbCr & "Pick last section point: ")
        
        '' section the solid
        Set objNew = objFirst.SectionSolid(varPnt1, varPnt2, varPnt3)
        If objNew Is Nothing Then
            MsgBox "Couldn't section using those points"
        Else
            '' highlight new solid
            objNew.Highlight False
            objNew.Color = acWhite
            objNew.Update
            
            '' move section region to new location
            .InitializeUserInput 1
            varPnt2 = .GetPoint(varPnt1, vbCr & "Pick a new location: ")
            objNew.Move varPnt1, varPnt2
        End If
        
        '' dehighlight entities
        objFirst.Highlight False
        objFirst.Update
    End With
    
    '' shade the view
    ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestMassProperties()
      Dim objEnt As Acad3DSolid
      Dim varPick As Variant
      Dim strMassProperties As String
      Dim varProperty As Variant
      Dim intI As Integer

On Error Resume Next
    '' let user pick a solid
    With ThisDrawing.Utility
        .GetEntity objEnt, varPick, vbCr & "Pick a solid: "
        If Err Then
            MsgBox "That is not an Acad3DSolid"
            Exit Sub
        End If
    End With
    
    '' format mass properties
    With objEnt
        strMassProperties = "Volume: "
        strMassProperties = strMassProperties & vbCr & "    " & .Volume
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Center Of Gravity: "
        For Each varProperty In .Centroid
            strMassProperties = strMassProperties & vbCr & "    "_
                                 & varProperty
        Next
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Moment Of Inertia: "
        For Each varProperty In .MomentOfInertia
            strMassProperties = strMassProperties & vbCr & "    " & _
                                varProperty
        Next
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Product Of Inertia: "
        For Each varProperty In .ProductOfInertia
            strMassProperties = strMassProperties & vbCr & "    " & _
                                varProperty
        Next
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Principal Moments: "
        For Each varProperty In .PrincipalMoments
            strMassProperties = strMassProperties & vbCr & "    " & _
                                varProperty
        Next
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Radii Of Gyration: "
        For Each varProperty In .RadiiOfGyration
            strMassProperties = strMassProperties & vbCr & "    " & _
                                varProperty
        Next
        strMassProperties = strMassProperties & vbCr & vbCr & _
                            "Principal Directions: "
        For intI = 0 To UBound(.PrincipalDirections) / 3
            strMassProperties = strMassProperties & vbCr & "    (" & _
                           .PrincipalDirections((intI - 1) * 3) & ", " & _
                           .PrincipalDirections((intI - 1) * 3 + 1) & "," & _
                           .PrincipalDirections((intI - 1) * 3 + 2) & ")"
        Next
        
    End With
    
    '' highlight entity
    objEnt.Highlight True
    objEnt.Update
    
    '' display properties
    MsgBox strMassProperties, , "Mass Properties"
    
    '' dehighlight entity
    objEnt.Highlight False
    objEnt.Update
End Sub

; CHAPTER 10 Editing Objects 
Public Sub CopyObject()
      Dim objDrawingObject As AcadEntity
      Dim objCopiedObject As Object
      Dim varEntityPickedPoint As Variant
      Dim varCopyPoint As Variant

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
        "Pick an entity to copy: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not pick an object"
        Exit Sub
    End If
    
    'Copy the object
    Set objCopiedObject = objDrawingObject.Copy()
    varCopyPoint = ThisDrawing.Utility.GetPoint(, "Pick point to copy to: ")

    'put the object in its new position
    objCopiedObject.Move varEntityPickedPoint, varCopyPoint
    objCopiedObject.Update
  
End Sub

Public Sub DeleteObject()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
    "Pick an entity to delete: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not pick an object."
        Exit Sub
    End If
    
    'delete the object
    objDrawingObject.Delete 
End Sub

Public Sub ExplodeRegion()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Please pick a region object."
    If objDrawingObject Is Nothing Or _
       objDrawingObject.ObjectName <> "AcDbRegion" Then
        MsgBox "You did not choose a region object."
        Exit Sub
    End If
   
Dim varObjectArray As Variant
Dim strObjectTypes As String
Dim intCount As Integer
        
    varObjectArray = objDrawingObject.Explode
    strObjectTypes = "The region you chose has been exploded " & _ 
             "into the following: " & UBound(varObjectArray) + 1 & " objects:"
    For intCount = 0 To UBound(varObjectArray)
        strObjectTypes = strObjectTypes & vbCrLf & _
                         varObjectArray(intCount).ObjectName
    Next
    MsgBox strObjectTypes  
End Sub

Public Sub ToggleHighlight()
      Dim objSelectionSet As AcadSelectionSet
      Dim objDrawingObject As AcadEntity

    'choose a selection set name that you only use as temporary storage and
    'ensure that it does not currently exist
    On Error Resume Next
    ThisDrawing.SelectionSets("TempSSet").Delete
    Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")

    'ask user to pick entities on the screen
    objSelectionSet.SelectOnScreen
    
    'change the highlight status of each entity selected
    For Each objDrawingObject In objSelectionSet
        objDrawingObject.Highlight True
        objDrawingObject.Update			‘not required for 2006
        MsgBox "Notice that the entity is highlighted"
        objDrawingObject.Highlight False			‘not required for 2006
        objDrawingObject.Update			‘not required for 2006
        MsgBox "Notice that the entity is not highlighted"
    Next
  
    objSelectionSet.Delete
End Sub

Public Sub MirrorObjects()
      Dim objSelectionSet As AcadSelectionSet
      Dim objDrawingObject As AcadEntity
      Dim objMirroredObject As AcadEntity
      Dim varPoint1 As Variant
      Dim varPoint2 As Variant
    
    ThisDrawing.SetVariable "MIRRTEXT", 0

    'choose a selection set name that you only use as temporary storage and
    'ensure that it does not currently exist
    On Error Resume Next
    ThisDrawing.SelectionSets("TempSSet").Delete
    Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")

    'ask user to pick entities on the screen
    ThisDrawing.Utility.Prompt "Pick objects to be mirrored." & vbCrLf
    objSelectionSet.SelectOnScreen
    
'change the highlight status of each entity selected
    varPoint1 = ThisDrawing.Utility.GetPoint(, _
        "Select a point on the mirror axis")
    varPoint2 = ThisDrawing.Utility.GetPoint(varPoint1, _
        "Select a point on the mirror axis")

    For Each objDrawingObject In objSelectionSet
        Set objMirroredObject = objDrawingObject.Mirror(varPoint1, varPoint2)
        objMirroredObject.Update
    Next
  
    objSelectionSet.Delete
End Sub

Public Sub MirrorObjectinXYplane()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant
      Dim objMirroredObject As AcadEntity
      Dim dblPlanePoint1(2) As Double
      Dim dblPlanePoint2(2) As Double
      Dim dblPlanePoint3(2) As Double

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Please an entity to reflect: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If

    'set plane of reflection to be the XY plane
    dblPlanePoint2(0) = 1#
    dblPlanePoint3(1) = 1#
    
    objDrawingObject.Mirror3D dblPlanePoint1, dblPlanePoint2, dblPlanePoint3       
End Sub

Public Sub MoveObjects()
      Dim varPoint1 As Variant
      Dim varPoint2 As Variant
      Dim objSelectionSet As AcadSelectionSet
      Dim objDrawingObject As AcadEntity

    'choose a selection set name that you only use as temporary storage and 
    'ensure that it does not currently exist
    On Error Resume Next
    ThisDrawing.SelectionSets("TempSSet").Delete
    Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")

    'ask user to pick entities on the screen
    objSelectionSet.SelectOnScreen

    varPoint1 = ThisDrawing.Utility.GetPoint(, vbCrLf _
        & "Base point of displacement: ")
    varPoint2 = ThisDrawing.Utility.GetPoint(varPoint1, vbCrLf _
        & "Second point of displacement: ")
  
    'move the selection of entities
    For Each objDrawingObject In objSelectionSet
     objDrawingObject.Move varPoint1, varPoint2
     objDrawingObject.Update
    Next
  
    objSelectionSet.Delete  
End Sub

Public Sub OffsetEllipse()
      Dim objEllipse As AcadEllipse
      Dim varObjectArray As Variant
      Dim dblCenter(2) As Double
      Dim dblMajor(2) As Double

    dblMajor(0) = 100#
    Set objEllipse = ThisDrawing.ModelSpace.AddEllipse(dblCenter, dblMajor, _
                                                       0.5)

    varObjectArray = objEllipse.Offset(50)
    MsgBox "The offset object is a " & varObjectArray(0).ObjectName

    varObjectArray = objEllipse.Offset(-25)
    MsgBox "The offset object is a " & varObjectArray(0).ObjectName  
End Sub

Public Sub RotateObject()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant
      Dim varBasePoint As Variant
      Dim dblRotationAngle As Double
  
On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Please pick an entity to rotate: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object."
        Exit Sub
    End If
    varBasePoint = ThisDrawing.Utility.GetPoint(, _
                   "Enter a base point for the rotation.")
    dblRotationAngle = ThisDrawing.Utility.GetReal( _
                       "Enter the rotation angle in degrees: ")
    
    'convert to radians
    dblRotationAngle = ThisDrawing.Utility. _
                       AngleToReal(CStr(dblRotationAngle), acDegrees)
  
    'Rotate the object
    objDrawingObject.Rotate varBasePoint, dblRotationAngle
    objDrawingObject.Update
End Sub

Public Sub Rotate3DObject()
Dim objDrawingObject As AcadEntity
Dim varEntityPickedPoint As Variant
Dim objMirroredObject As AcadEntity
Dim varAxisPoint1 As Variant
Dim varAxisPoint2 As Variant
Dim dblRotationAxis As Double

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, _
        varEntityPickedPoint, "Please an entity to rotate: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If

   'ask user for axis points and angle of rotation
   varAxisPoint1 = ThisDrawing.Utility.GetPoint(, _
                                "Enter first point of axis of rotation: ")
   varAxisPoint2 = ThisDrawing.Utility.GetPoint(, _
                                "Enter second point of axis of rotation: ")
   dblRotationAxis = ThisDrawing.Utility.GetReal( _
                                "Enter angle of rotation in degrees")
   'convert to radians
   dblRotationAxis = ThisDrawing.Utility.AngleToReal(CStr(dblRotationAxis), _
                                                     acDegrees)
    
   objDrawingObject.Rotate3D varAxisPoint1, varAxisPoint2, dblRotationAxis       
End Sub

Public Sub ScaleObject()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant
      Dim varBasePoint As Variant
      Dim dblScaleFactor As Double
  
On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Please pick an entity to scale: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If
    
    varBasePoint = ThisDrawing.Utility.GetPoint(, _
                   "Pick a base point for the scale:")
    dblScaleFactor = ThisDrawing.Utility.GetReal("Enter the scale factor: ")
    
    'Scale the object
    objDrawingObject.ScaleEntity varBasePoint, dblScaleFactor
    objDrawingObject.Update
End Sub

Public Sub CreatePolarArray()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant
      Dim varArrayCenter As Variant
      Dim lngNumberofObjects As Long
      Dim dblAngletoFill As Double
      Dim varPolarArray As Variant
      Dim intCount As Integer

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                       "Please select an entity to form the basis of a polar array"
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If
    
    varArrayCenter = ThisDrawing.Utility.GetPoint(, _
                     "Pick the center of the array: ")
    lngNumberofObjects = ThisDrawing.Utility.GetInteger( _
                     "Enter total number of objects required in the array: ")
    dblAngletoFill = ThisDrawing.Utility.GetReal( _
          "Enter an angle (in degrees less than 360) over which the array should extend: ")
    dblAngletoFill = ThisDrawing.Utility.AngleToReal _
                            (CStr(dblAngletoFill), acDegrees)
    If dblAngletoFill > 359 Then
      MsgBox “Angle must be less than 360 degrees”, vbCritical
      Exit Sub
    End If
    varPolarArray = objDrawingObject.ArrayPolar(lngNumberofObjects, _
                                            dblAngletoFill, varArrayCenter)

    For intCount = 0 To UBound(varPolarArray)
        varPolarArray(intCount).Color = acRed
        varPolarArray(intCount).Update
    Next intCount
End Sub

Public Sub Create2DRectangularArray()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant
      Dim lngNoRows As Long
      Dim lngNoColumns As Long
      Dim dblDistRows As Long
      Dim dblDistCols As Long
      Dim varRectangularArray As Variant
      Dim intCount As Integer

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
          "Please pick an entity to form the basis of a rectangular array: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If
    
    lngNoRows = ThisDrawing.Utility.GetInteger( _
                             "Enter the required number of rows: ")
    lngNoColumns = ThisDrawing.Utility.GetInteger( _
                             "Enter the required number of columns: ")
    dblDistRows = ThisDrawing.Utility.GetReal( _
                             "Enter the required distance between rows: ")
    dblDistCols = ThisDrawing.Utility.GetReal( _
                             "Enter the required distance between columns: ")
    
    varRectangularArray = objDrawingObject.ArrayRectangular(lngNoRows, _
                             lngNoColumns, 1, dblDistRows, dblDistCols, 0)
    For intCount = 0 To UBound(varRectangularArray)
        varRectangularArray(intCount).Color = acRed
        varRectangularArray(intCount).Update
    Next
End Sub

Public Sub ColorGreen()
Dim objSelectionSet As AcadSelectionSet
Dim objDrawingObject As AcadEntity

    'choose a selection set name that you only use as temporary storage and
    'ensure that it does not currently exist
    On Error Resume Next
    ThisDrawing.SelectionSets("TempSSet").Delete
    Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")

    'ask user to pick entities on the screen
    objSelectionSet.SelectOnScreen
 
    For Each objDrawingObject In objSelectionSet
        objDrawingObject.Color = acGreen
        objDrawingObject.Update
    Next

    objSelectionSet.Delete 
End Sub

Public Sub Example_TrueColor()
    ' This example draws a line and returns the RGB values
    Dim color As AcadAcCmColor
    Set color = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.16")
    Call color.SetRGB(80, 100, 244)
    
    Dim line As AcadLine
    Dim startPoint(0 To 2) As Double
    Dim endPoint(0 To 2) As Double
        
    startPoint(0) = 1#: startPoint(1) = 1#: startPoint(2) = 0#
    endPoint(0) = 5#: endPoint(1) = 5#: endPoint(2) = 0#
        
    Set line = ThisDrawing.ModelSpace.AddLine(startPoint, endPoint)
    ZoomAll
    
    line.TrueColor = color
    Dim retcolor As AcadAcCmColor
    Set retcolor = line.TrueColor
    
    MsgBox "Red = " & retcolor.Red & vbCrLf & _
    "Green = " & retcolor.Green & vbCrLf & _
    "Blue = " & retcolor.Blue
End Sub
Dim fColor As AcadAcCmColor
Set fColor = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.16")
Call color.SetRGB(80, 100, 244) _    
Call fColor.SetColorBookColor("PANTONE Yellow", _
      "Pantone solid colors-uncoated.acb") 
objCircle.TrueColor = fColor

Public Sub ChangeLayer()
      Dim objNewLayer As AcadLayer
      Dim objCircle1 As AcadCircle
      Dim objCircle2 As AcadCircle
      Dim dblCenter1(2) As Double
      Dim dblCenter2(2) As Double
      dblCenter2(0) = 10#

    'reference a layer called "New Layer" is it exists or 
    'add a new layer if it does not
    Set objNewLayer = ThisDrawing.Layers.Add("New Layer")
    
    objNewLayer.Color = acBlue

    ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")
    Set objCircle1 = ThisDrawing.ModelSpace.AddCircle(dblCenter1, 10#)
    Set objCircle2 = ThisDrawing.ModelSpace.AddCircle(dblCenter2, 10#)
    objCircle1.Color = acRed
    objCircle1.Update
    
    objCircle1.Layer = "New Layer"
    objCircle2.Layer = "New Layer"
End Sub

Public Sub ChangeLinetype()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Please pick an object"
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If

    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Pick an entity to change linetype: "
    objDrawingObject.Linetype = "Continuous"
    objDrawingObject.Update
    
End Sub

Public Sub ToggleVisibility()
      Dim objDrawingObject As AcadEntity
      Dim varEntityPickedPoint As Variant

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
                                  "Choose an object to toggle visibility: "
    If objDrawingObject Is Nothing Then
        MsgBox "You did not choose an object"
        Exit Sub
    End If
  
  objDrawingObject.Visible = False
  objDrawingObject.Update
  MsgBox "The object was made invisible!"
  objDrawingObject.Visible = True
  objDrawingObject.Update
  MsgBox "Now it is visible again!"
End Sub

; CHAPTER 11 Dimensions and Annotations . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 231
Public Sub NewDimStyle
Dim objDimStyle As AcadDimStyle

    Set objDimStyle = ThisDrawing.DimStyles.Add("NewDimStyle")
    SetVariable "DIMCLRD", acRed
    SetVariable "DIMCLRE", acBlue
    SetVariable "DIMCLRT", acWhite
    SetVariable "DIMLWD", acLnWtByLwDefault
    objDimStyle.CopyFrom ThisDrawing
End Sub

Public Sub ChangeDimStyle()
Dim objDimension As AcadDimension
Dim varPickedPoint As Variant
Dim objDimStyle As AcadDimStyle
Dim strDimStyles As String
Dim strChosenDimStyle As String

On Error Resume Next
    ThisDrawing.Utility.GetEntity objDimension, varPickedPoint, _
                              "Pick a dimension whose style you wish to set"
    If objDimension Is Nothing Then
        MsgBox "You failed to pick a dimension object"
        Exit Sub
    End If
    
    For Each objDimStyle In ThisDrawing.DimStyles
        strDimStyles = strDimStyles & objDimStyle.Name & vbCrLf
    Next objDimStyle
    strChosenDimStyle = InputBox("Choose one of the following " & _
                        "Dimension styles to apply" & vbCrLf & strDimStyles)
    If strChosenDimStyle = "" Then Exit Sub
    
    objDimension.StyleName = strChosenDimStyle
End Sub 

Public Sub SetActiveDimStyle()
Dim strDimStyles As String
Dim strChosenDimStyle As String
Dim objDimStyle As AcadDimStyle

    For Each objDimStyle In ThisDrawing.DimStyles
        strDimStyles = strDimStyles & objDimStyle.Name & vbCrLf
    Next

    strChosenDimStyle = InputBox("Choose one of the following Dimension " &  _ 
     “styles:" & vbCr & vbCr & strDimStyles, "Existing Dimension style is: " &  _
     ThisDrawing.ActiveDimStyle.Name, ThisDrawing.ActiveDimStyle.Name)
    If strChosenDimStyle = "" Then Exit Sub

On Error Resume Next
    ThisDrawing.ActiveDimStyle = ThisDrawing.DimStyles(strChosenDimStyle)
    If Err Then MsgBox "Dimension style was not recognized"
End Sub

Public Sub Add3PointAngularDimension()
Dim varAngularVertex As Variant
Dim varFirstPoint As Variant
Dim varSecondPoint As Variant
Dim varTextLocation As Variant
Dim objDim3PointAngular As AcadDim3PointAngular
    
    'Define the dimension
    varAngularVertex = ThisDrawing.Utility.GetPoint(, _
                                                 "Enter the center point: ")
    varFirstPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
                                                 "Select first point: ")
    varSecondPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
                                                 "Select second point: ")
    varTextLocation = ThisDrawing.Utility.GetPoint(varAngularVertex, _
                "Pick dimension text location: ")
    
    Set objDim3PointAngular = ThisDrawing.ModelSpace.AddDim3PointAngular( _
        varAngularVertex, varFirstPoint, varSecondPoint, varTextLocation)
    objDim3PointAngular.Update
End Sub

Public Sub AddAlignedDimension()
Dim varFirstPoint As Variant
Dim varSecondPoint As Variant
Dim varTextLocation As Variant
Dim objDimAligned As AcadDimAligned
    'Define the dimension
    varFirstPoint = ThisDrawing.Utility.GetPoint(, "Select first point: ")
    varSecondPoint = ThisDrawing.Utility.GetPoint(varFirstPoint, _
                                                  "Select second point: ")
    varTextLocation = ThisDrawing.Utility.GetPoint(, _
                                           "Pick dimension text location: ")
      
    'Create an aligned dimension
    Set objDimAligned = ThisDrawing.ModelSpace.AddDimAligned(varFirstPoint, _
                                            varSecondPoint, varTextLocation)
    objDimAligned.Update
    
    MsgBox "Now we will change to Engineering units format"
    objDimAligned.UnitsFormat = acDimLEngineering
    objDimAligned.Update
End Sub

Public Sub AddAngularDimension()
Dim varAngularVertex As Variant
Dim varFirstPoint As Variant
Dim varSecondPoint As Variant
Dim varTextLocation As Variant
Dim objDimAngular As AcadDimAngular
  
    'Define the dimension
    varAngularVertex = ThisDrawing.Utility.GetPoint(, _
                                                 "Enter the center point: ")
    varFirstPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
                                                 "Select first point: ")
    varSecondPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
                                                  "Select second point: ")
    varTextLocation = ThisDrawing.Utility.GetPoint(varAngularVertex, _
            "Pick dimension text location: ")
    
    'Create an angular dimension
    Set objDimAngular = ThisDrawing.ModelSpace.AddDimAngular( _
           varAngularVertex, varFirstPoint, varSecondPoint, varTextLocation)
    objDimAngular.AngleFormat = acGrads
    objDimAngular.Update
    MsgBox "Angle measured in GRADS"
      
    objDimAngular.AngleFormat = acDegreeMinuteSeconds
    objDimAngular.TextPrecision = acDimPrecisionFour
    objDimAngular.Update
    MsgBox "Angle measured in Degrees Minutes Seconds"
End Sub

Public Sub AddDiametricDimension()
Dim varFirstPoint As Variant
Dim varSecondPoint As Variant
Dim dblLeaderLength As Double
Dim objDimDiametric As AcadDimDiametric
Dim intOsmode As Integer
    
    'get original object snap settings
    intOsmode = ThisDrawing.GetVariable("osmode")
    ThisDrawing.SetVariable "osmode", 512 ' Near
    
    With ThisDrawing.Utility
        varFirstPoint = .GetPoint(, "Select first point on circle: ")
        ThisDrawing.SetVariable "osmode", 128 ' Per
        varSecondPoint = .GetPoint(varFirstPoint, _
                                   "Select a point opposite the first: ")
        dblLeaderLength = .GetDistance(varFirstPoint, _
                                   "Enter leader length from first point: ")
    End With
    
    Set objDimDiametric = ThisDrawing.ModelSpace.AddDimDiametric( _
                             varFirstPoint, varSecondPoint, dblLeaderLength)
    objDimDiametric.UnitsFormat = acDimLEngineering
    objDimDiametric.PrimaryUnitsPrecision = acDimPrecisionFive
    objDimDiametric.FractionFormat = acNotStacked
    objDimDiametric.Update
    
    'reinstate original object snap settings
    ThisDrawing.SetVariable "osmode", intOsmode
End Sub

Public Sub AddOrdinateDimension()
Dim varBasePoint As Variant
Dim varLeaderEndPoint As Variant
Dim blnUseXAxis As Boolean
Dim strKeywordList As String
Dim strAnswer As String
Dim objDimOrdinate As AcadDimOrdinate

    strKeywordList = "X Y"    
    'Define the dimension
    varBasePoint = ThisDrawing.Utility.GetPoint(, _
                                     "Select ordinate dimension position: ")
    ThisDrawing.Utility.InitializeUserInput 1, strKeywordList
    strAnswer = ThisDrawing.Utility.GetKeyword("Along Which Axis? <X/Y>: ")
    
    If strAnswer = "X" Then
      varLeaderEndPoint = ThisDrawing.Utility.GetPoint(varBasePoint, _
                                      "Select X point for dimension text: ")
      blnUseXAxis = True
      Else
        varLeaderEndPoint = ThisDrawing.Utility.GetPoint(varBasePoint, _
                                      "Select Y point for dimension text: ")
        blnUseXAxis = False
    End If
    
    'Create an ordinate dimension
    Set objDimOrdinate = ThisDrawing.ModelSpace.AddDimOrdinate( _
                               varBasePoint, varLeaderEndPoint, blnUseXAxis)
    objDimOrdinate.TextSuffix = "units"
    objDimOrdinate.Update
End Sub

Public Sub AddRadialDimension()
Dim objUserPickedEntity As Object
Dim varEntityPickedPoint As Variant
Dim varEdgePoint As Variant
Dim dblLeaderLength As Double
Dim objDimRadial As AcadDimRadial
Dim intOsmode As Integer
    
    intOsmode = ThisDrawing.GetVariable("osmode")
    ThisDrawing.SetVariable "osmode", 512 ' Near
    
    'Define the dimension
    On Error Resume Next
    With ThisDrawing.Utility
        .GetEntity objUserPickedEntity, varEntityPickedPoint, _
                   "Pick Arc or Circle:"
        If objUserPickedEntity Is Nothing Then
            MsgBox "You did not pick an entity"
            Exit Sub
        End If
        varEdgePoint = .GetPoint(objUserPickedEntity.Center, _
                                 "Pick edge point")
        dblLeaderLength = .GetReal("Enter leader length from this point: ")
    End With
    
    'Create the radial dimension
    Set objDimRadial = ThisDrawing.ModelSpace.AddDimRadial( _
                  objUserPickedEntity.Center, varEdgePoint, dblLeaderLength)
    objDimRadial.ArrowheadType = acArrowArchTick
    objDimRadial.Update
    'reinstate original setting
    ThisDrawing.SetVariable "osmode", intOsmode
End Sub

Public Sub AddRotatedDimension()
Dim varFirstPoint As Variant
Dim varSecondPoint As Variant
Dim varTextLocation As Variant
Dim strRotationAngle As String
Dim objDimRotated As AcadDimRotated
    
    'Define the dimension
    With ThisDrawing.Utility
      varFirstPoint = .GetPoint(, "Select first point: ")
      varSecondPoint = .GetPoint(varFirstPoint, "Select second point: ")
      varTextLocation = .GetPoint(, "Pick dimension text location: ")
      strRotationAngle = .GetString(False, "Enter rotation angle in degrees")
    End With
    'Create a rotated dimension
    Set objDimRotated = ThisDrawing.ModelSpace.AddDimRotated(varFirstPoint, _
      varSecondPoint, varTextLocation, _
      ThisDrawing.Utility.AngleToReal(strRotationAngle, acDegrees))
    objDimRotated.DecimalSeparator = ","
    objDimRotated.Update
End Sub

Public Sub CreateTolerance()
Dim strToleranceText As String
Dim varInsertionPoint As Variant
Dim varTextDirection As Variant
Dim intI As Integer
Dim objTolerance As AcadTolerance

  strToleranceText = InputBox("Please enter the text for the tolerance")
  varInsertionPoint = ThisDrawing.Utility.GetPoint(, _
                       "Please enter the insertion point for the tolerance")
  varTextDirection = ThisDrawing.Utility.GetPoint(varInsertionPoint, _
                               "Please enter a direction for the tolerance")
    
  For intI = 0 To 2
    varTextDirection(intI) = varTextDirection(intI) - varInsertionPoint(intI)
  Next
    
  Set objTolerance = ThisDrawing.ModelSpace.AddTolerance(strToleranceText, _
                                        varInsertionPoint, varTextDirection)

End Sub 

Public Sub AddTextStyle
Dim objTextStyle As AcadTextStyle

    Set objTextStyle = ThisDrawing.TextStyles.Add("Bold Greek Symbols")
    objTextStyle.SetFont "Symbol", True, False, 0, 0
End Sub

Public Sub GetTextSettings()
Dim objTextStyle As AcadTextStyle
Dim strTextStyleName As String
Dim strTextStyles As String
Dim strTypeFace As String
Dim blnBold As Boolean
Dim blnItalic As Boolean
Dim lngCharacterSet As Long
Dim lngPitchandFamily As Long
Dim strText As String
    
    ' Get the name of each text style in the drawing
    For Each objTextStyle In ThisDrawing.TextStyles
        strTextStyles = strTextStyles & vbCr & objTextStyle.Name
    Next
    ' Ask the user to select the Text Style to look at
    strTextStyleName = InputBox("Please enter the name of the TextStyle " & _
                "whose setting you would like to see" & vbCr & _
                strTextStyles,"TextStyles", ThisDrawing.ActiveTextStyle.Name)
    ' Exit the program if the user input was cancelled or empty
    If strTextStyleName = "" Then Exit Sub

On Error Resume Next
    Set objTextStyle = ThisDrawing.TextStyles(strTextStyleName)
    ' Check for existence the text style
    If objTextStyle Is Nothing Then
        MsgBox "This text style does not exist"
        Exit Sub
    End If
    
    ' Get the Font properties
    objTextStyle.GetFont strTypeFace, blnBold, blnItalic, lngCharacterSet, _
                lngPitchandFamily
    ' Check for Type face
    If strTypeFace = "" Then  ' No True type
        MsgBox "Text Style: " & objTextStyle.Name & vbCr & _
                "Using file font: " & objTextStyle.fontFile, _
                vbInformation, "Text Style: " & objTextStyle.Name
    
        Else
        ' True Type font info
       strText = "The text style: " & strTextStyleName & " has " & vbCrLf & _
                 "a " & strTypeFace & " type face"
       If blnBold Then strText = strText & vbCrLf & " and is bold"
       If blnItalic Then strText = strText & vbCrLf & " and is italicized"
       MsgBox strText & vbCr & "Using file font: " & objTextStyle.fontFile, _
              vbInformation, "Text Style: " & objTextStyle.Name
    End If
End Sub

Public Sub SetFontFile
Dim objTextStyle As AcadTextStyle

    Set objTextStyle = ThisDrawing.TextStyles.Add("Roman")
    objTextStyle.fontFile = "romand.shx"
End Sub 

Public Sub ChangeTextStyle()
Dim strTextStyles As String
Dim objTextStyle As AcadTextStyle
Dim objLayer As AcadLayer
Dim strLayerName As String
Dim strStyleName As String
Dim objAcadObject As AcadObject
    
On Error Resume Next
    For Each objTextStyle In ThisDrawing.TextStyles
        strTextStyles = strTextStyles & vbCr & objTextStyle.Name
    Next
    strStyleName = InputBox("Enter name of style to apply:" & vbCr & _
              strTextStyles, "TextStyles", ThisDrawing.ActiveTextStyle.Name)
    Set objTextStyle = ThisDrawing.TextStyles(strStyleName)
    If objTextStyle Is Nothing Then
        MsgBox "Style does not exist"
        Exit Sub
    End If

    For Each objAcadObject In ThisDrawing.ModelSpace
        If objAcadObject.ObjectName = "AcDbMText" Or _
            objAcadObject.ObjectName = "AcDbText" Then
            objAcadObject.StyleName = strStyleName
            objAcadObject.Update
        End If
    Next
End Sub

Public Sub SetDefaultTextStyle()
Dim strTextStyles As String
Dim objTextStyle As AcadTextStyle
Dim strTextStyleName As String

    For Each objTextStyle In ThisDrawing.TextStyles
        strTextStyles = strTextStyles & vbCr & objTextStyle.Name
    Next
    strTextStyleName = InputBox("Enter name of style to apply:" & vbCr & _
              strTextStyles, "TextStyles", ThisDrawing.ActiveTextStyle.Name)
    If strTextStyleName = "" Then Exit Sub

On Error Resume Next
    Set objTextStyle = ThisDrawing.TextStyles(strTextStyleName)
    If objTextStyle Is Nothing Then
        MsgBox "This text style does not exist"
        Exit Sub
    End If
    
    ThisDrawing.ActiveTextStyle = objTextStyle
End Sub

Public Sub CreateStraightLeaderWithNote()
Dim dblPoints(5) As Double
Dim varStartPoint As Variant
Dim varEndPoint As Variant
Dim intLeaderType As Integer
Dim objAcadLeader As AcadLeader
Dim objAcadMtext As AcadMText
Dim strMtext As String
Dim intI As Integer
    
    intLeaderType = acLineWithArrow
    varStartPoint = ThisDrawing.Utility.GetPoint(, _
                                              "Select leader start point: ")
    varEndPoint = ThisDrawing.Utility.GetPoint(varStartPoint, _
                                               "Select leader end point: ")
    
    For intI = 0 To 2
      dblPoints(intI) = varStartPoint(intI)
      dblPoints(intI + 3) = varEndPoint(intI)
    Next
    
    strMtext = InputBox("Notes:", "Leader Notes")
    If strMtext = "" Then Exit Sub
    ' Create the text for the leader
    Set objAcadMtext = ThisDrawing.ModelSpace.AddMText(varEndPoint, _
    Len(strMtext) * ThisDrawing.GetVariable("dimscale"), strMtext)
    ' Flip the alignment direction of the text
    If varEndPoint(0) > varStartPoint(0) Then
        objAcadMtext.AttachmentPoint = acAttachmentPointMiddleLeft
    Else
        objAcadMtext.AttachmentPoint = acAttachmentPointMiddleRight
    End If
    objAcadMtext.InsertionPoint = varEndPoint
    
    'Create the leader object
    Set objAcadLeader = ThisDrawing.ModelSpace.AddLeader(dblPoints, _
                                                objAcadMtext, intLeaderType)
    objAcadLeader.Update
End Sub

; CHAPTER 12 Selection Sets and Groups . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 259
Public Sub TestAddSelectionSet()
      Dim objSS As AcadSelectionSet
      Dim strName As String
    
On Error Resume Next
    '' get a name from user
    strName = InputBox("Enter a new selection set name: ")
    If "" = strName Then Exit Sub
            
    '' create it
    Set objSS = ThisDrawing.SelectionSets.Add(strName)
    
    '' check if it was created
    If objSS Is Nothing Then
        MsgBox "Unable to Add '" & strName & "'"
    Else
        MsgBox "Added selection set '" & objSS.Name & "'"
    End If
End Sub

Public Sub ListSelectionSets()
      Dim objSS As AcadSelectionSet
      Dim strSSList As String

    For Each objSS In ThisDrawing.SelectionSets
        strSSList = strSSList & vbCr & objSS.Name
    Next
    MsgBox strSSList, , "List of Selection Sets"
End Sub

Public Sub TestSelect()
      Dim objSS As AcadSelectionSet
      Dim varPnt1 As Variant
      Dim varPnt2 As Variant
      Dim strOpt As String
      Dim lngMode As Long

On Error GoTo Done
    With ThisDrawing.Utility
    
        '' get input for mode
        .InitializeUserInput 1, "Window Crossing Previous Last All"
        strOpt = .GetKeyword(vbCr & _
        "Select [Window/Crossing/Previous/Last/All]: ")
        
        '' convert keyword into mode
        Select Case strOpt
        Case "Window":  lngMode = acSelectionSetWindow
        Case "Crossing":  lngMode = acSelectionSetCrossing
        Case "Previous":  lngMode = acSelectionSetPrevious
        Case "Last":  lngMode = acSelectionSetLast
        Case "All":  lngMode = acSelectionSetAll
        End Select
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectSS")
                
        '' if it's window or crossing, get the points
        If "Window" = strOpt Or "Crossing" = strOpt Then
            
            '' get first point
            .InitializeUserInput 1
            varPnt1 = .GetPoint(, vbCr & "Pick the first corner: ")
            
            '' get corner, using dashed lines if crossing
            .InitializeUserInput 1 + IIf("Crossing" = strOpt, 32, 0)
            varPnt2 = .GetCorner(varPnt1, vbCr & "Pick other corner: ")
            
            '' select entities using points
            objSS.Select lngMode, varPnt1, varPnt2
        Else
        
            '' select entities using mode
            objSS.Select lngMode
        End If
            
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .GetString False, vbCr & "Enter to continue"
        
        '' dehighlight the entities
        objSS.Highlight False
    
    End With
        
Done:

    '' if the selectionset was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestSelectionSetFilter()
      Dim objSS As AcadSelectionSet
      Dim intCodes(0) As Integer
      Dim varCodeValues(0) As Variant
      Dim strName As String
        
On Error GoTo Done

    With ThisDrawing.Utility
        strName = .GetString(True, vbCr & "Layer name to filter: ")
        If "" = strName Then Exit Sub
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectionSetFilter")
                
        '' set the code for layer
        intCodes(0) = 8
        
        '' set the value specified by user
        varCodeValues(0) = strName
        
        '' filter the objects
        objSS.Select acSelectionSetAll, , , intCodes, varCodeValues
                
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .Prompt vbCr & objSS.Count & " entities selected"
        .GetString False, vbLf & "Enter to continue "
        
        '' dehighlight the entities
        objSS.Highlight False
    End With
    
Done:
    
    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestSelectionSetOperator()
      Dim objSS As AcadSelectionSet
      Dim intCodes() As Integer
      Dim varCodeValues As Variant
      Dim strName As String
        
On Error GoTo Done
    With ThisDrawing.Utility
        strName = .GetString(True, vbCr & "Layer name to exclude: ")
        If "" = strName Then Exit Sub
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectionSetOperator")
        
        '' using 9 filters
        ReDim intCodes(9):  ReDim varCodeValues(9)
        
        '' set codes and values - indented for clarity
        intCodes(0) = -4:  varCodeValues(0) = "<and"
        intCodes(1) = -4:    varCodeValues(1) = "<or"
        intCodes(2) = 0:       varCodeValues(2) = "line"
        intCodes(3) = 0:       varCodeValues(3) = "arc"
        intCodes(4) = 0:       varCodeValues(4) = "circle"
        intCodes(5) = -4:    varCodeValues(5) = "or>"
        intCodes(6) = -4:    varCodeValues(6) = "<not"
        intCodes(7) = 8:       varCodeValues(7) = strName
        intCodes(8) = -4:    varCodeValues(8) = "not>"
        intCodes(9) = -4:  varCodeValues(9) = "and>"
        
        '' filter the objects
        objSS.Select acSelectionSetAll, , , intCodes, varCodeValues
                
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .Prompt vbCr & objSS.Count & " entities selected"
        .GetString False, vbLf & "Enter to continue "
        
        '' dehighlight the entities
        objSS.Highlight False
    End With
    
Done:
    
    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestSelectOnScreen()
      Dim objSS As AcadSelectionSet
    
      On Error GoTo Done
    With ThisDrawing.Utility
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectOnScreen")
                
        '' let user select entities interactively
        objSS.SelectOnScreen
        
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .Prompt vbCr & objSS.Count & " entities selected"
        .GetString False, vbLf & "Enter to continue "
        
        '' dehighlight the entities
        objSS.Highlight False
    
    End With
        
Done:

    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestSelectAtPoint()
      Dim varPick As Variant
      Dim objSS As AcadSelectionSet
    
      On Error GoTo Done

    With ThisDrawing.Utility
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectAtPoint")
                
        '' get a point of selection from the user
        varPick = .GetPoint(, vbCr & "Select entities at a point: ")
                
        '' let user select entities interactively
        objSS.SelectAtPoint varPick
        
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .Prompt vbCr & objSS.Count & " entities selected"
        .GetString False, vbLf & "Enter to continue "
        
        '' dehighlight the entities
        objSS.Highlight False
    
    End With
        
Done:

    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestSelectByPolygon()
      Dim objSS As AcadSelectionSet
      Dim strOpt As String
      Dim lngMode As Long
      Dim varPoints As Variant
    
      On Error GoTo Done
    With ThisDrawing.Utility
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectByPolygon1")
                
        '' get the mode from the user
        .InitializeUserInput 1, "Fence Window Crossing"
        strOpt = .GetKeyword(vbCr & "Select by [Fence/Window/Crossing]: ")
                
        '' convert keyword into mode
        Select Case strOpt
        	Case "Fence":  lngMode = acSelectionSetFence
        	Case "Window":  lngMode = acSelectionSetWindowPolygon
        	Case "Crossing":  lngMode = acSelectionSetCrossingPolygon
        End Select
        
        '' let user digitize points
        varPoints = InputPoints()
        
        '' select entities using mode and points specified
        objSS.SelectByPolygon lngMode, varPoints
        
        '' highlight the selected entities
        objSS.Highlight True
    
        '' pause for the user
        .Prompt vbCr & objSS.Count & " entities selected"
        .GetString False, vbLf & "Enter to continue "
        
        '' dehighlight the entities
        objSS.Highlight False
    
    End With
        
Done:

    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Function InputPoints() As Variant
      Dim varStartPoint As Variant
      Dim varNextPoint As Variant
      Dim varWCSPoint As Variant
      Dim lngLast As Long
      Dim dblPoints() As Double
    
      On Error Resume Next
    '' get first points from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varStartPoint = .GetPoint(, vbLf & "Pick the start point: ")
    
        '' setup initial point
        ReDim dblPoints(2)
        dblPoints(0) = varStartPoint(0)
        dblPoints(1) = varStartPoint(1)
        dblPoints(2) = varStartPoint(2)
        varNextPoint = varStartPoint
        
        '' append vertexes in a loop
        Do
            '' translate picked point to UCS for basepoint below
            varWCSPoint = .TranslateCoordinates(varNextPoint, acWorld, _
                                                acUCS, True)
            
            '' get user point for new vertex, use last pick as basepoint
            varNextPoint = .GetPoint(varWCSPoint, vbCr & _
                                     "Pick another point <exit>: ")
            
            '' exit loop if no point picked
            If Err Then Exit Do
            
            '' get the upper bound
            lngLast = UBound(dblPoints)
            
            '' expand the array
            ReDim Preserve dblPoints(lngLast + 3)
                        
            '' add the new point
            dblPoints(lngLast + 1) = varNextPoint(0)
            dblPoints(lngLast + 2) = varNextPoint(1)
            dblPoints(lngLast + 3) = varNextPoint(2)
        Loop
    End With
    
    '' return the points
    InputPoints = dblPoints
End Function

Public Sub TestSelectAddRemoveClear()
      Dim objSS As AcadSelectionSet
      Dim objSStmp As AcadSelectionSet
      Dim strType As String
      Dim objEnts() As AcadEntity
      Dim intI As Integer
   
      On Error Resume Next

    With ThisDrawing.Utility
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("ssAddRemoveClear")
        If Err Then GoTo Done
                
        '' create a new temporary selection
        Set objSStmp = ThisDrawing.SelectionSets.Add("ssAddRemoveClearTmp")
        If Err Then GoTo Done
        
        '' loop until the user has finished
        Do
            '' clear any pending errors
            Err.Clear
        
            '' get input for type
            .InitializeUserInput 1, "Add Remove Clear Exit"
            strType = .GetKeyword(vbCr & "Select [Add/Remove/Clear/Exit]: ")
                
            '' branch based on input
            If "Exit" = strType Then
            
                '' exit if requested
                Exit Do
                
            ElseIf "Clear" = strType Then
                
                '' dehighlight the main selection
                objSS.Highlight False
                
                '' clear the main set
                objSS.Clear
            
            '' otherwise, we're adding/removing
            Else
            
                '' clear the temporary selection
                objSStmp.Clear
        
                objSStmp.SelectOnScreen
                '' highlight the temporary selection
                objSStmp.Highlight True
                
                '' convert temporary selection to array

                '' resize the entity array to the selection size
                ReDim objEnts(objSStmp.Count - 1)
    
                '' copy entities from the selection to entity array
                For intI = 0 To objSStmp.Count - 1
                    Set objEnts(intI) = objSStmp(intI)
                Next
    
                '' add/remove items from main selection using entity array
                If "Add" = strType Then
                    objSS.AddItems objEnts
                Else
                    objSS.RemoveItems objEnts
                End If
                                    
                '' dehighlight the temporary selection
                objSStmp.Highlight False
                
                '' highlight the main selection
                objSS.Highlight True
            End If
        Loop
    End With
        
Done:

    '' if the selections were created, delete them
    If Not objSS Is Nothing Then
               
        '' dehighlight the entities
        objSS.Highlight False

        '' delete the main selection
        objSS.Delete
    End If
    
    If Not objSStmp Is Nothing Then
    
        '' delete the temporary selection
        objSStmp.Delete
    End If
End Sub

Public Sub TestSelectErase()
      Dim objSS As AcadSelectionSet
    
      On Error GoTo Done
    With ThisDrawing.Utility
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestSelectErase")
                
        '' let user select entities interactively
        objSS.SelectOnScreen
        
        '' highlight the selected entities
        objSS.Highlight True
    
        '' erase the selected entities
        objSS.Erase
        
        '' prove that the selection is empty (but still viable)
        .Prompt vbCr & objSS.Count & " entities selected"
        
    End With
        
Done:

    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Private Sub AcadDocument_SelectionChanged()
      Dim objSS As AcadSelectionSet
      Dim dblStart As Double
    
    '' get the pickfirst selection from drawing
    Set objSS = ThisDrawing.PickfirstSelectionSet
    
    '' highlight the selected entities
    objSS.Highlight True
    MsgBox "There are " & objSS.Count & " objects in selection set: " & objSS.Name
    '' delay for 1/2 second
    dblStart = Timer
    Do While Timer < dblStart + 0.5
    Loop
        
    '' dehighlight the selected entities
    objSS.Highlight False
End Sub

Public Sub TestAddGroup()
      Dim objGroup As AcadGroup
      Dim strName As String
    
      On Error Resume Next
    '' get a name from user
    strName = InputBox("Enter a new group name: ")
    If "" = strName Then Exit Sub
            
    Set objGroup = ThisDrawing.Groups.Item(strName)
    '' create it
    If Not objGroup Is Nothing Then
        MsgBox "Group already exists"
        Exit Sub
    End If
    
    Set objGroup = ThisDrawing.Groups.Add(strName)
   
    '' check if it was created
    If objGroup Is Nothing Then
        MsgBox "Unable to Add '" & strName & "'"
    Else
        MsgBox "Added group '" & objGroup.Name & "'"
    End If
End Sub

Public Sub ListGroups()
      Dim objGroup As AcadGroup
      Dim strGroupList As String

    For Each objGroup In ThisDrawing.Groups
        strGroupList = strGroupList & vbCr & objGroup.Name
    Next
    MsgBox strGroupList, vbOKOnly, "List of Groups"
End Sub

Public Sub TestGroupAppendRemove()
      Dim objSS As AcadSelectionSet
      Dim objGroup As AcadGroup
      Dim objEnts() As AcadEntity
      Dim strName As String
      Dim strOpt As String
      Dim intI As Integer
 
      On Error Resume Next
    '' set pickstyle to NOT select groups
    ThisDrawing.SetVariable "Pickstyle", 2
    
    With ThisDrawing.Utility
        '' get group name from user
        strName = .GetString(True, vbCr & "Group name: ")
        If Err Or "" = strName Then GoTo Done
        
        '' get the existing group or add new one
        Set objGroup = ThisDrawing.Groups.Add(strName)
        
        '' pause for the user
        .Prompt vbCr & "Group contains: " & objGroup.Count & " entities" & _
                vbCrLf
        
        '' get input for mode
        .InitializeUserInput 1, "Append Remove"
        strOpt = .GetKeyword(vbCr & "Option [Append/Remove]: ")
        If Err Then GoTo Done
        
        '' create a new selectionset
        Set objSS = ThisDrawing.SelectionSets.Add("TestGroupAppendRemove")
        If Err Then GoTo Done
        
        '' get a selection set from user
        objSS.SelectOnScreen
        
        '' convert selection set to array
        '' resize the entity array to the selection size
        ReDim objEnts(objSS.Count - 1)
    
        '' copy entities from the selection to entity array
        For intI = 0 To objSS.Count - 1
            Set objEnts(intI) = objSS(intI)
        Next
    
        '' append or remove entities based on input
        If "Append" = strOpt Then
            objGroup.AppendItems objEnts
        Else
            objGroup.RemoveItems objEnts
        End If
    
        '' pause for the user
        .Prompt vbCr & "Group contains: " & objGroup.Count & " entities"
        
        '' dehighlight the entities
        objSS.Highlight False
    End With
        
Done:
    If Err Then MsgBox "Error occurred: " & Err.Description
    '' if the selection was created, delete it
    If Not objSS Is Nothing Then
        objSS.Delete
    End If
End Sub

Public Sub TestGroupDelete()
      Dim objGroup As AcadGroup
      Dim strName As String
    
      On Error Resume Next
    With ThisDrawing.Utility
    
        strName = .GetString(True, vbCr & "Group name: ")
        If Err Or "" = strName Then Exit Sub
        
        '' get the existing group
        Set objGroup = ThisDrawing.Groups.Item(strName)
        If Err Then
            .Prompt vbCr & "Group does not exist "
            Exit Sub
        End If
        
        '' delete the group
        objGroup.Delete
        If Err Then
            .Prompt vbCr & "Error deleting group "
            Exit Sub
        End If
    
        '' pause for the user
        .Prompt vbCr & "Group deleted"
    End With
End Sub

; CHAPTER 13 Blocks, Attributes, and External References . . . . . . . . . . . . . . . . . . . 285
Public Sub ListBlocks()
Dim objBlock As AcadBlock
Dim strBlockList As String

    strBlockList = "List of blocks: "
    
    For Each objBlock In ThisDrawing.Blocks
        strBlockList = strBlockList & vbCr & objBlock.Name
    Next
    
    MsgBox strBlockList
End Sub

Public Sub AddBlock()
Dim dblOrigin(2) As Double
Dim objBlock As AcadBlock
Dim strName As String
    
    '' get a name from user
    strName = InputBox("Enter a new block name: ")
    If "" = strName Then Exit Sub          ' exit if no old name
 
    '' set the origin point
    dblOrigin(0) = 0:  dblOrigin(1) = 0:   dblOrigin(2) = 0
    
    ''check if block already exists
On Error Resume Next
    Set objBlock = ThisDrawing.Blocks.Item(strName)
    If Not objBlock Is Nothing Then
        MsgBox "Block already exists"
        Exit Sub
    End If
    
    '' create the block
    Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, strName)
    
    '' then add entities (circle)
    objBlock.AddCircle dblOrigin, 10    
End Sub

Public Sub TestCopyObjects()
Dim objSS As AcadSelectionSet
Dim varBase As Variant
Dim objBlock As AcadBlock
Dim strName As String
Dim strErase As String
Dim varEnt As Variant
Dim objSourceEnts() As Object
Dim varDestEnts As Variant
Dim dblOrigin(2) As Double
Dim intI As Integer

    'choose a selection set name that you only use as temporary storage and
    'ensure that it does not currently exist
On Error Resume Next
    ThisDrawing.SelectionSets.Item("TempSSet").Delete
    Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
    objSS.SelectOnScreen
    
    '' get the other user input
    With ThisDrawing.Utility
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "Enter a block name: ")
        .InitializeUserInput 1
        varBase = .GetPoint(, vbCr & "Pick a base point: ")
        .InitializeUserInput 1, "Yes No"
        strErase = .GetKeyword(vbCr & "Erase originals [Yes/No]? ")
    End With
        
    '' set WCS origin
    dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
    
    '' create the block
    Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, strName)
    
    '' put selected entities into an array for CopyObjects
    ReDim objSourceEnts(objSS.Count - 1)
    For intI = 0 To objSS.Count - 1
        Set objSourceEnts(intI) = objSS(intI)
    Next
    
    '' copy the entities into block
    varDestEnts = ThisDrawing.CopyObjects(objSourceEnts, objBlock)
    
    '' move copied entities so that base point becomes origin
    For Each varEnt In varDestEnts
        varEnt.Move varBase, dblOrigin
    Next
    
    '' if requested, erase the originals
    If strErase = "Yes" Then
        objSS.Erase
    End If
    
    '' we're done - prove that we did it
    ThisDrawing.SendCommand "._-insert" & vbCr & strName & vbCr

    '' clean up selection set
    objSS.Delete
End Sub

Public Sub RenameBlock()
Dim strName As String
Dim objBlock As AcadBlock
    
On Error Resume Next                   ' handle exceptions inline
    strName = InputBox("Original Block name: ")
    If "" = strName Then Exit Sub          ' exit if no old name
    
    Set objBlock = ThisDrawing.Blocks.Item(strName)
    If objBlock Is Nothing Then            ' exit if not found
        MsgBox "Block '" & strName & "' not found"
        Exit Sub
    End If
    
    strName = InputBox("New Block name: ")
    If "" = strName Then Exit Sub          ' exit if no new name
    
    objBlock.Name = strName                ' try and change name
    If Err Then                            ' check if it worked
        MsgBox "Unable to rename block: " & vbCr & Err.Description
    Else
        MsgBox "Block renamed to '" & strName & "'"
    End If
End Sub

Public Sub DeleteBlock()
Dim strName As String
Dim objBlock As AcadBlock
    
On Error Resume Next                   ' handle exceptions inline
    strName = InputBox("Block name to delete: ")
    If "" = strName Then Exit Sub          ' exit if no old name

    Set objBlock = ThisDrawing.Blocks.Item(strName)
    If objBlock Is Nothing Then            ' exit if not found
        MsgBox "Block '" & strName & "' not found"
        Exit Sub
    End If

    objBlock.Delete                        ' try to delete it
    If Err Then                            ' check if it worked
        MsgBox "Unable to delete Block: " & vbCr & Err.Description
    Else
        MsgBox "Block '" & strName & "' deleted"
    End If
End Sub

Public Sub TestInsertBlock()
Dim strName As String
Dim varInsertionPoint As Variant
Dim dblX As Double
Dim dblY As Double
Dim dblZ As Double
Dim dblRotation As Double
          
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "Block or file name: ")
        .InitializeUserInput 1
        varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
        .InitializeUserInput 1 + 2
        dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
        .InitializeUserInput 1 + 2
        dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
        .InitializeUserInput 1 + 2
        dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
        .InitializeUserInput 1
        dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
    End With
    
    '' create the object
On Error Resume Next

    ThisDrawing.ModelSpace.InsertBlock varInsertionPoint, strName, dblX, _
                                       dblY, dblZ, dblRotation
    If Err Then MsgBox "Unable to insert this block."    
End Sub

Private Sub CommandButton1_Click()
    Dim objBlockRef As AcadBlockReference
    Dim varInsertionPoint As Variant
    Dim dblX As Double
    Dim dblY As Double
    Dim dblZ As Double
    Dim dblRotation As Double
          
    '' get input from user
    dlgOpenFile.Filter = "AutoCAD Blocks (*.DWG) | *.dwg"
    dlgOpenFile.InitDir = Application.Path     
    dlgOpenFile.ShowOpen

    If dlgOpenFile.FileName = "" Then Exit Sub
    Me.Hide
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
        .InitializeUserInput 1 + 2
        dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
        .InitializeUserInput 1 + 2
        dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
        .InitializeUserInput 1 + 2
        dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
        .InitializeUserInput 1
        dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
    End With
    
    '' create the object
On Error Resume Next
    Set objBlockRef = ThisDrawing.ModelSpace.InsertBlock _
        (varInsertionPoint, dlgOpenFile.FileName, dblX, _
        dblY, dblZ, dblRotation)
    
    If Err Then
        MsgBox "Unable to insert this block"
        Exit Sub
    End If
    objBlockRef.Update
    Me.Show
End Sub

Public Sub TestExplode()
    Dim objBRef As AcadBlockReference
    Dim varPick As Variant
    Dim varNew As Variant
    Dim varEnts As Variant
    Dim intI As Integer

On Error Resume Next
    '' get an entity and new point from user
    With ThisDrawing.Utility
        .GetEntity objBRef, varPick, vbCr & "Pick a block reference: "
        If Err Then Exit Sub
        varNew = .GetPoint(varPick, vbCr & "Pick a new location: ")
        If Err Then Exit Sub
    End With
    
    '' explode the blockref
    varEnts = objBRef.Explode
    If Err Then
        MsgBox "Error has occurred: " & Err.Description
        Exit Sub
    End If
        
    '' move resulting entities to new location
    For intI = 0 To UBound(varEnts)
        varEnts(intI).Move varPick, varNew
    Next
End Sub

Public Sub TestWBlock()
Dim objSS As AcadSelectionSet
Dim varBase As Variant
Dim dblOrigin(2) As Double
Dim objEnt As AcadEntity
Dim strFilename As String
    
    'choose a selection set name that you only use as temporary storage and
    'ensure that it does not currently exist
On Error Resume Next
    ThisDrawing.SelectionSets("TempSSet").Delete
    Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
    objSS.SelectOnScreen

    With ThisDrawing.Utility
        .InitializeUserInput 1
        strFilename = .GetString(True, vbCr & "Enter a filename: ")
        .InitializeUserInput 1
        varBase = .GetPoint(, vbCr & "Pick a base point: ")
    End With
    
    '' WCS origin
    dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
    
    '' move selection to the origin
    For Each objEnt In objSS
        objEnt.Move varBase, dblOrigin
    Next
    
    '' wblock selection to filename
    ThisDrawing.Wblock strFilename, objSS
        
    '' move selection back
    For Each objEnt In objSS
        objEnt.Move dblOrigin, varBase
    Next
    
    '' clean up selection set
    objSS.Delete
End Sub

Public Sub TestAddMInsertBlock()
Dim strName As String
Dim varInsertionPoint As Variant
Dim dblX As Double
Dim dblY As Double
Dim dblZ As Double
Dim dR As Double
Dim lngNRows As Long
Dim lngNCols As Long
Dim dblSRows As Double
Dim dblSCols As Double
          
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "Block or file name: ")
        .InitializeUserInput 1
        varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
        .InitializeUserInput 1 + 2
        dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
        .InitializeUserInput 1 + 2
        dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
        .InitializeUserInput 1 + 2
        dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
        .InitializeUserInput 1
        dR = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
        .InitializeUserInput 1 + 2 + 4
        lngNRows = .GetInteger(vbCr & "Number of rows: ")
        .InitializeUserInput 1 + 2 + 4
        lngNCols = .GetInteger(vbCr & "Number of columns: ")
        .InitializeUserInput 1 + 2
        dblSRows = .GetDistance(varInsertionPoint, vbCr & "Row spacing: ")
        .InitializeUserInput 1 + 2
        dblSCols = .GetDistance(varInsertionPoint, vbCr & "Column spacing: ")
    End With
    
    '' create the object
    ThisDrawing.ModelSpace.AddMInsertBlock varInsertionPoint, strName, _
                 dblX, dblY, dblZ, dR, lngNRows, lngNCols, dblSRows, dblSCols
End Sub

Public Sub TestEditobjMInsertBlock()
Dim objMInsert As AcadMInsertBlock
Dim varPick As Variant
Dim lngNRows As Long
Dim lngNCols As Long
Dim dblSRows As Double
Dim dblSCols As Double
        
On Error Resume Next
    '' get an entity and input from user
    With ThisDrawing.Utility
        .GetEntity objMInsert, varPick, vbCr & "Pick an MInsert: "
        If objMInsert Is Nothing Then
            MsgBox "You did not choose an MInsertBlock object"
            Exit Sub
        End If
        .InitializeUserInput 1 + 2 + 4
        lngNRows = .GetInteger(vbCr & "Number of rows: ")
        .InitializeUserInput 1 + 2 + 4
        lngNCols = .GetInteger(vbCr & "Number of columns: ")
        .InitializeUserInput 1 + 2
        dblSRows = .GetDistance(varPick, vbCr & "Row spacing: ")
        .InitializeUserInput 1 + 2
        dblSCols = .GetDistance(varPick, vbCr & "Column spacing: ")
    End With
    
    '' update the objMInsert
    With objMInsert
        .Rows = lngNRows
        .Columns = lngNCols
        .RowSpacing = dblSRows
        .ColumnSpacing = dblSCols
        .Update
    End With
End Sub

Public Sub GetDynamicBlockProps()
Dim oBlockRef As IAcadBlockReference2
Dim Point As Variant
Dim oEntity As AcadEntity
Dim Props As Variant
Dim Index As Long

  On Error Resume Next
  
  ThisDrawing.Utility.GetEntity oEntity, Point, "Select block ..."

  If oEntity.ObjectName = "AcDbBlockReference" Then
    Set oBlockRef = oEntity
    
    Else
      'no block reference selected
      Exit Sub
  End If
  
  With oBlockRef
    If .IsDynamicBlock = True Then
      Props = .GetDynamicBlockProperties
          
      For Index = LBound(Props) To UBound(Props)
        Dim oProp As AcadDynamicBlockReferenceProperty
        
        Set oProp = Props(Index)
        
        'is the value an array for an insertion point
        If IsArray(oProp.Value) Then
          Dim SubIndex As Long
          
          For SubIndex = LBound(oProp.Value) To UBound(oProp.Value)
            Debug.Print oProp.PropertyName & ", " & oProp.Value(SubIndex)
          Next SubIndex
          
          Else
            Debug.Print oProp.PropertyName & ", " & oProp.Value
        End If
      Next Index
    End If
  End With
End Sub

Public Sub TestAttachExternalReference()
Dim strPath As String
Dim strName As String
Dim varInsertionPoint As Variant
Dim dblX As Double
Dim dblY As Double
Dim dblZ As Double
Dim dblRotation As Double
Dim strInput As String
Dim blnOver As Boolean
          
    '' get input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        strPath = .GetString(True, vbCr & "External file name: ")
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "Block name to create: ")
        .InitializeUserInput 1
        varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
        .InitializeUserInput 1 + 2
        dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
        .InitializeUserInput 1 + 2
        dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
        .InitializeUserInput 1 + 2
        dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
        .InitializeUserInput 1
        dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
        .InitializeUserInput 1, "Attach Overlay"
        strInput = .GetKeyword(vbCr & "Type [Attach/Overlay]: ")
        blnOver = IIf("Overlay" = strInput, True, False)
    End With
    
    '' create the object
    ThisDrawing.ModelSpace.AttachExternalReference strPath, strName, _  
                   varInsertionPoint, dblX, dblY, dblZ, dblRotation, blnOver
End Sub

Public Sub TestExternalReference()
Dim strName As String
Dim strOpt As String
Dim objBlock As AcadBlock
          
On Error Resume Next    '' get input from user
    With ThisDrawing.Utility
        '' get the block name
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "External reference name: ")
        If Err Then Exit Sub
        
        '' get the block definition
        Set objBlock = ThisDrawing.Blocks.Item(strName)
        
        '' exit if not found
        If Err Then
            MsgBox "Unable to get block " & strName
            Exit Sub
        End If
        
        '' exit if not an xref
        If Not objBlock.IsXRef Then
            MsgBox "That is not an external reference"
            Exit Sub
        End If
        
        '' get the operation
        .InitializeUserInput 1, "Detach Reload Unload"
        strOpt = .GetKeyword(vbCr & "Option [Detach/Reload/Unload]: ")
        If Err Then Exit Sub
        
        '' perform operation requested
        If strOpt = "Detach" Then
            objBlock.Detach
        ElseIf strOpt = "Reload" Then
            objBlock.Reload
        Else
            objBlock.Unload
        End If
    End With
End Sub

Public Sub TestBind()
Dim strName As String
Dim strOpt As String
Dim objBlock As AcadBlock

On Error Resume Next
          
    '' get input from user
    With ThisDrawing.Utility
        
        '' get the block name
        .InitializeUserInput 1
        strName = .GetString(True, vbCr & "External reference name: ")
        If Err Then Exit Sub
        
        '' get the block definition
        Set objBlock = ThisDrawing.Blocks.Item(strName)
        
        '' exit if not found
        If Err Then
            MsgBox "Unable to get block " & strName
            Exit Sub
        End If
        
        '' exit if not an xref
        If Not objBlock.IsXRef Then
            MsgBox "That is not an external reference"
            Exit Sub
        End If
        
        '' get the option
        .InitializeUserInput 1, "Prefix Merge"
        strOpt = .GetKeyword(vbCr & "Dependent entries [Prefix/Merge]: ")
        If Err Then Exit Sub
            
        '' perform the bind, using option entered
        objBlock.Bind ("Merge" = strOpt)
    End With
End Sub

Public Sub TestAddAttribute()
    Dim dblOrigin(2) As Double
    Dim dblEnt(2) As Double
    Dim dblHeight As Double
    Dim lngMode As Long
    Dim strTag As String
    Dim strPrompt As String
    Dim strValue As String
    Dim objBlock As AcadBlock
    Dim objEnt As AcadEntity
      
    '' create the block
    dblOrigin(0) = 0:  dblOrigin(1) = 0:  dblOrigin(2) = 0
    Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, "Affirmations")
    
    '' delete existing entities (in case we've run before)
    For Each objEnt In objBlock
        objEnt.Delete
    Next
    
    '' create an ellipse in the block
    dblEnt(0) = 4:  dblEnt(1) = 0:  dblEnt(2) = 0
    objBlock.AddEllipse dblOrigin, dblEnt, 0.5
    
    '' set the height for all attributes
    dblHeight = 0.25
    dblEnt(0) = -1.5:  dblEnt(1) = 0:  dblEnt(2) = 0
    
    '' create a regular attribute
    lngMode = acAttributeModeNormal
    strTag = "Regular"
    strPrompt = "Enter a value"
    strValue = "I'm regular"
    dblEnt(1) = 1
    objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
                          strValue
      
    '' create an invisible attribute
    lngMode = acAttributeModeInvisible
    strTag = "Invisible"
    strPrompt = "Enter a hidden value"
    strValue = "I'm invisible"
    dblEnt(1) = 0.5
    objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
                          strValue
    
    '' create a constant attribute
    lngMode = acAttributeModeConstant
    strTag = "Constant"
    strPrompt = "Don't bother"
    strValue = "I'm set"
    dblEnt(1) = 0
    objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
                          strValue
    
    '' create a verify attribute
    lngMode = acAttributeModeVerify
    strTag = "Verify"
    strPrompt = "Enter an important value"
    strValue = "I'm important"
    dblEnt(1) = -0.5
    objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
                          strValue
    
    '' create a preset attribute
    lngMode = acAttributeModePreset
    strTag = "Preset"
    strPrompt = "No question"
    strValue = "I've got values"
    dblEnt(1) = -1
    objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
                          strValue

    '' now insert block interactively using sendcommand
    ThisDrawing.SendCommand "._-insert" & vbCr & "Affirmations" & vbCr
End Sub

Public Sub TestGetAttributes()
Dim varPick As Variant
Dim objEnt As AcadEntity
Dim objBRef As AcadBlockReference
Dim varAttribs As Variant
Dim strAttribs As String
Dim intI As Integer

On Error Resume Next
    With ThisDrawing.Utility
        
        '' get an entity from user
        .GetEntity objEnt, varPick, vbCr & "Pick a block with attributes: "
        If Err Then Exit Sub
        
        '' cast it to a blockref
        Set objBRef = objEnt
        
        '' exit if not a block
        If objBRef Is Nothing Then
            .Prompt vbCr & "That wasn't a block."
            Exit Sub
        End If
        
        '' exit if it has no attributes
        If Not objBRef.HasAttributes Then
            .Prompt vbCr & "That block doesn't have attributes."
            Exit Sub
        End If
        
        '' get the attributerefs
        varAttribs = objBRef.GetAttributes
        
        '' show some information about each
        strAttribs = "Block Name: " & objBRef.Name & vbCrLf
        For intI = LBound(varAttribs) To UBound(varAttribs)
           strAttribs = strAttribs & " Tag(" & intI & "): " & _
           varAttribs(intI).TagString & vbTab & "  Value(" & intI & "): " & _
           varAttribs(intI).TextString & vbCrLf
        Next

    End With
    MsgBox strAttribs

End Sub

Public Sub TestGetConstantAttributes()
Dim varPick As Variant
Dim objEnt As AcadEntity
Dim objBRef As AcadBlockReference
Dim varAttribs As Variant
Dim strAttribs As String
Dim intI As Integer

    On Error Resume Next
    With ThisDrawing.Utility
        
        '' get an entity from user
        .GetEntity objEnt, varPick, vbCr & _
                   "Pick a block with constant attributes: "
        If Err Then Exit Sub
        
        '' cast it to a blockref
        Set objBRef = objEnt
        
        '' exit if not a block
        If objBRef Is Nothing Then
            .Prompt vbCr & "That wasn't a block."
            Exit Sub
        End If
        
        '' exit if it has no attributes
        If Not objBRef.HasAttributes Then
            .Prompt vbCr & "That block doesn't have attributes."
            Exit Sub
        End If
        
        '' get the constant attributes
        varAttribs = objBRef.GetConstantAttributes
        
        '' show some information about each
        strAttribs = "Block Name: " & objBRef.Name & vbCrLf
        For intI = LBound(varAttribs) To UBound(varAttribs)
            strAttribs = strAttribs & " Tag(" & intI & "): " & _
            varAttribs(intI).TagString & vbTab & "Value(" & intI & "): " & _
            varAttribs(intI).TextString
        Next
    
    End With
    MsgBox strAttribs
End Sub

Function GetAttributes(objBlock As AcadBlock) As Collection
Dim objEnt As AcadEntity
Dim objAttribute As AcadAttribute
Dim coll As New Collection
                
    '' iterate the block
    For Each objEnt In objBlock
    
        '' if it's an attribute
        If objEnt.ObjectName = "AcDbAttributeDefinition" Then
        
            '' cast to an attribute
            Set objAttribute = objEnt
        
            '' add attribute to the collection
            coll.Add objAttribute, objAttribute.TagString
        End If
    Next
    
    '' return collection
    Set GetAttributes = coll
End Function

Public Sub DemoGetAttributes()
Dim objAttribs As Collection
Dim objAttrib As AcadAttribute
Dim objBlock As AcadBlock
Dim strAttribs As String

    '' get the block
    Set objBlock = ThisDrawing.Blocks.Item("Affirmations")
    
    '' get the attributes
    Set objAttribs = GetAttributes(objBlock)
    
    '' show some information about each
    
    For Each objAttrib In objAttribs
        strAttribs = objAttrib.TagString & vbCrLf
        strAttribs = strAttribs & "Tag: " & objAttrib.TagString & vbCrLf & _
"Prompt: " & objAttrib.PromptString & vbCrLf & " Value: " & _
 objAttrib.TextString & vbCrLf & "  Mode: " & _
 objAttrib.Mode
        MsgBox strAttribs
    Next
    
    '' find specific attribute by TagString
    Set objAttrib = objAttribs.Item("PRESET")
    
    '' prove that we have the right one
    strAttribs = "Tag: " & objAttrib.TagString & vbCrLf & "Prompt: " & _
objAttrib.PromptString & vbCrLf & "Value: " & objAttrib.TextString & _
vbCrLf & "Mode: " & objAttrib.Mode
    MsgBox strAttribs
End Sub

Public Sub TestInsertAndSetAttributes()
Dim objBRef As AcadBlockReference
Dim varAttribRef As Variant
Dim varInsertionPoint As Variant
Dim dblX As Double
Dim dblY As Double
Dim dblZ As Double
Dim dblRotation As Double
          
    '' get block input from user
    With ThisDrawing.Utility
        .InitializeUserInput 1
        varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
        .InitializeUserInput 1 + 2
        dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
        .InitializeUserInput 1 + 2
        dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
        .InitializeUserInput 1 + 2
        dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
        .InitializeUserInput 1
        dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
    End With
    
    '' insert the block
    Set objBRef = ThisDrawing.ModelSpace.InsertBlock(varInsertionPoint, _
                              "Affirmations", dblX, dblY, dblZ, dblRotation)
        
    '' interate the attributerefs
    For Each varAttribRef In objBRef.GetAttributes
        
        '' change specific values based on Tag
        Select Case varAttribRef.TagString
        Case "Regular":
            varAttribRef.TextString = "I have new values"
        Case "Invisible":
            varAttribRef.TextString = "I'm still invisible"
        Case "Verify":
            varAttribRef.TextString = "No verification needed"
        Case "Preset":
            varAttribRef.TextString = "I can be changed"
        End Select
    Next
End Sub

; CHAPTER 14 Views and Viewports. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 321

; CHAPTER 15 Layout and Plot Configurations . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 337

; CHAPTER 16 Controlling Menus and Toolbars . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 355

; CHAPTER 17 Drawing Security . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 383

; CHAPTER 18 Using the Windows API. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 391

; CHAPTER 19 Connecting to External Applications . . . . . . . . . . . . . . . . . . . . . . . . . . 403

; CHAPTER 20 Creating Tables . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 415

; CHAPTER 21 The SummaryInfo Object. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 427

; CHAPTER 22 An Illustrative VBA Application . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 437

; APPENDIX A AutoCAD Object Summary . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 447

; APPENDIX B AutoCAD Constants Reference . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 631


; APPENDIX C System Variables . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 671


