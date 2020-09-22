VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl EDriveListBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ScaleHeight     =   600
   ScaleWidth      =   3825
   ToolboxBitmap   =   "EDriveListBox.ctx":0000
   Begin MSComctlLib.ImageList imgList 
      Left            =   3150
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":0532
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":0A74
            Key             =   "Floppy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":0EC6
            Key             =   "Cd"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":1318
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":176A
            Key             =   "NetWork"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":1BBC
            Key             =   "Removable"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":200E
            Key             =   "Unrecognized"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":2460
            Key             =   "Ram"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EDriveListBox.ctx":28B2
            Key             =   "DriveNotFound"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo imgCboDrive 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Enached Drive List Box"
      Top             =   0
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "List of Drives"
   End
End
Attribute VB_Name = "EDriveListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'api for disk
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Remember: the GetVolumeInformations Api require:
'Ucase letter and ":" and "\"!
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
'events-----------------------
Public Event Change()
Public Event Scroll()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)


'------Ide infos:
'are you running a compiled Ocx?
'are you in design time of a standard exe?
'are you running a standard exe (in ide or compiled) ?

Enum enumEDInInde
   EDIDevelopeControl = 0 'you're developing Ocx.Exe is in ide, running or not
   EDIDevelopeContainer = 1 'the ocx is compiled, and the standard exe is in
                             'design time (not running)
   EDICompiled = 3 'this means ocx is compiled, not standard exe
                    'your standard exe is running. It might be running from ide or
                    'as compiled exe
End Enum

Enum sizeUnit
   suByte = 0
   suKb = 1
   suMb = 2
   suGb = 3
End Enum

'Properties
Private mIsInIde As enumEDInInde 'readonly

Private mDriveLetter As String   'current selected DriveLetter
Private mDrive As String 'current drive text, similar to drive1 object
'should hide some type of disk? Here what you can hide:
Private mHideCd As Boolean
Private mHideRemovable As Boolean
Private mHideFixed As Boolean
Private mHideNetWork As Boolean
Private mHideRam As Boolean
Private mHideUnknown As Boolean

Private Sub UserControl_Initialize()
    'infos about ide
   mIsInIde = getIde
   Set imgCboDrive.ImageList = imgList
   getAllDRives
End Sub

Private Sub UserControl_Paint()

   imgCboDrive.Move 0, 0
   UserControl.Size imgCboDrive.Width, imgCboDrive.Height

End Sub

Private Sub UserControl_Resize()
If Width < 150 Then
   Width = 150
   Exit Sub
End If
If Height <> imgCboDrive.Height Then
   Height = imgCboDrive.Height
   Exit Sub
End If
'size control
imgCboDrive.Move 0, 0, Width
End Sub

Public Property Get IsInIde() As enumEDInInde
Attribute IsInIde.VB_Description = "Tell if control is in developing or compiled, and if container is at design time or running time. ReadOnly"
    IsInIde = mIsInIde
End Property

Public Property Get Drive() As String
   'show selected Drive as text displayed
   Drive = imgCboDrive.Text
End Property

Public Property Let Drive(ByVal sDrive As String)
   'Set driveletter (next property)
   DriveLetter = sDrive
End Property


Public Property Get DriveLetter() As String
   'Holds selected DriveLetter
   DriveLetter = mDriveLetter
End Property


Public Property Let DriveLetter(ByVal sDriveLetter As String)
Attribute DriveLetter.VB_Description = "Set/return selected Drive letter, without dots or slash "
Attribute DriveLetter.VB_ProcData.VB_Invoke_PropertyPut = ";Data"
Attribute DriveLetter.VB_MemberFlags = "200"
   'Holds selected DriveLetter
   
   sDriveLetter = LCase(Trim(sDriveLetter))
   If Len(sDriveLetter) > 1 Then
        sDriveLetter = Left$(sDriveLetter, 1)
   End If
   If SelectADrive(sDriveLetter) Then
      If mDriveLetter <> sDriveLetter Then
        mDriveLetter = sDriveLetter
        RaiseEvent Change
        PropertyChanged DriveLetter
     End If
        
        
   Else
      If mDriveLetter <> "" Then
         mDriveLetter = ""
         RaiseEvent Change
         PropertyChanged DriveLetter
      End If
        'DriveLetter not listed in listImage
        'may be you choosed to hide it?
        'this control should never raise errors:
        'the DriveListBox does not(!) -in case you need,
        'uncomment following lines
        '---------------------------------------------
'        If mIsInIde = EDICompiled Or mIsInIde = EDIDevelopeContainer Then
'            'Err.Raise works in compiled ocx only,
'            'in develope mode of ocx, you will not be able to
'            'handle this error!
'            Err.Raise 68, TypeName(Me), "Device " & sDriveLetter & " unavailable"
'        ElseIf mIsInIde = EDIDevelopeControl Then
'            'msgbox, and not raiseError to avoid stopping
'            MsgBox "Err 68: the DriveLetter " & sDriveLetter & " is unavailable. In compiled Ocx this would raise error 68.", vbOKOnly + vbCritical, TypeName(Me)
'        End If
         '---------------------------------------------
   End If
   
End Property

Public Property Get HideCd() As Boolean
Attribute HideCd.VB_Description = "Boolean. (True)Hide/(False)Show Cd drives from list of drives"
   HideCd = mHideCd
End Property

Public Property Let HideCd(ByVal HideCdRom As Boolean)
   If HideCdRom <> mHideCd Then
      mHideCd = HideCdRom
      getAllDRives
      PropertyChanged HideCd
   End If
End Property

Public Property Get HideRemovable() As Boolean
Attribute HideRemovable.VB_Description = "Boolean. (True)Hide/(False)Show Removable drives from list of drives. Note: Floppy drives are detected as removable ones!"
   HideRemovable = mHideRemovable
End Property

Public Property Let HideRemovable(ByVal HideRemovableDisk As Boolean)
   If HideRemovableDisk <> mHideRemovable Then
      mHideRemovable = HideRemovableDisk
      getAllDRives
      PropertyChanged HideRemovable
   End If
End Property

Public Property Get HideFixed() As Boolean
Attribute HideFixed.VB_Description = "Boolean. (True)Hide/(False)Show Fixed drives from list of drives"
   HideFixed = mHideFixed
End Property

Public Property Let HideFixed(ByVal HideFixedDisk As Boolean)
   If HideFixedDisk <> mHideFixed Then
      mHideFixed = HideFixedDisk
      getAllDRives
      PropertyChanged HideFixed
   End If
End Property

Public Property Get HideNetWork() As Boolean
Attribute HideNetWork.VB_Description = "Boolean. (True)Hide/(False)Show Network drives from list of drives"
   HideNetWork = mHideNetWork
End Property

Public Property Let HideNetWork(ByVal HideNetWorkDisk As Boolean)
   If mHideNetWork <> HideNetWorkDisk Then
      mHideNetWork = HideNetWorkDisk
      getAllDRives
     PropertyChanged HideNetWork
   End If
End Property

Public Property Get HideRam() As Boolean
Attribute HideRam.VB_Description = "Boolean. (True)Hide/(False)Show Ram drives from list of drives"
   HideRam = mHideRam
End Property

Public Property Let HideRam(ByVal HideRamDisk As Boolean)
   If HideRamDisk <> mHideRam Then
      mHideRam = HideRamDisk
      getAllDRives
      PropertyChanged HideRam
   End If
End Property

Public Property Get HideUnknown() As Boolean
   HideUnknown = mHideRemovable
End Property

Public Property Let HideUnknown(ByVal HideUnknownDisk As Boolean)
   If HideUnknownDisk <> mHideUnknown Then
      mHideUnknown = HideUnknownDisk
      getAllDRives
      PropertyChanged HideUnknown
   End If
End Property


Private Sub imgCboDrive_Click()
  DriveLetter = Left$(imgCboDrive.Text, 1)
  
End Sub

Private Sub imgCboDrive_Dropdown()
   RaiseEvent Scroll
End Sub

Private Sub imgCboDrive_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub imgCboDrive_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub imgCboDrive_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub



Private Sub getAllDRives()
   Dim LDs As Long
   Dim Counter As Long
   Dim sLetter As String
   Dim ImgKey As String
  
   Dim VName As String
   Dim SelThis As String
    
   imgCboDrive.ComboItems.Clear
   LDs = GetLogicalDrives
   
   For Counter = 0 To 25
   
        If (LDs And 2 ^ Counter) <> 0 Then
            
            sLetter = Chr$(Asc("a") + Counter)
            
            'the letter is available
            VName = ""
            Select Case GetDriveType(UCase(sLetter) & ":\")
               Case 2
                  If mHideRemovable Then
                     ImgKey = ""
                  Else
                     If LCase(sLetter) <> "a" Then
                        ImgKey = "Removable"
                     Else
                        ImgKey = "Floppy"
                     End If
                  End If
               Case 3
                  If mHideFixed Then
                     ImgKey = ""
                  Else
                     ImgKey = "Fixed"
                     VName = getExtraInfos(UCase(sLetter) & ":\")
                  End If
                Case Is = 4
                  If mHideNetWork Then
                     ImgKey = ""
                  Else
                    ImgKey = "NetWork"
                  End If
                Case Is = 5
                  If mHideCd Then
                     ImgKey = ""
                  Else
                    ImgKey = "Cd"
                  End If
                Case Is = 6
                  If mHideRam Then
                     ImgKey = ""
                  Else
                    ImgKey = "Ram"
                  End If
                Case Else
                  If mHideUnknown Then
                     ImgKey = ""
                  Else
                     ImgKey = "Unrecognized"
                  End If
            End Select
            If Len(ImgKey) > 0 Then
               imgCboDrive.ComboItems.Add , sLetter, sLetter & ":" & VName, ImgKey
            End If
           
        End If
    Next Counter
   Dim tmpDLetter As String
   If imgCboDrive.ComboItems.Count > 0 Then
      'use app.path to determine which is current disk
      'but be sure you have it, else use fiorst you have if any
       
       tmpDLetter = mDriveLetter
       If tmpDLetter = "" Then
         tmpDLetter = LCase(Left$(App.Path, 1))
       End If
       'select the current DriveLetter or the setted one if you can
       If Not SelectADrive(tmpDLetter) Then
            tmpDLetter = ""
       End If
       
   Else
      tmpDLetter = ""
      
   End If
   If tmpDLetter <> mDriveLetter Then
      DriveLetter = tmpDLetter
   End If
End Sub

Private Sub UserControl_Terminate()
   'TMouse.Enabled = False
End Sub

Private Function SelectADrive(ByVal sDriveLetter As String) As Boolean
    Dim Counter As Integer
    
    imgCboDrive.Text = ""
    For Counter = 1 To imgCboDrive.ComboItems.Count
         If imgCboDrive.ComboItems.Item(Counter).Key = sDriveLetter Then
            imgCboDrive.ComboItems(Counter).Selected = True
            SelectADrive = True
            Exit For
         End If
    Next
    
    
    If imgCboDrive.Text = "" Then
        imgCboDrive.Text = "-Choose a DriveLetter-"
    End If
End Function

Private Function getIde() As enumEDInInde
    On Error Resume Next
     Debug.Print 1 / 0
     If Err.Number <> 0 Then
        Err.Clear
        getIde = EDIDevelopeControl 'you are developing the Ocx,
                                    'do not know if form is running in ide ro not
        
     Else
        If Ambient.UserMode Then
            getIde = EDICompiled 'you are using a compiled ocx,
                                 'and you are at least running the standard project
        Else
            getIde = EDIDevelopeContainer 'you are in design time of your standard exe
        End If
     End If
End Function

Private Sub UserControl_InitProperties()
    'default values of properties
    DriveLetter = mDriveLetter
    HideCd = False
    HideRemovable = False
    HideFixed = False
    HideNetWork = False
    HideRam = False
    HideUnknown = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save properties as ocx will restart many time
    'while stepping throug designing standard exe and
    'launching it
    PropBag.WriteProperty "DriveLetter", DriveLetter, mDriveLetter
    PropBag.WriteProperty "HideCd", HideCd, False
    PropBag.WriteProperty "HideRemovable", HideRemovable, False
    PropBag.WriteProperty "HideFixed", HideFixed, False
    PropBag.WriteProperty "HideNetWork", HideNetWork, False
    PropBag.WriteProperty "HideRam", HideRam, False
    PropBag.WriteProperty "HideUnknown", HideUnknown, True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'reading properties may result in an error
    'as the frm file might be modified manually
    'by external editor with invalid values for
    'these properties...
    On Error Resume Next
    DriveLetter = PropBag.ReadProperty("DriveLetter", mDriveLetter)
    HideCd = PropBag.ReadProperty("HideCd", False)
    HideRemovable = PropBag.ReadProperty("HideRemovable", False)
    HideFixed = PropBag.ReadProperty("HideFixed", False)
    HideNetWork = PropBag.ReadProperty("HideNetWork", False)
    HideRam = PropBag.ReadProperty("HideRam", False)
    HideUnknown = PropBag.ReadProperty("HideUnknown", True)
End Sub

Public Function driveFreeSpace(sDriveLetter As String, Unit As sizeUnit) As Currency
   Dim freeToCaller As Currency
   Dim TotBytes As Currency
   Dim FreeBytes As Currency
   
   Select Case Len(sDriveLetter)
      Case 1
         sDriveLetter = sDriveLetter & ":"
      Case 2
         'Ok
      Case 3
         sDriveLetter = Left$(sDriveLetter, 2)
       Case Else
         Exit Function
   End Select
    'Retrieve information about the disk
    GetDiskFreeSpaceEx UCase(sDriveLetter), freeToCaller, TotBytes, FreeBytes
        
    'Freeb = FreeC& * Sectors& * Bytes&
    Dim rMult As Double
    Select Case Unit
      Case suByte
         rMult = 10000 'the currency has 4 decimal places
      Case suKb
         rMult = 10000 / 1024
      Case suMb
         rMult = CDbl(10000) / CDbl(1048576)
      Case suGb
       rMult = CDbl(10000) / CDbl(1073741824)
    End Select
    
    
      driveFreeSpace = FreeBytes * rMult
   
End Function

Public Function ReturnFirstReadyDiskLetter() As String
   'now...
   Dim Counter As Integer
   Dim tmpVal As Double
   On Error Resume Next
   For Counter = 1 To imgCboDrive.ComboItems.Count
      If isDiskReady(UCase(imgCboDrive.ComboItems(Counter).Key) & ":\") Then
         ReturnFirstReadyDiskLetter = imgCboDrive.ComboItems(Counter).Key
         Exit For
      End If
   Next
End Function

Private Function isDiskReady(ByVal sDriveLetter As String) As Boolean
  
  
   Dim DrvVolumeName As String
   Dim UnusedVal1 As Long
   Dim UnusedVal2 As Long
   Dim UnusedVal3 As Long
   Dim UnusedStr As String
  
   DrvVolumeName = Space$(14)
   UnusedStr = Space$(32)

  
   If GetVolumeInformation(sDriveLetter, _
                            DrvVolumeName, _
                            Len(DrvVolumeName), _
                            UnusedVal1, UnusedVal2, _
                            UnusedVal3, _
                            UnusedStr, Len(UnusedStr)) > 0 Then

      isDiskReady = True

   End If

End Function

Private Function getExtraInfos(ByVal sLetter As String) As String
   'get extra infos
   'Create buffers
   Dim tmpName As String
   Dim tmpFSName As String
   Dim tmpSerial As Long
   tmpName = String$(255, Chr$(0))
   tmpFSName = String$(255, Chr$(0))
   'Get the volume information
   GetVolumeInformation sLetter, tmpName, 255, tmpSerial, 0, 0, tmpFSName, 255
   'Strip the extra chr$(0)'s
   tmpName = Left$(tmpName, InStr(1, tmpName, Chr$(0)) - 1)
   tmpFSName = Left$(tmpFSName, InStr(1, tmpFSName, Chr$(0)) - 1)
   If Len(tmpName) > 0 Then
      tmpName = " [" & tmpName & "]"
   End If
   getExtraInfos = tmpName
End Function


