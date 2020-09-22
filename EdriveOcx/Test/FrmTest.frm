VERSION 5.00
Object = "{DDF24EB4-85E1-478B-A375-C3A9D4E65E54}#6.0#0"; "EnanchedDriveLIstBox.ocx"
Begin VB.Form FrmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Enanched DriveList Box"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide Unknown"
      Height          =   240
      Index           =   5
      Left            =   3225
      TabIndex        =   13
      Top             =   2475
      Value           =   1  'Checked
      Width           =   3165
   End
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide Ram"
      Height          =   240
      Index           =   4
      Left            =   3225
      TabIndex        =   11
      Top             =   2175
      Width           =   3165
   End
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide Network"
      Height          =   240
      Index           =   3
      Left            =   3225
      TabIndex        =   10
      Top             =   1875
      Width           =   3165
   End
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide Fixed Disks"
      Height          =   240
      Index           =   2
      Left            =   0
      TabIndex        =   9
      Top             =   2550
      Width           =   3165
   End
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide CdRoms"
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   2250
      Width           =   3165
   End
   Begin VB.CheckBox ChkHide 
      Caption         =   "Hide Removables or Floppy"
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1950
      Width           =   3165
   End
   Begin VB.Frame FraMore 
      Caption         =   "Extra Infos"
      Height          =   1290
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   6315
      Begin VB.ComboBox CboSize 
         Height          =   315
         Left            =   3450
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   675
         Width           =   765
      End
      Begin VB.Label LblDriveLetter 
         AutoSize        =   -1  'True
         Caption         =   "Drive Letter"
         Height          =   195
         Left            =   3450
         TabIndex        =   6
         Top             =   300
         Width           =   825
      End
      Begin VB.Label LblDrive 
         AutoSize        =   -1  'True
         Caption         =   "Drive"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   300
         Width           =   375
      End
      Begin VB.Label LblFreeSize 
         AutoSize        =   -1  'True
         Caption         =   "Free space"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   750
         Width           =   795
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3075
      TabIndex        =   2
      Top             =   0
      Width           =   3390
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   3015
   End
   Begin EnanchedDriveLIstBox.EDriveListBox Edsk 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      HideUnknown     =   0   'False
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboSize_Click()
     LblFreeSize.Caption = "Free space : " & Edsk.driveFreeSpace(Edsk.DriveLetter, CboSize.ListIndex) & Space(1) & CboSize.Text
End Sub

Private Sub ChkHide_Click(Index As Integer)
   With Edsk
      .HideRemovable = CBool(ChkHide(0).Value = vbChecked)
      .HideCd = CBool(ChkHide(1).Value = vbChecked)
      .HideFixed = CBool(ChkHide(2).Value = vbChecked)
      .HideNetWork = CBool(ChkHide(3).Value = vbChecked)
      .HideRam = CBool(ChkHide(4).Value = vbChecked)
      .HideUnknown = CBool(ChkHide(5).Value = vbChecked)
   End With
End Sub



Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Edsk_Change()
   With Edsk
      LblDrive.Caption = "Drive Property is = " & .Drive
      LblDriveLetter.Caption = "DriveLetter Property  is = " & .DriveLetter
      LblFreeSize.Caption = "Free space : " & .driveFreeSpace(.DriveLetter, CboSize.ListIndex) & Space(1) & CboSize.Text
      'sync drive and folder
      'You need errHandler, as it could return invalid value (="")
      'when user choose to make disappear some types of disk
      On Error GoTo errHandler
      Dir1.Path = .DriveLetter & ":"
   
   End With
   Exit Sub
errHandler:
   Dim tmpDRive As String
   tmpDRive = Edsk.ReturnFirstReadyDiskLetter
   If Len(tmpDRive) > 0 Then
      Edsk.DriveLetter = tmpDRive
   Else
      MsgBox "No ready disks To show!"
   End If
End Sub

Private Sub Form_Load()
   
   CboSize.AddItem "Byte"
   CboSize.AddItem "Kb"
   CboSize.AddItem "Mb"
   CboSize.AddItem "GB"
   CboSize.ListIndex = 2
   'call  it at load first time if you need to retrieve infos,
   'Call Edsk_Change
   'or assign a starting disk
   
   Edsk.DriveLetter = Left$(App.Path, 1)
   
End Sub
