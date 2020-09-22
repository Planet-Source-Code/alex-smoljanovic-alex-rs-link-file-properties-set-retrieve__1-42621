VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link Properties"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   5040
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   660
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   5040
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   4455
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   5055
      Begin VB.TextBox txtIconInd 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   4020
         Width           =   495
      End
      Begin VB.TextBox txtIconPath 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   3660
         Width           =   3855
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbShowCmd 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   1380
         List            =   "Form1.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2280
         Width           =   3555
      End
      Begin VB.TextBox txtWorkingDirectory 
         Height          =   315
         Left            =   1380
         TabIndex        =   15
         Top             =   3120
         Width           =   3555
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   900
         Width           =   3855
      End
      Begin VB.TextBox txtHotkey 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2700
         Width           =   3555
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Top             =   1740
         Width           =   3855
      End
      Begin VB.TextBox txtArg 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label9 
         Caption         =   "Icon Index:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Icon Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4920
         Y1              =   3540
         Y2              =   3540
      End
      Begin VB.Label lblLinkName 
         Height          =   195
         Left            =   1080
         TabIndex        =   18
         Top             =   360
         Width           =   3840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Directory:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   675
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4920
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Show Command:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Path:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hotkey:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2820
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Argument:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1380
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   315
      Left            =   4740
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.TextBox txtLink 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Link:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'This application was explicitly developed for
'PSC(Planet Source Code) Users as an Open Source Project.
'This code is the property of it's author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************


Dim iLink As ShellLinkObject, IconPath$, IconInd&
'General Declarations

Private Sub GetLinkInfo(Optional Path$ = "")
On Error GoTo errh 'on the event of an error jump to label errh
Dim iShell As Shell, iFolder As Folder, ifItem As FolderItem
'Dimensionalize iShell as object Shell, iFolder as object Folder, ifItem as object FolderItem
 cd.DialogTitle = "Open Link" 'Initialize the CommonDialog control's DialogTitle property
  cd.Filter = "Link File|*.lnk|All Files|*.*"
  'Update the controls file type extension; syntax: Description|Pattern|Description|Pattern
   cd.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
   'See Microsofts CommonDialog control documentaion for description of common dialog flags (http://msdn.microsoft.com)
    If Path$ = "" Then cd.ShowOpen Else cd.FileName = Path$: If Dir(Path$) = "" Then Exit Sub
    'if the optional argument evaluates to "" then call common dialog control's ShowOpen method
    'to show the common dialog open modal dialog, otherwise initialize object cd's FileName property with the value of the path argument
     Set iShell = New Shell 'initialize iShell with a new instance of the Shell class
      Set iFolder = iShell.NameSpace(Left$(cd.FileName, InStrRev(cd.FileName, "\", , 1) - 1))
      'initialize iFolder with the Folder class return of Shell's NameSpace method, which will return a Folder object who's path evaluates to CD.FileName's member value
       Set ifItem = iFolder.ParseName(cd.FileTitle)
       'initialize ifItem(Folder Item) with the FolderItem class returned by iFolder's ParseName method, this function will return the specified folder item(CD's FileTitle property which returns the name of the file specified by CD's FileName property)
        If ifItem.IsLink = True Then
        'Verify that the folder item ifItem is a link object
         Set iLink = ifItem.GetLink
         'initialize iLink(ShellLinkObject class) with the ShellLinkObject returned by ifItem's GetLink method
        End If
         txtLink = cd.FileName: txtLink.SelStart = Len(txtLink) - 1
         'set the text box txtLink's text property to the CD's FileName property, move the caret position to the end of the object's text property
          txtArg = iLink.Arguments 'update the text boxes text property with the link's argument attribute value
           txtDescription = iLink.Description '...
            txtHotkey = iLink.Hotkey '...
             txtPath = iLink.Path '...
              Select Case iLink.ShowCommand
               Case 1: cmbShowCmd.ListIndex = 0 'Normal Show Command
               Case 7: cmbShowCmd.ListIndex = 1 'Minimzed Show Command
               Case 3: cmbShowCmd.ListIndex = 2 'Maximized Show Command
              End Select
               lblLinkName = cd.FileTitle '...
                txtWorkingDirectory = iLink.WorkingDirectory '...
                 IconPath = String(260, 0)
                 'Allocate memory to the variable IconPath,
                 'the number argument of the function String is significant
                 'since 260 is the maximum length of characters a valid path can consist of.
                  IconInd = iLink.GetIconLocation(IconPath)
                  'ShellLinkObject's GetIconLocation method returns the Icon Index
                  'It will also initialize the variable to which the argument pbs points
                  'to if the Icon Path isn't equal to the target of the icon
                   If InStr(1, IconPath, Chr(0), 1) > 0 Then
                   'determine the position(starting) of a sub-string within a string
                    IconPath = Left(IconPath, InStr(1, IconPath, Chr(0), 1) - 1)
                    'return the string removing the null-terminating characters
                   End If
                    If IconPath = "" Then IconPath = iLink.Path
                    'if the function returned a Null String, then the Icon Path of this Link File evaluates
                    'to the target of the link rather than another module
                     picIcon.Cls 'clear the picturebox control
                      GetIcon cd.FileName, picIcon
                      'see GetLargeIcon for more info...
                       picIcon.Refresh 'refresh the picturebox control since its AutoRedraw(Persistent Bitmap) property evaluates to true
                        txtIconPath = IconPath '...
                         txtIconInd = IconInd '...
                          Set iShell = Nothing: Set iFolder = Nothing: Set ifItem = Nothing
                          'Destroy class instances...
                           Exit Sub 'discontinue execution of this procedure
errh: 'label errh
 If Err.Number = 32755 Then Exit Sub
 'The user cancelled the Common Dialog's dialog, exit this procedure
  MsgBox "An error occured." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Error " & "[" & Err.Number & "]"
  'Inform user of the error
   Set iShell = Nothing: Set iFolder = Nothing: Set ifItem = Nothing
   'destroy shell's class instances
End Sub

Private Sub cmdOpen_Click()
 GetLinkInfo 'See GetLinkInfo for more info...
End Sub

Private Sub cmdSave_Click()
If Dir(txtPath) = "" Then MsgBox "The links path property must point to a valid file.", vbExclamation, "Invalid Path": Exit Sub
'If the text property of the text box txtPath point to a non-existent path then call MsgBox function to show a modal dialog informing the user, then exit this procedure....
 If txtWorkingDirectory <> "" And Dir(txtWorkingDirectory, vbDirectory) = "" Then MsgBox "The links Working Directory must be a valid directory.", vbExclamation, "Invalid Directory": Exit Sub '...
  If Dir(txtIconPath) = "" Then MsgBox "The links Icon Path property must point to a valid file.", vbExclamation, "Invalid Path": Exit Sub '...
   iLink.Arguments = txtArg 'Update the ShellLinkObject iLink's Arguments property
    iLink.Description = txtDescription '...
     iLink.Path = txtPath '...
      Select Case cmbShowCmd.ListIndex
       Case 0: iLink.ShowCommand = 1 'Normal Show Command
       Case 1: iLink.ShowCommand = 7 'Minimized Show Command
       Case 2: iLink.ShowCommand = 3 'Maximized Show Command
      End Select
       If txtWorkingDirectory <> "" Then iLink.WorkingDirectory = IIf(Right(txtWorkingDirectory, 1) <> "\", txtWorkingDirectory & "\", txtWorkingDirectory)
       'if txtWorkingDirectory's text property is unequal to a null string(""),
       'initialize iLink's WorkingDirectory property with the return of the IIf operator,
       'which will conditionally append the character "\" to the string if the last character isn't "\"
        iLink.SetIconLocation txtIconPath, txtIconInd
        'call iLink's SetIconLocation method to update the link object's Icon Path and Icon Index
         iLink.Save 'Save the updated link object's information
          GetLinkInfo txtLink 'Refresh
          'GetLinkInfo's optional argument path is passes so that this procedure doesn't show the open file modal dialog called by CD(CommonDialog control)
End Sub

Private Sub Command1_Click()
 Unload Me 'Destroy this window
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Not (iLink Is Nothing) Then Set iLink = Nothing
 'Determine if object iLink has been initialized,
 'if it has been initialized with an instance of a class, then destroy its instance
End Sub
