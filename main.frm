VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SpeedConnection"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   409
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws4 
      Left            =   6480
      Top             =   2.45745e5
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws3 
      Left            =   5400
      Top             =   2.45745e5
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws2 
      Left            =   4800
      Top             =   2.45745e5
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   3960
      Top             =   2.45745e5
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fm4 
      Height          =   5535
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   7215
      Begin VB.CommandButton cdsave 
         Caption         =   "Save Picture"
         Height          =   495
         Left            =   1440
         TabIndex        =   16
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cdpclear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   4920
         Width           =   1215
      End
      Begin VB.PictureBox precieve 
         Height          =   4575
         Left            =   120
         ScaleHeight     =   301
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   461
         TabIndex        =   14
         Top             =   240
         Width           =   6975
         Begin VB.CommandButton cdcenter 
            Height          =   195
            Left            =   6735
            TabIndex        =   25
            Top             =   4335
            Width           =   180
         End
         Begin VB.HScrollBar hs 
            Height          =   210
            LargeChange     =   20
            Left            =   0
            Max             =   0
            SmallChange     =   20
            TabIndex        =   24
            Top             =   4305
            Width           =   6735
         End
         Begin VB.VScrollBar vs 
            Height          =   4335
            LargeChange     =   20
            Left            =   6705
            Max             =   0
            SmallChange     =   20
            TabIndex        =   23
            Top             =   0
            Width           =   210
         End
         Begin VB.Image igrecieve 
            Height          =   3855
            Left            =   0
            Top             =   0
            Width           =   6135
         End
      End
   End
   Begin VB.Frame fm3 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7215
      Begin VB.Frame fm 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   6975
         Begin VB.Frame fm 
            Height          =   855
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   1200
            Width           =   5655
            Begin VB.TextBox txstatus4 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   39
               Text            =   "Recieving  (10%)"
               Top             =   525
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txstatus3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   38
               Text            =   "Sending     (10%)"
               Top             =   195
               Visible         =   0   'False
               Width           =   1335
            End
            Begin MSComctlLib.ProgressBar pb4 
               Height          =   255
               Left            =   1680
               TabIndex        =   37
               Top             =   510
               Visible         =   0   'False
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
            Begin MSComctlLib.ProgressBar pb3 
               Height          =   255
               Left            =   1680
               TabIndex        =   36
               Top             =   180
               Visible         =   0   'False
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
         End
         Begin VB.CommandButton cdsendfile 
            Caption         =   "Send"
            Height          =   375
            Left            =   5880
            TabIndex        =   30
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txfilename 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label lbfilesize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ". . . . . . . . . ."
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lb 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File size"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lb 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "File name and path"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.FileListBox fl1 
         Height          =   2625
         Left            =   3600
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.DriveListBox dv1 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6975
      End
      Begin VB.DirListBox dr1 
         Height          =   2565
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame fm2 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7215
      Begin VB.TextBox txsend 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   240
         Width           =   6975
      End
      Begin VB.Frame fm 
         Height          =   735
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   2400
         Width           =   4935
         Begin MSComctlLib.ProgressBar pb2 
            Height          =   255
            Left            =   1380
            TabIndex        =   34
            Top             =   420
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.TextBox txstatus2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "Recieving(10%)"
            Top             =   420
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox txstatus1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Sending   (10%)"
            Top             =   150
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   255
            Left            =   1380
            TabIndex        =   21
            Top             =   120
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.CommandButton cdsendpicture 
         Caption         =   "Send Picture"
         Height          =   615
         Left            =   6000
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cdclear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txrecieve 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3360
         Width           =   6975
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   7080
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   120
         X2              =   7080
         Y1              =   3225
         Y2              =   3225
      End
   End
   Begin VB.Frame fm1 
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7215
      Begin VB.Frame fmstatus 
         Caption         =   "Status"
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   3480
         TabIndex        =   31
         Top             =   360
         Width           =   3615
         Begin VB.TextBox txstatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   1215
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.CommandButton cddisconnect 
         Caption         =   "Disconnect"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton op 
         Caption         =   "Listen"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   5895
      End
      Begin VB.CommandButton cdconnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txip 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote IP"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   750
      End
   End
   Begin MSComctlLib.TabStrip tb1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10821
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Connectiom"
            Key             =   "tab1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chat"
            Key             =   "tab2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files"
            Key             =   "tab3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Picture"
            Key             =   "tab4"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'chat
Dim typesend As String
Dim typerecieve As String
'picture
Dim picdatasend As String
Dim picdatarecieve As String
Dim picrecievesize
Dim totalrecieve As Long
'files
Dim filedatasend As String
Dim filedatarecieve As String
Dim filerecievesize
Dim totalrecieve1 As Long
Dim fileextention As String
Dim sendfilename As String
'open file
Dim filename As String
Dim filesize As Long
Dim onlyfilename As String
'connection control
Dim commanddatasend
Dim commanddatarecieve

Dim ctr
Dim cmd
Dim v, n, m, b, t, s, p, q, c, o, o1
Dim k As String
Dim pb1value
Dim pb3value
Dim connected As Boolean


Public Sub openfile()



Dim filebox As OPENFILENAME
Dim fn As String
Dim retval As Long

filebox.lStructSize = Len(filebox)
filebox.hwndOwner = main.hWnd
filebox.lpstrTitle = "Open File"

filebox.lpstrFilter = "BMP Files" & vbNullChar & "*.BMP" & vbNullChar & "GIF Files" & vbNullChar & "*.GIF" & vbNullChar & "JPG Files" & vbNullChar & "*.JPG" & vbNullChar & vbNullChar
filebox.lpstrFile = Space(255)
filebox.nMaxFile = 255
filebox.lpstrFileTitle = Space(255)
filebox.nMaxFileTitle = 255

filebox.flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY


retval = GetOpenFileName(filebox)
If retval <> 0 Then
  fn = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
  filename = fn
  filesize = FileLen(fn)
End If
For t = filesize To 1 Step -1
   If Mid(filename, t, 1) = "\" Then
     Exit For
   End If
Next t

onlyfilename = Mid(filename, t + 1, t)

Exit Sub

err:
  Exit Sub
  
End Sub

Private Sub cdcenter_Click()

igrecieve.Top = 0
igrecieve.Left = 0

hs.Value = 0
vs.Value = 0

End Sub

Private Sub cdclear_Click()

txrecieve = ""

End Sub

Private Sub cdconnect_Click()

If txip = "" Then
  Exit Sub
End If
  
If ws1.State <> 0 Then
  ws1.Close
End If
If ws2.State <> 0 Then
  ws2.Close
End If
If ws3.State <> 0 Then
  ws3.Close
End If
If ws4.State <> 0 Then
  ws4.Close
End If

ws1.RemoteHost = txip
ws2.RemoteHost = txip
ws3.RemoteHost = txip
ws4.RemoteHost = txip
ws1.RemotePort = 2000
ws2.RemotePort = 2100
ws3.RemotePort = 2200
ws4.RemotePort = 2300
ws1.LocalPort = 0
ws2.LocalPort = 0
ws3.LocalPort = 0
ws4.LocalPort = 0
ws1.Connect
ws2.Connect
ws3.Connect
ws4.Connect



End Sub

Private Sub cddisconnect_Click()

If ws1.State <> 0 Then
  ws1.Close
End If
If ws2.State <> 0 Then
  ws2.Close
End If
If ws3.State <> 0 Then
  ws3.Close
End If
If ws4.State <> 0 Then
  ws4.Close
End If

op.Value = False
op.Enabled = True
txip.Enabled = True
cdconnect.Enabled = True
cddisconnect.Enabled = False

connected = False
tb1.TabIndex = 1
If tb1.TabIndex = 1 Then
  tb1.Tabs.Item(1).Selected = True
  fm1.Visible = True
  fm2.Visible = False
  fm3.Visible = False
  fm4.Visible = False
End If
txstatus = txstatus + "Disconnected.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)

End Sub

Private Sub cdpclear_Click()

igrecieve.Picture = LoadPicture("")
igrecieve.Top = 0
igrecieve.Left = 0
hs.Max = 0
vs.Max = 0


End Sub

Private Sub cdsave_Click()

On Error Resume Next

Dim filebox As OPENFILENAME  ' passes data to and from the function
Dim fname As String  ' receives path and filename of selected file
Dim retval As Long  ' return value

filebox.lStructSize = Len(filebox)  ' size of the structure
filebox.hwndOwner = Me.hWnd  ' handle of the window opening the box
filebox.lpstrTitle = "Save File"  ' text to display in the title bar
' Set the File Type drop-box values to Text Files and All Files
filebox.lpstrFilter = "BMP File" & vbNullChar & "*.BMP" & vbNullChar & vbNullChar
filebox.lpstrFile = Space(255)  ' receives path and filename of selected file
filebox.nMaxFile = 255  ' size of the path and filename buffer
filebox.lpstrFileTitle = Space(255)  ' receives filename of selected file
filebox.nMaxFileTitle = 255  ' size of the filename buffer
filebox.lpstrDefExt = "txt"  ' default file extension
' Allow only existing paths, warn if file already exists, hide read-only box
filebox.flags = OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

retval = GetSaveFileName(filebox)  ' open the Save File dialog box
If retval <> 0 Then  ' the user chose a file
  ' Extract the filename from the buffer and put it into fname
  fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
End If

precieve.AutoSize = True
precieve.Picture = igrecieve.Picture
SavePicture precieve.Image, fname

precieve = LoadPicture("")

precieve.AutoSize = False
precieve.Width = 6975
precieve.Height = 4575

txstatus = txstatus + "Picture recieved.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub cdsendfile_Click()

If txfilename = "" Then
  Exit Sub
End If

commanddatasend = Trim(Str(FileLen(txfilename)) + ">" + Trim("file" & Right(txfilename, 3)))



ws4.SendData commanddatasend

cmd = 2

cdsendfile.Enabled = False
txstatus3.Visible = True
pb3.Visible = True

sendfilename = txfilename

cdsendpicture.Enabled = False

txstatus = txstatus + "Sending file.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub cdsendpicture_Click()

openfile

If Len(filename) < 1 Then
  Exit Sub
End If



'txsend = Str(filesize)

commanddatasend = Trim(Str(filesize) + ">" + "picture")

ws4.SendData commanddatasend
cmd = 1


cdsendpicture.Enabled = False

txstatus1.Visible = True
pb1.Visible = True


cdsendfile.Enabled = False

txstatus = txstatus + "Sending Picture.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub dr1_Change()

fl1.Path = dr1.Path
End Sub

Private Sub dv1_Change()

On Error GoTo err

dr1.Path = dv1.Drive

err:
   dv1.Drive = dr1.Path

End Sub

Private Sub fl1_Click()

If Mid(dr1.Path, Len(dr1.Path), 1) <> "\" Then
  txfilename.Text = dr1.Path + "\" + fl1.List(fl1.ListIndex)
  Else
  txfilename.Text = dr1.Path + fl1.List(fl1.ListIndex)
End If

lbfilesize.Caption = FileLen(txfilename) & " Bytes."




End Sub

Private Sub Form_Load()

'If App.PrevInstance Then
'  Unload Me
'  End
'End If


If ws1.State <> 0 Then
  ws1.Close
End If

tb1.TabIndex = 0
If tb1.TabIndex = 0 Then
  fm1.Visible = True
  fm2.Visible = False
  fm3.Visible = False
  fm4.Visible = False
End If

totalrecieve = 1
totalrecieve1 = 1

connected = False

End Sub

Private Sub Form_Resize()

On Error GoTo err

Me.Width = 7590
Me.Height = 6540

err:
   Exit Sub
   
End Sub


Private Sub hs_Change()

igrecieve.Top = -(vs.Value)
igrecieve.Left = -(hs.Value)

End Sub

Private Sub hs_Scroll()

igrecieve.Top = -(vs.Value)
igrecieve.Left = -(hs.Value)

End Sub


Private Sub op_Click()

If ws1.State <> 0 Then
  ws1.Close
End If
If ws2.State <> 0 Then
  ws2.Close
End If
If ws3.State <> 0 Then
  ws3.Close
End If
If ws4.State <> 0 Then
  ws4.Close
End If

  ws1.LocalPort = 2000
  ws2.LocalPort = 2100
  ws3.LocalPort = 2200
  ws4.LocalPort = 2300
  ws1.Listen
  ws2.Listen
  ws3.Listen
  ws4.Listen
  
  txip.Enabled = False
  cdconnect.Enabled = False
  
  cddisconnect.Enabled = True
  
txstatus = txstatus + "Listning.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub


Private Sub tb1_Click()

If connected = False Then
 
  Exit Sub
End If


If tb1.SelectedItem.Index = 1 Then
  fm1.Visible = True
  fm2.Visible = False
  fm3.Visible = False
  fm4.Visible = False
End If
If tb1.SelectedItem.Index = 2 Then
  fm1.Visible = False
  fm2.Visible = True
  fm3.Visible = False
  fm4.Visible = False
End If
If tb1.SelectedItem.Index = 3 Then
  fm1.Visible = False
  fm2.Visible = False
  fm3.Visible = True
  fm4.Visible = False
End If
If tb1.SelectedItem.Index = 4 Then
  fm1.Visible = False
  fm2.Visible = False
  fm3.Visible = False
  fm4.Visible = True
End If


End Sub

Private Sub txsend_Change()


txsend.SelStart = Len(txsend)
typesend = txsend
ws1.SendData typesend

End Sub

Private Sub vs_Change()

igrecieve.Top = -(vs.Value)
igrecieve.Left = -(hs.Value)

End Sub

Private Sub vs_Scroll()

igrecieve.Top = -(vs.Value)
igrecieve.Left = -(hs.Value)

End Sub


Private Sub ws1_Close()

connected = False
tb1.TabIndex = 1
If tb1.TabIndex = 1 Then
  tb1.Tabs.Item(1).Selected = True
  fm1.Visible = True
  fm2.Visible = False
  fm3.Visible = False
  fm4.Visible = False
End If

op.Value = False
op.Enabled = True
txip.Enabled = True
cdconnect.Enabled = True
cddisconnect.Enabled = False

txstatus = txstatus + "Disconnected.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub ws1_Connect()

op.Enabled = False
txip.Enabled = False
cdconnect.Enabled = False
cddisconnect.Enabled = True

tb1.TabIndex = 2
If tb1.TabIndex = 2 Then
  tb1.Tabs.Item(2).Selected = True
  fm1.Visible = False
  fm2.Visible = True
  fm3.Visible = False
  fm4.Visible = False
End If

txsend.Enabled = True

connected = True

txstatus = txstatus + "Connected to.." + ws1.RemoteHostIP + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub ws1_ConnectionRequest(ByVal requestID As Long)

If ws1.State <> 0 Then
  ws1.Close
End If

ws1.Accept requestID


tb1.TabIndex = 2
If tb1.TabIndex = 2 Then
  tb1.Tabs.Item(2).Selected = True
  fm1.Visible = False
  fm2.Visible = True
  fm3.Visible = False
  fm4.Visible = False
End If

txsend.Enabled = True

connected = True

txstatus = txstatus + "Connected to.." + ws1.RemoteHostIP + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub


Private Sub ws1_DataArrival(ByVal bytesTotal As Long)




ws1.GetData typerecieve, vbString

txrecieve.Refresh
txrecieve = typerecieve + Chr(10)
txrecieve.SelStart = Len(txrecieve)



End Sub


Private Sub ws2_ConnectionRequest(ByVal requestID As Long)

If ws2.State <> 0 Then
  ws2.Close
End If

ws2.Accept requestID




End Sub

Private Sub ws2_DataArrival(ByVal bytesTotal As Long)





ws2.GetData picdatarecieve, vbString

If totalrecieve < picrecievesize Then
  Open App.Path + "\picture.tmp" For Binary Access Write As #1
  Put #1, totalrecieve, picdatarecieve
  Close #1
  totalrecieve = totalrecieve + bytesTotal
  txstatus2.Visible = True
  pb2.Visible = True
End If

pb2.Value = Int((totalrecieve / picrecievesize) * 100)
txstatus2 = "Recieving " & Str(pb2.Value) & "%"

If totalrecieve - 1 = picrecievesize Then
  igrecieve.Picture = LoadPicture(App.Path + "\picture.tmp")
  totalrecieve = 1
  Kill App.Path + "\picture.tmp"
  hs.Max = 0
  vs.Max = 0
  If igrecieve.Width > (precieve.Width / 15) Then
    hs.Max = (igrecieve.Width - (precieve.Width / 15)) + 20
    igrecieve.Top = 0
    igrecieve.Left = 0
  End If
  If igrecieve.Height > (precieve.Height / 15) Then
    vs.Max = (igrecieve.Height - (precieve.Height / 15)) + 20
    igrecieve.Top = 0
    igrecieve.Left = 0
  End If
  txstatus2.Visible = False
  pb2.Visible = False
  
  tb1.TabIndex = 4
  If tb1.TabIndex = 4 Then
    tb1.Tabs.Item(4).Selected = True
    fm1.Visible = False
    fm2.Visible = False
    fm3.Visible = False
    fm4.Visible = True
  End If
  
  
  Exit Sub
  
End If


Me.Refresh
End Sub


Private Sub ws2_SendComplete()

cdsendpicture.Enabled = True
o = 0

txstatus1.Visible = False
pb1.Visible = False

cdsendfile.Enabled = True

txstatus = txstatus + "Picture sended.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)


End Sub

Private Sub ws2_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

o = o + bytesSent
pb1value = Int((o / filesize) * 100)

pb1.Value = pb1value
txstatus1 = "Sending " & Str(pb1value) & "%"

End Sub

Private Sub ws3_ConnectionRequest(ByVal requestID As Long)

If ws3.State <> 0 Then
  ws3.Close
End If

ws3.Accept requestID

End Sub

Private Sub ws3_DataArrival(ByVal bytesTotal As Long)



ws3.GetData filedatarecieve, vbString


If totalrecieve1 < filerecievesize Then 'Len(totalrecieve1)
  Open App.Path + "\file.tmp" For Binary Access Write As #2
  Put #2, totalrecieve1, filedatarecieve
  Close #2
  totalrecieve1 = totalrecieve1 + bytesTotal
  txstatus4.Visible = True
  pb4.Visible = True
End If

pb4.Value = Int((totalrecieve1 / filerecievesize) * 100)
txstatus4 = "Recieving " & Str(pb4.Value) & "%"

If totalrecieve1 - 1 = filerecievesize Then
  
  totalrecieve1 = 1
  cdsendfile.Enabled = True
  savefile
  txstatus4.Visible = False
  pb4.Visible = False
  Exit Sub
End If


Me.Refresh


End Sub


Private Sub ws3_SendComplete()

cdsendfile.Enabled = True
o1 = 0

txstatus3.Visible = False
pb3.Visible = False

cdsendpicture.Enabled = True


txstatus = txstatus + "File sended.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)

End Sub

Private Sub ws3_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

o1 = o1 + bytesSent
pb3value = Int((o1 / FileLen(sendfilename)) * 100)

pb3.Value = pb3value
txstatus3 = "Sending " & Str(pb3value) & "%"



End Sub

Private Sub ws4_ConnectionRequest(ByVal requestID As Long)

If ws4.State <> 0 Then
  ws4.Close
End If

ws4.Accept requestID

End Sub

Private Sub ws4_DataArrival(ByVal bytesTotal As Long)


ws4.GetData commanddatarecieve, vbString


For t = 1 To Len(commanddatarecieve)
   If Mid(commanddatarecieve, t, 1) = ">" Then
     Exit For
   End If
Next t


If Mid(commanddatarecieve, t + 1, -(t - Len(commanddatarecieve))) = "picture" Then
  picrecievesize = Val(Mid(commanddatarecieve, 1, t - 1))
End If

'txsend = Trim(Right(commanddatarecieve, 3)) 'Mid(commanddatarecieve, t + 1, (-(t - Len(commanddatarecieve)) - 3))
If Mid(commanddatarecieve, t + 1, (-(t - Len(commanddatarecieve)) - 3)) = "file" Then
  filerecievesize = Val(Mid(commanddatarecieve, 1, t - 1))
  fileextention = Trim(Right(commanddatarecieve, 3))
  
End If


End Sub

Private Sub ws4_SendComplete()

If cmd = 1 Then
  Open filename For Binary Access Read As #1

  picdatasend = String(filesize, " ")
  Get #1, , picdatasend
  ws2.SendData picdatasend
  Close #1
  filename = ""
End If

If cmd = 2 Then
  Open txfilename For Binary Access Read As #1

  filedatasend = String(FileLen(txfilename), " ")
  Get #1, , filedatasend
  ws3.SendData filedatasend
  Close #1
  txfilename = ""
End If

End Sub

Public Sub savefile()



Dim filebox As OPENFILENAME  ' passes data to and from the function
Dim fname As String  ' receives path and filename of selected file
Dim retval As Long  ' return value

Dim a As String


filebox.lStructSize = Len(filebox)  ' size of the structure
filebox.hwndOwner = Me.hWnd  ' handle of the window opening the box
filebox.lpstrTitle = "Save File"  ' text to display in the title bar
' Set the File Type drop-box values to Text Files and All Files
filebox.lpstrFilter = Trim(fileextention) + " File" & vbNullChar & "*." + Trim(fileextention) & vbNullChar & vbNullChar
filebox.lpstrFile = Space(255)  ' receives path and filename of selected file
filebox.nMaxFile = 255  ' size of the path and filename buffer
filebox.lpstrFileTitle = Space(255)  ' receives filename of selected file
filebox.nMaxFileTitle = 255  ' size of the filename buffer
filebox.lpstrDefExt = "lll"  ' default file extension
' Allow only existing paths, warn if file already exists, hide read-only box
filebox.flags = OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

retval = GetSaveFileName(filebox)  ' open the Save File dialog box
If retval <> 0 Then  ' the user chose a file
  ' Extract the filename from the buffer and put it into fname
  fname = Left(filebox.lpstrFile, InStr(filebox.lpstrFile, vbNullChar) - 1)
End If

If Len(fname) < 1 Then
  Exit Sub
End If

Open App.Path + "\file.tmp" For Binary Access Read As #1
Open fname For Binary Access Write As #2
a = String(FileLen(App.Path + "\file.tmp"), " ")

Get #1, , a
Put #2, , a

Close #1, #2

Kill App.Path + "\file.tmp"

txstatus = txstatus + "File recieved.." + Chr(13) + Chr(10)
txstatus.SelStart = Len(txstatus)

End Sub
