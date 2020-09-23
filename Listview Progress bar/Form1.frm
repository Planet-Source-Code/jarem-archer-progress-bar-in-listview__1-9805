VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress Bar inside ListView"
   ClientHeight    =   3585
   ClientLeft      =   2910
   ClientTop       =   2640
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "GO! "
      Height          =   630
      Left            =   390
      TabIndex        =   0
      Top             =   300
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2010
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   210
      ScaleWidth      =   2085
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   75
      TabIndex        =   2
      Top             =   1110
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4260
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Downloading:"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Progress"
         Object.Width           =   3705
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3555
      Top             =   2085
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   435
      Top             =   1500
      Width           =   2340
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentPercent As Integer
Private Sub Command1_Click()
Dim lstEntry As ListItem
ListView1.ListItems.Add , , "Q3Demo.zip"
ListView1.ListItems.Item(1).SubItems(1) = ""
'lstEntry.SubItems(1) = ""
'lstEntry.ListSubItems.Item(1).ReportIcon = 0
Timer1.Enabled = True
End Sub

Public Function UpdateProgress(pb As Control, ByVal Percent)
'Replacement for progress bar..looks nicer also
Dim Num$ 'use percent
If Not pb.AutoRedraw Then 'picture in memory ?
pb.AutoRedraw = -1 'no, make one
End If
pb.Cls 'clear picture in memory
pb.ScaleWidth = 100 'new sclaemodus
pb.DrawMode = 10 'not XOR Pen Modus
Num$ = Format$(Percent, "###") + "%"
pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
pb.Print Num$ 'print percent
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh 'show differents
End Function

Private Sub Timer1_Timer()
'On Error Resume Next
CurrentPercent = CurrentPercent + 1
If CurrentPercent < 101 Then
UpdateProgress Picture1, CurrentPercent

'First unbound
ListView1.SmallIcons = Nothing

'Next Make changes to imagelist1
ImageList1.ListImages.Clear
ImageList1.ListImages.Add , , Picture1.Image

'After Rebound it listview
ListView1.SmallIcons = ImageList1

ListView1.ListItems.Item(1).ListSubItems.Item(1).ReportIcon = 1

Else
Timer1.Enabled = False
CurrentPercent = 0
End If

End Sub
