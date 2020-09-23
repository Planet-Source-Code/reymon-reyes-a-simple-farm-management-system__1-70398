VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   2910
   ClientTop       =   2115
   ClientWidth     =   8775
   ClipControls    =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   6855
      Left            =   -600
      TabIndex        =   0
      Top             =   0
      Width           =   9585
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   4020
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Timer timer2 
         Index           =   0
         Interval        =   1
         Left            =   7440
         Top             =   3120
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1935
         Left            =   8400
         Shape           =   2  'Oval
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   4
         Height          =   615
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   -240
         Width           =   5535
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   4
         Height          =   620
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   3940
         Width           =   7335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ACME FRUITS CORPORATION"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   915
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Width           =   4500
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0 S. 2008"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   2640
         Width           =   2250
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "This  System version 1.0 S, 2008 is Exclusive only for Acme Fruits Corp."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   5
         Top             =   3600
         Width           =   8655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   8760
         TabIndex        =   4
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "The_Etc..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The_Etc... Development Team"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   3000
         Width           =   3390
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BackStyle       =   0  'Transparent
         Caption         =   "BANANA PRODUCTION"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   3360
         TabIndex        =   1
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   3225
         Left            =   840
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0080FFFF&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   4
         Height          =   375
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   3375
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2535
         Left            =   600
         Shape           =   2  'Oval
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1695
         Left            =   600
         Top             =   2880
         Width           =   855
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   1335
         Left            =   8880
         Top             =   0
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note: Please understand the forms before trying to understand the codes
'      so that it's much easier for us to trace the flow of the forms.
'      Also consider that there are 1 and a million ways to do something
'      so the codes here might be unorthodox but the IMPORTANT thing is that
'      the desired output is achieved.
'      The approach taken here is to use adodc instead of pure ado code.
'      Don't forget to place comments in your code so that other people who sees your
'      code doesn't suffer from painfully trying to understand it.
'      Ayaw liboga imo sarili kay daghan na ta! Sayuna lang para dili maglisod!
' Peace! :D
' the_Etc Development Team
' Mabuhay ang Pilipinas!

Option Explicit
Public StartUp As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If StartUp = False Then Unload Me
End Sub

Private Sub Form_Load()
'Label8.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'Label4.Caption = App.Title
If StartUp = False Then Exit Sub
frmSplash.MousePointer = vbHourglass
ProgressBar1.Visible = True
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub timer2_Timer(Index As Integer)
On Error Resume Next
ProgressBar1.Value = ProgressBar1 + 1
Label1.Caption = ProgressBar1.Value & "%"
frmSplash.MousePointer = vbHourglass
    
    If ProgressBar1 = 100 Then
        frmSplash.MousePointer = vbDefault
        ProgressBar1.Visible = False
        MDIForm1.Show
        timer2(0).Enabled = False
        StartUp = False
        Unload Me
    End If
End Sub
