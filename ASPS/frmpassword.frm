VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPassword 
   BorderStyle     =   0  'None
   Caption         =   "Login Details"
   ClientHeight    =   3240
   ClientLeft      =   4860
   ClientTop       =   3075
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2768
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2400
         Width           =   1185
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1448
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   1185
      End
      Begin MSAdodcLib.Adodc adopassword 
         Height          =   330
         Left            =   2280
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "adopassword"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txtUserName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataField       =   "username"
         DataSource      =   "Adodc1"
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
         Height          =   435
         Left            =   1763
         TabIndex        =   1
         Top             =   1320
         Width           =   2325
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataField       =   "password"
         DataSource      =   "Adodc1"
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
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1763
         PasswordChar    =   "#"
         TabIndex        =   2
         Top             =   1800
         Width           =   2325
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   4163
         Picture         =   "frmpassword.frx":0000
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enter User Name and Password to access the system"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   690
         Left            =   248
         TabIndex        =   7
         Top             =   480
         Width           =   4905
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   2055
         Left            =   4440
         Shape           =   2  'Oval
         Top             =   150
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   0
         Left            =   518
         TabIndex        =   6
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Index           =   1
         Left            =   623
         TabIndex        =   5
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0FFC0&
         Height          =   2055
         Left            =   0
         Shape           =   2  'Oval
         Top             =   1020
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000C000&
         BorderWidth     =   15
         X1              =   0
         X2              =   5425
         Y1              =   25
         Y2              =   25
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         BorderWidth     =   15
         X1              =   0
         X2              =   5425
         Y1              =   3175
         Y2              =   3175
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   0
         Top             =   2280
         Width           =   495
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000C000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   4920
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "username"
      DataSource      =   "adopassword"
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Private iCount As Integer

Private Sub adopassword_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
End Sub
Private Sub log()
On Error GoTo ErrHandler

With adopassword.Recordset
.MoveFirst
    Do Until .EOF
        If txtUserName.Text = .Fields("Username") And txtPassword.Text = .Fields("Password") Then
            modAdoconnect.user = txtUserName.Text
            Unload Me
            frmSplash.Show
            frmSplash.StartUp = True
            Exit Sub
        Else
            .MoveNext
        End If
    Loop
    
    iCount = iCount - 1
    MsgBox "Invalid username/password." & vbCr & iCount & " attempts left."
    txtUserName.SetFocus
    SendKeys "{Home}+{End}"
If iCount = 0 Then
    MsgBox "The system will now terminate"
    Unload Me
End If
End With


Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdOK_Click()
log
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

iCount = 3
Call ConnectDB(adopassword)
adopassword.CommandType = adCmdTable
adopassword.RecordSource = "tblpassword"
adopassword.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

