VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChangePassword 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4815
   Begin VB.TextBox txtUsername 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtUserID 
      DataField       =   "id"
      DataSource      =   "adoPassword"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoPassword 
      Height          =   375
      Left            =   120
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      Caption         =   "adoPassword"
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
   Begin VB.TextBox txtReTypePassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtNewPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "#"
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtOldPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Index           =   1
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   -360
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Index           =   0
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   810
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5295
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-Type Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   465
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   555
      TabIndex        =   5
      Top             =   360
      Width           =   1365
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adoPassword_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdChangePassword.FontBold = False
cmdCancel.FontBold = True
End Sub

Private Sub cmdChangePassword_Click()
On Error GoTo ErrHandler

If txtOldPassword.Text = "" Then MsgBox "Please enter Old Password.": txtOldPassword.SetFocus: Exit Sub
If txtNewPassword.Text = "" Then MsgBox "Please enter New Password.": txtNewPassword.SetFocus: Exit Sub
If txtReTypePassword.Text = "" Then MsgBox "Please Re-Type Password.": txtReTypePassword.SetFocus: Exit Sub
If txtNewPassword.Text <> txtReTypePassword.Text Then
    MsgBox "Either the new password or re-type password does not match"
    txtNewPassword.SetFocus
    Exit Sub
End If
If Len(txtNewPassword.Text) < 6 Then
    MsgBox "A new password requires a minimum of 6 characters."
    txtNewPassword.SetFocus
    Exit Sub
End If
With adoPassword.Recordset
.MoveFirst
    Do Until .EOF
        If modAdoconnect.user = adoPassword.Recordset.Fields("username") And txtOldPassword.Text = adoPassword.Recordset.Fields("password") Then
            adoPassword.Recordset.Fields("password") = txtNewPassword.Text
            adoPassword.Recordset.Update
            adoPassword.Refresh
            
            txtOldPassword.Text = ""
            txtNewPassword.Text = ""
            txtReTypePassword.Text = ""
            
            MsgBox "Password changed successfully."
            Exit Sub
        Else
            .MoveNext
        End If
    Loop
    
If .EOF Then
    MsgBox "Invalid Username/Password."
    txtOldPassword.SetFocus
    SendKeys "{Home}+{End}"
End If
End With

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdChangePassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdChangePassword.FontBold = True
cmdCancel.FontBold = False
End Sub

Private Sub Form_Load()
Call CenterForm(frmChangePassword, MDIForm1)
Call ConnectDB(adoPassword)
adoPassword.CommandType = adCmdTable
adoPassword.RecordSource = "tblpassword"
adoPassword.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdChangePassword.FontBold = False
cmdCancel.FontBold = False
End Sub
