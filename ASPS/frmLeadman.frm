VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLeadman 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leadman Profile"
   ClientHeight    =   5160
   ClientLeft      =   2955
   ClientTop       =   2745
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9315
   Begin MSAdodcLib.Adodc adoleadman 
      Height          =   375
      Left            =   7680
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "adoleadman"
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      Left            =   7800
      TabIndex        =   15
      ToolTipText     =   "Close this form"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelleadman 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
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
      Left            =   7800
      TabIndex        =   14
      ToolTipText     =   "Cancel the current operation"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteleadman 
      Caption         =   "&Delete"
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
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "Delete the current Lead Man"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditLeadman 
      Caption         =   "&Edit"
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
      Left            =   5160
      TabIndex        =   12
      ToolTipText     =   "Modify the current Lead Man"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLeadman 
      Caption         =   "&Add"
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
      Left            =   3840
      TabIndex        =   11
      ToolTipText     =   "Add a new Lead Man"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Fieldman_Address"
      DataSource      =   "adoleadman"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtmint 
      DataField       =   "Fieldman_MiddleInit"
      DataSource      =   "adoleadman"
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
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtFname 
      DataField       =   "Fieldman_Fname"
      DataSource      =   "adoleadman"
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtLname 
      DataField       =   "Fieldman_Lname"
      DataSource      =   "adoleadman"
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtleadman 
      DataField       =   "Fieldman_ID"
      DataSource      =   "adoleadman"
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdLeadman 
      Bindings        =   "frmLeadman.frx":0000
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8454143
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   975
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   6255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   975
      Left            =   4200
      Shape           =   2  'Oval
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   600
      Top             =   -120
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   360
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   855
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   4125
      TabIndex        =   5
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Initial:"
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
      Left            =   3630
      TabIndex        =   4
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
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
      Left            =   3855
      TabIndex        =   3
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
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
      Left            =   3870
      TabIndex        =   2
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Leadman ID:"
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
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmLeadman"
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
Public UpdateMode As Boolean

Private Sub adoleadman_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'prevents the display of errors generated by the adodc
fCancelDisplay = True
End Sub

Private Sub cmdAddLeadman_Click() 'adds a new record
On Error GoTo ErrHandler

If cmdAddLeadman.Caption = "&Add" Then
    UpdateMode = True
    cmdAddLeadman.Caption = "&Save"
    cmdDeleteleadman.Enabled = False
    cmdEditLeadman.Enabled = False
    cmdCancelleadman.Enabled = True
    cmdClose.Enabled = False
    grdLeadman.Enabled = False
    
   txtleadman.Locked = False
   txtLname.Locked = False
   txtFname.Locked = False
   txtmint.Locked = False
   txtAddress.Locked = False
    
   adoleadman.Recordset.AddNew
   txtLname.SetFocus
   
ElseIf cmdAddLeadman.Caption = "&Save" Then
    If txtLname.Text = "" Then
        MsgBox "Last Name is required."
        txtLname.SetFocus
        Exit Sub
    ElseIf txtFname.Text = "" Then
        MsgBox "First Name is required."
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdAddLeadman.Caption = "&Add"
        adoleadman.Recordset.Update
        'enable edit and delete
        cmdEditLeadman.Enabled = True
        cmdDeleteleadman.Enabled = True
        'disable cancel
        cmdCancelleadman.Enabled = False
        cmdClose.Enabled = True
        grdLeadman.Enabled = True
        'lock fields
         txtleadman.Locked = True
         txtLname.Locked = True
         txtFname.Locked = True
         txtmint.Locked = True
         txtAddress.Locked = True
        adoleadman.Recordset.Resync
        adoleadman.Refresh
        Me.MousePointer = vbDefault
        MsgBox "Record successfully updated."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    Me.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub cmdAddLeadman_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddLeadman.FontBold = True
cmdEditLeadman.FontBold = False
cmdDeleteleadman.FontBold = False
cmdCancelleadman.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdCancelleadman_Click() 'cancels the addition or editing of a record
On Error GoTo ErrHandler

adoleadman.Recordset.Cancel
adoleadman.Refresh
cmdAddLeadman.Caption = "&Add"
cmdEditLeadman.Enabled = True
cmdDeleteleadman.Enabled = True
cmdCancelleadman.Enabled = False
cmdClose.Enabled = True
grdLeadman.Enabled = True
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelleadman_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddLeadman.FontBold = False
cmdEditLeadman.FontBold = False
cmdDeleteleadman.FontBold = False
cmdCancelleadman.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub cmdClose_Click() 'close the form
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddLeadman.FontBold = False
cmdEditLeadman.FontBold = False
cmdDeleteleadman.FontBold = False
cmdCancelleadman.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdDeleteleadman_Click() 'deletes a record
On Error GoTo errorHandler

'present the user with a message box to confirm deletion of a record
If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    adoleadman.Recordset.Delete
    adoleadman.Recordset.Requery
    'adoleadman.Refresh
End If

Exit Sub
errorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        adoleadman.Refresh
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub cmdDeleteleadman_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddLeadman.FontBold = False
cmdEditLeadman.FontBold = False
cmdDeleteleadman.FontBold = True
cmdCancelleadman.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdEditLeadman_Click() 'edits a record
On Error GoTo ErrHandler

UpdateMode = True
'unlock fields
txtLname.Locked = False
txtFname.Locked = False
txtmint.Locked = False
txtAddress.Locked = False
cmdAddLeadman.Caption = "&Save"
cmdDeleteleadman.Enabled = False
cmdEditLeadman.Enabled = False
cmdCancelleadman.Enabled = True
cmdClose.Enabled = False
grdLeadman.Enabled = False
txtLname.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditLeadman_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddLeadman.FontBold = False
cmdEditLeadman.FontBold = True
cmdDeleteleadman.FontBold = False
cmdCancelleadman.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub Form_Deactivate()
If Me.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
'    MsgBox "Please save or cancel current operation to continue."
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Call CenterForm(frmLeadman, MDIForm1)
'setup the adodc
Call ConnectDB(adoleadman)
adoleadman.CommandType = adCmdTable
adoleadman.RecordSource = "FieldManNames"
adoleadman.Refresh
End Sub

Private Sub grdLeadman_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
