VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGrower 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Growers"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   ControlBox      =   0   'False
   FillColor       =   &H0000C000&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   10020
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataSource      =   "adoGrower"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancelGrower 
      Caption         =   "Ca&ncel"
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
      Left            =   8640
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoGrower 
      Height          =   375
      Left            =   5280
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "adogrower"
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
   Begin MSDataGridLib.DataGrid grdGrower 
      Bindings        =   "frmGrower.frx":0000
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
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
         DataField       =   "Grower"
         Caption         =   "Grower"
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
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4410.142
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEditGrower 
      Caption         =   "E&dit"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdCloseGrower 
      Caption         =   "&Close"
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
      Left            =   8640
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeleteGrower 
      Caption         =   "Delete"
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
      Left            =   7440
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddGrower 
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
      Left            =   5040
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtGrowerName 
      DataField       =   "Grower"
      DataSource      =   "adoGrower"
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtGrowerID 
      DataField       =   "Growers_ID"
      DataSource      =   "adoGrower"
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   5565
      TabIndex        =   11
      Top             =   1560
      Width           =   810
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   855
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   855
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   720
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   960
      Top             =   -120
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   975
      Left            =   3480
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grower Name:"
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
      Left            =   4980
      TabIndex        =   10
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GrowerID:"
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
      Left            =   5415
      TabIndex        =   9
      Top             =   480
      Width           =   960
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   5775
   End
End
Attribute VB_Name = "frmGrower"
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

Private Sub adoGrower_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'prevents the display of errors generated by the adodc
fCancelDisplay = True
End Sub

Private Sub cmdAddGrower_Click() 'adds a new record
On Error GoTo ErrHandler

If cmdAddGrower.Caption = "&Add" Then
    UpdateMode = True
    cmdAddGrower.Caption = "&Save"
    'disable edit and delete
    cmdEditGrower.Enabled = False
    cmdDeleteGrower.Enabled = False
    'enable cancel
    cmdCancelGrower.Enabled = True
    cmdCloseGrower.Enabled = False
    grdGrower.Enabled = False
    'unlock fields
    txtGrowerName.Locked = False
    txtAddress.Locked = False
    'set ado to addnew
    adoGrower.Recordset.AddNew
    txtGrowerName.SetFocus
ElseIf cmdAddGrower.Caption = "&Save" Then
    If txtGrowerName.Text = "" Then
        MsgBox "Grower Name is required."
        txtGrowerName.SetFocus
        Exit Sub
    Else
    Me.MousePointer = vbHourglass
    cmdAddGrower.Caption = "&Add"
    adoGrower.Recordset.Update
    'enable edit and delete
    cmdEditGrower.Enabled = True
    cmdDeleteGrower.Enabled = True
    'disable cancel
    cmdCancelGrower.Enabled = False
    cmdCloseGrower.Enabled = True
    grdGrower.Enabled = True
    'lock fields
    txtGrowerName.Locked = True
    txtAddress.Locked = True
    adoGrower.Recordset.Resync
    'adoGrower.Refresh
    Me.MousePointer = vbDefault
    MsgBox "Record successfully updated."
    UpdateMode = False
    End If
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdAddGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = True
cmdEditGrower.FontBold = False
cmdDeleteGrower.FontBold = False
cmdCancelGrower.FontBold = False
cmdCloseGrower.FontBold = False
End Sub

Private Sub cmdCancelGrower_Click() 'cancels the addition or editing of a record
On Error GoTo ErrHandler

UpdateMode = False
adoGrower.Recordset.Cancel
adoGrower.Refresh
cmdAddGrower.Caption = "&Add"
cmdEditGrower.Enabled = True
cmdDeleteGrower.Enabled = True
cmdCancelGrower.Enabled = False
cmdCloseGrower.Enabled = True
grdGrower.Enabled = True
txtGrowerName.Locked = True
txtAddress.Locked = True

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancelGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = False
cmdEditGrower.FontBold = False
cmdDeleteGrower.FontBold = False
cmdCancelGrower.FontBold = True
cmdCloseGrower.FontBold = False
End Sub

Private Sub cmdCloseGrower_Click() 'close the form
Unload Me
End Sub

Private Sub cmdCloseGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = False
cmdEditGrower.FontBold = False
cmdDeleteGrower.FontBold = False
cmdCancelGrower.FontBold = False
cmdCloseGrower.FontBold = True
End Sub

Private Sub cmdDeleteGrower_Click() 'delete a record
On Error GoTo errorHandler

'present the user with a message box to confirm deletion of a record
If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    adoGrower.Recordset.Delete
    adoGrower.Recordset.Requery
End If

Exit Sub
errorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        adoGrower.Refresh
    Else
        MsgBox Err.Description
    End If

End Sub

Private Sub cmdDeleteGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = False
cmdEditGrower.FontBold = False
cmdDeleteGrower.FontBold = True
cmdCancelGrower.FontBold = False
cmdCloseGrower.FontBold = False
End Sub

Private Sub cmdEditGrower_Click() 'edit a record
On Error GoTo ErrHandler

UpdateMode = True
cmdAddGrower.Caption = "&Save"
cmdDeleteGrower.Enabled = False
cmdEditGrower.Enabled = False
cmdCancelGrower.Enabled = True
cmdCloseGrower.Enabled = False
grdGrower.Enabled = False
txtGrowerName.Locked = False
txtAddress.Locked = False
txtGrowerName.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEditGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = False
cmdEditGrower.FontBold = True
cmdDeleteGrower.FontBold = False
cmdCancelGrower.FontBold = False
cmdCloseGrower.FontBold = False
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
Call CenterForm(frmGrower, MDIForm1)
'setup the adodc
Call ConnectDB(adoGrower)
adoGrower.CommandType = adCmdTable
adoGrower.RecordSource = "tblGrower"
adoGrower.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddGrower.FontBold = False
cmdEditGrower.FontBold = False
cmdDeleteGrower.FontBold = False
cmdCancelGrower.FontBold = False
cmdCloseGrower.FontBold = False
End Sub

Private Sub grdGrower_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub
