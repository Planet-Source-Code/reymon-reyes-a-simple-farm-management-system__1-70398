VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSuppliers 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   3810
   ClientLeft      =   5445
   ClientTop       =   4755
   ClientWidth     =   5685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5685
   Begin VB.TextBox txtContactNum 
      DataField       =   "ContactNum"
      DataSource      =   "adoSuppliers"
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin MSAdodcLib.Adodc adoSuppliers 
      Height          =   375
      Left            =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmSuppliers.frx":0000
      OLEDBString     =   $"frmSuppliers.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblSupplier"
      Caption         =   "adosuppliers"
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
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3135
      TabIndex        =   11
      ToolTipText     =   "Move to the last record"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2175
      TabIndex        =   10
      ToolTipText     =   "Move to the next record"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1215
      TabIndex        =   9
      ToolTipText     =   "Move to the previous record"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   255
      TabIndex        =   8
      ToolTipText     =   "Move to the first record"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   4215
      TabIndex        =   12
      ToolTipText     =   "Close this form"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   4215
      TabIndex        =   7
      ToolTipText     =   "Cancel current operation"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   2895
      TabIndex        =   6
      ToolTipText     =   "Delete the current record"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
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
      Left            =   1575
      TabIndex        =   5
      ToolTipText     =   "Modify the current record"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
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
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Add a new record"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "Address"
      DataSource      =   "adoSuppliers"
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtSupplier 
      DataField       =   "Name"
      DataSource      =   "adoSuppliers"
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox txtSupplierID 
      DataField       =   "Supplier_ID"
      DataSource      =   "adoSuppliers"
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact #:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   945
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
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   855
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   240
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   480
      Top             =   -120
      Width           =   3255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   975
      Left            =   1680
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Left            =   390
      TabIndex        =   15
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
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
      Left            =   330
      TabIndex        =   14
      Top             =   960
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SupplierID:"
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
      Left            =   135
      TabIndex        =   13
      Top             =   480
      Width           =   1065
   End
End
Attribute VB_Name = "frmSuppliers"
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

Private Sub adoSuppliers_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrHandler

If cmdAdd.Caption = "&Add" Then
    UpdateMode = True
    cmdAdd.Caption = "&Save"
    'unlock fields
    txtSupplier.Locked = False
    txtAddress.Locked = False
    txtContactNum.Locked = False
    'disable edit, delete
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdClose.Enabled = False
    
    cmdCancel.Enabled = True
    adoSuppliers.Recordset.AddNew
    txtSupplier.SetFocus
ElseIf cmdAdd.Caption = "&Save" Then
    If txtSupplier.Text = "" Then
        MsgBox "Supplier is required."
        txtSupplier.SetFocus
        Exit Sub
    End If
    If txtAddress.Text = "" Then
        MsgBox "Address is required."
        txtSupplier.SetFocus
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    cmdAdd.Caption = "&Add"
    'update the record
    adoSuppliers.Recordset.Update
    'lock fields
    txtSupplier.Locked = True
    txtAddress.Locked = True
    txtContactNum.Locked = True
    'enable edit, delete buttons
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    cmdClose.Enabled = True
    cmdCancel.Enabled = False
    adoSuppliers.Recordset.Resync
    'adoSuppliers.Recordset.MoveLast
    Me.MousePointer = vbDefault
    MsgBox "Record updated successfully."
    UpdateMode = False
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.FontBold = True
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandler

adoSuppliers.Recordset.Cancel
adoSuppliers.Refresh
'lock fields
txtSupplier.Locked = True
txtAddress.Locked = True
txtContactNum.Locked = True
'enable edit, delete buttons
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdAdd.Caption = "&Add"
cmdCancel.Enabled = False
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdClose.Enabled = True
UpdateMode = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCancel.FontBold = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose.FontBold = True
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbQuestion, "Delete") = vbOK Then
    Me.MousePointer = vbHourglass
    adoSuppliers.Recordset.Delete
    adoSuppliers.Recordset.Requery
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
If Err.Number = -2147467259 Then
        Me.MousePointer = vbDefault
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        adoSuppliers.Refresh
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdDelete.FontBold = True
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrHandler

UpdateMode = True
'unlock fields
txtSupplier.Locked = False
txtAddress.Locked = False
txtContactNum.Locked = False
'disable delete
cmdDelete.Enabled = False
cmdAdd.Caption = "&Save"
cmdCancel.Enabled = True
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdClose.Enabled = False
cmdEdit.Enabled = False
txtSupplier.SetFocus

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdit.FontBold = True
End Sub

Private Sub cmdFirst_Click()
On Error Resume Next
cmdNext.Enabled = True
adoSuppliers.Recordset.MoveFirst
End Sub

Private Sub cmdLast_Click()
cmdPrevious.Enabled = True
adoSuppliers.Recordset.MoveLast
End Sub

Private Sub cmdNext_Click()
adoSuppliers.Recordset.MoveNext
If adoSuppliers.Recordset.EOF Then
    adoSuppliers.Recordset.MoveLast
    cmdPrevious.Enabled = True
    cmdNext.Enabled = False
Else
    cmdPrevious.Enabled = True
End If
End Sub

Private Sub cmdPrevious_Click()
adoSuppliers.Recordset.MovePrevious
If adoSuppliers.Recordset.BOF Then
    adoSuppliers.Recordset.MoveFirst
    cmdPrevious.Enabled = False
    cmdNext.Enabled = True
Else
    cmdNext.Enabled = True
End If
End Sub

Private Sub Form_Deactivate()
If frmSuppliers.UpdateMode = True Then
    Me.ZOrder (vbBringToFront)
'    MsgBox "Please save or cancel current operation to continue."
    Me.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Call CenterForm(frmSuppliers, MDIForm1)
Call ConnectDB(adoSuppliers)
adoSuppliers.CommandType = adCmdTable
adoSuppliers.RecordSource = "tblSupplier"
adoSuppliers.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = False
cmdDelete.FontBold = False
cmdEdit.FontBold = False
End Sub
