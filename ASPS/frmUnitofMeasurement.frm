VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUnitofMeasurement 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unit of Measurement"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUnitofMeasurement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      ToolTipText     =   "Modify the current record"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      ToolTipText     =   "Cancel current operation"
      Top             =   1200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoUOM 
      Height          =   330
      Left            =   840
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      Connect         =   $"frmUnitofMeasurement.frx":000C
      OLEDBString     =   $"frmUnitofMeasurement.frx":00A2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblunittype"
      Caption         =   "adoUOM"
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
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Close this form"
      Top             =   1800
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
      Left            =   2760
      TabIndex        =   7
      ToolTipText     =   "Delete the current record"
      Top             =   1200
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
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Add a new record"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2940
      TabIndex        =   5
      ToolTipText     =   "Go to the last record"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      ToolTipText     =   "Go to the next record"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      ToolTipText     =   "Go to the previous record"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Go to the first record"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtUnitofmeasurement 
      DataField       =   "Unit"
      DataSource      =   "adoUOM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2759
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   480
      Top             =   -240
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   240
      Shape           =   3  'Circle
      Top             =   -240
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   855
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   -360
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   615
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit of Measurement:"
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
      Left            =   601
      TabIndex        =   1
      Top             =   720
      Width           =   2025
   End
End
Attribute VB_Name = "frmUnitofMeasurement"
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

Private Sub adoUOM_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAdd_Click()
On Error GoTo ErrHandle
If cmdAdd.Caption = "&Add" Then
    UpdateMode = True
    cmdAdd.Caption = "&Save"
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    cmdClose.Enabled = False
    cmdCancel.Enabled = True
    txtUnitofmeasurement.Locked = False
    txtUnitofmeasurement.SetFocus
    adoUOM.Recordset.AddNew
ElseIf cmdAdd.Caption = "&Save" Then
    If txtUnitofmeasurement.Text = "" Or IsNull(txtUnitofmeasurement) Then
        MsgBox "Unit of measurement is required."
        txtUnitofmeasurement.SetFocus
        Exit Sub
    Else
        Me.MousePointer = vbHourglass
        cmdAdd.Caption = "&Add"
        adoUOM.Recordset.Update
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdClose.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        cmdCancel.Enabled = False
        txtUnitofmeasurement.Locked = True
        adoUOM.Recordset.Resync
        'adoUOM.Refresh
        Me.MousePointer = vbDefault
        MsgBox "Record updated successfully."
        UpdateMode = False
    End If
End If

Exit Sub
ErrHandle:
    MsgBox Err.Description
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.FontBold = True
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandler

adoUOM.Recordset.Cancel
adoUOM.Refresh
cmdAdd.Caption = "&Add"
txtUnitofmeasurement.Locked = True
cmdEdit.Enabled = True
cmdFirst.Enabled = True
cmdPrevious.Enabled = True
cmdNext.Enabled = True
cmdLast.Enabled = True
cmdClose.Enabled = True
cmdEdit.Enabled = True
cmdDelete.Enabled = True
cmdCancel.Enabled = False
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
On Error GoTo errorHandler

If MsgBox("The selected RECORD will be deleted, are you sure you want to continue?" & vbCr & "This action is irreversible.", vbOKCancel + vbExclamation, "Delete") = vbOK Then
    adoUOM.Recordset.Delete
    adoUOM.Recordset.Requery
End If

Exit Sub
errorHandler:
    If Err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it is being used in" & vbCr & _
            "one or more existing records."
        adoUOM.Refresh
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
cmdAdd.Caption = "&Save"
cmdDelete.Enabled = False
cmdFirst.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdLast.Enabled = False
cmdClose.Enabled = False
cmdCancel.Enabled = True
txtUnitofmeasurement.Locked = False
txtUnitofmeasurement.SetFocus
cmdEdit.Enabled = False

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdit.FontBold = True
End Sub

Private Sub cmdFirst_Click()
adoUOM.Recordset.MoveFirst
cmdNext.Enabled = True
If adoUOM.Recordset.BOF Then
    cmdFirst.Enabled = False
    cmdLast.Enabled = True
End If
End Sub

Private Sub cmdLast_Click()
adoUOM.Recordset.MoveLast
cmdPrevious.Enabled = True
If adoUOM.Recordset.EOF Then
    cmdFirst.Enabled = True
    cmdLast.Enabled = False
End If
End Sub

Private Sub cmdNext_Click()
cmdPrevious.Enabled = True
adoUOM.Recordset.MoveNext
If adoUOM.Recordset.EOF Then
    cmdNext.Enabled = False
    adoUOM.Recordset.MoveLast
End If
End Sub

Private Sub cmdPrevious_Click()
cmdNext.Enabled = True
adoUOM.Recordset.MovePrevious
If adoUOM.Recordset.BOF Then
    cmdPrevious.Enabled = False
    cmdLast.Enabled = True
    adoUOM.Recordset.MoveFirst
End If
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
Call CenterForm(frmUnitofMeasurement, MDIForm1)
Call ConnectDB(adoUOM)
adoUOM.CommandType = adCmdTable
adoUOM.RecordSource = "tblunittype"
adoUOM.Refresh
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAdd.FontBold = False
cmdCancel.FontBold = False
cmdClose.FontBold = False
cmdDelete.FontBold = False
cmdEdit.FontBold = False
End Sub
