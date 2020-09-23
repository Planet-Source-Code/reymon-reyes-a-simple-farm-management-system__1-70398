VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptGrower 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Withdrawals of Growers"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5175
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin MSComCtl2.DTPicker pkrDateTo 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39506
      End
      Begin MSComCtl2.DTPicker pkrDateFrom 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39506
      End
      Begin MSComCtl2.DTPicker pkrDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   39506
      End
      Begin VB.OptionButton optDateFromTo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date From:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date         :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   960
         Width           =   180
      End
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
      Left            =   3840
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdPreviewGrower 
      Caption         =   "Preview"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoGrowers 
      Height          =   375
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "adoGrowers"
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
   Begin MSDataListLib.DataCombo dcboGrowers 
      Bindings        =   "frmRptGrower.frx":0000
      Height          =   360
      Left            =   1620
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Grower"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   735
      Index           =   1
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   615
      Index           =   0
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   135
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grower:"
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
      Left            =   780
      TabIndex        =   3
      Top             =   360
      Width           =   765
   End
End
Attribute VB_Name = "frmRptGrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loadGrower As Boolean

Private Sub adoGrowers_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewGrower.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdPreviewGrower_Click()
Dim FromDate As String, ToDate As String

On Error GoTo ErrHandler

If dcboGrowers.Text = "" Then
    MsgBox "Please select a Grower."
    Exit Sub
End If

If optDate.Value = True Then
    Me.MousePointer = vbHourglass
    If envAmadeus.rscmdRptGrowerDate.State = adStateOpen Then
        envAmadeus.rscmdRptGrowerDate.Close
        Unload rptGrower
    End If
    
    envAmadeus.cmdRptGrowerDate adoGrowers.Recordset.Fields("Growers_ID"), pkrDate.Value
    
    If envAmadeus.rscmdRptGrowerDate.RecordCount = 0 Then
        envAmadeus.rscmdRptGrowerDate.Close
        MsgBox "Cannot find record."
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    rptGrower.Show
    rptGrower.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optDateFromTo.Value = True Then
    Me.MousePointer = vbHourglass
    If envAmadeus.rscmdRptGrowerDateFromTo.State = adStateOpen Then
        envAmadeus.rscmdRptGrowerDateFromTo.Close
        Unload rptGrowerFromTo
    End If
    
    envAmadeus.cmdRptGrowerDateFromTo adoGrowers.Recordset.Fields("Growers_ID"), pkrDateFrom.Value, pkrDateTo.Value
    
    If envAmadeus.rscmdRptGrowerDateFromTo.RecordCount = 0 Then
        envAmadeus.rscmdRptGrowerDateFromTo.Close
        MsgBox "Cannot find record."
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    envAmadeus.rscmdRptGrowerDateFromTo.MoveFirst
    FromDate = envAmadeus.rscmdRptGrowerDateFromTo.Fields("Date")
    envAmadeus.rscmdRptGrowerDateFromTo.MoveLast
    ToDate = envAmadeus.rscmdRptGrowerDateFromTo.Fields("Date")
    'envAmadeus.rscmdRptGrowerDateFromTo.Resync adAffectCurrent, adResyncAllValues
    rptGrowerFromTo.Title = "Grower Material Withdrawals from " & FromDate & " to " & ToDate
    rptGrowerFromTo.Show
    rptGrowerFromTo.WindowState = vbMaximized
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPreviewGrower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewGrower.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub dcboGrowers_Change()
If loadGrower = False Then Exit Sub
If adoGrowers.Recordset.Bookmark = 0 Then Exit Sub
adoGrowers.Recordset.Bookmark = dcboGrowers.SelectedItem
'MsgBox adoGrowers.Recordset.Fields("Growers_ID")
End Sub

Private Sub dcboGrowers_GotFocus()
loadGrower = True
End Sub

Private Sub dcboGrowers_LostFocus()
loadGrower = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

Call CenterForm(frmRptGrower, MDIForm1)
loadGrower = False

Call ConnectDB(adoGrowers)
adoGrowers.CommandType = adCmdTable
adoGrowers.RecordSource = "tblGrower"
adoGrowers.Refresh
Set dcboGrowers.RowSource = adoGrowers.Recordset
dcboGrowers.ListField = "Grower"

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewGrower.FontBold = False
cmdClose.FontBold = False
End Sub
