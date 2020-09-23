VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRptWithdrawals 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Withdrawals"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5295
   Begin VB.CommandButton cmdPreviewWithdrawals 
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
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
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
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Withdrawals by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin MSComCtl2.DTPicker pkrTo 
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSComCtl2.DTPicker pkrFrom 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.TextBox txtAnnual 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtMonthlyYear 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker pkrDaily 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.OptionButton optFromTo 
         BackColor       =   &H00C0FFC0&
         Caption         =   "From       :"
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
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton optAnnually 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Annually :"
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
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optMonthly 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Monthly  :"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optDaily 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Daily       :"
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
         TabIndex        =   1
         Top             =   480
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
         TabIndex        =   14
         Top             =   1920
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   5295
   End
End
Attribute VB_Name = "frmRptWithdrawals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub LoadMonth()
cboMonth.AddItem "January"
cboMonth.AddItem "February"
cboMonth.AddItem "March"
cboMonth.AddItem "April"
cboMonth.AddItem "May"
cboMonth.AddItem "June"
cboMonth.AddItem "July"
cboMonth.AddItem "August"
cboMonth.AddItem "September"
cboMonth.AddItem "October"
cboMonth.AddItem "November"
cboMonth.AddItem "December"
cboMonth.ListIndex = 0
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewWithdrawals.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdPreviewWithdrawals_Click()
Dim strMonth As String, strYear As String
Dim FromDate As String, ToDate As String

On Error GoTo ErrHandler

If optDaily.Value = True Then
    Me.MousePointer = vbHourglass
    If envAmadeus.rscmdRptRequestDaily.State = adStateOpen Then
        envAmadeus.rscmdRptRequestDaily.Close
        Unload rptRequestDaily
    End If

    envAmadeus.cmdRptRequestDaily pkrDaily.Value
    
    If envAmadeus.rscmdRptRequestDaily.RecordCount = 0 Then
        envAmadeus.rscmdRptRequestDaily.Close
        MsgBox "Cannot find record."
        pkrDaily.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    rptRequestDaily.Show
    rptRequestDaily.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optMonthly.Value = True Then
    Me.MousePointer = vbHourglass
    If txtMonthlyYear.Text = "" Then
        MsgBox "Invalid year."
        txtMonthlyYear.SetFocus
        Exit Sub
    End If
    
    If envAmadeus.rscmdRptRequestMonthlyYearly.State = adStateOpen Then
        envAmadeus.rscmdRptRequestMonthlyYearly.Close
        Unload rptRequestMonthlyYearly
    End If

    envAmadeus.cmdRptRequestMonthlyYearly 0, cboMonth.ListIndex + 1, txtMonthlyYear.Text
    
    If envAmadeus.rscmdRptRequestMonthlyYearly.RecordCount = 0 Then
        envAmadeus.rscmdRptRequestMonthlyYearly.Close
        MsgBox "Cannot find record."
        txtMonthlyYear.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    strMonth = MonthName(Month(envAmadeus.rscmdRptRequestMonthlyYearly.Fields("Date")))
    strYear = Year(envAmadeus.rscmdRptRequestMonthlyYearly.Fields("Date"))
    rptRequestMonthlyYearly.Title = "Material Withdrawals as of " & strMonth & " " & strYear
    rptRequestMonthlyYearly.Show
    rptRequestMonthlyYearly.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optAnnually.Value = True Then
    Me.MousePointer = vbHourglass
    If envAmadeus.rscmdRptRequestMonthlyYearly.State = adStateOpen Then
        envAmadeus.rscmdRptRequestMonthlyYearly.Close
        Unload rptRequestMonthlyYearly
    End If

    envAmadeus.cmdRptRequestMonthlyYearly txtAnnual.Text, 0, 0
    
    If envAmadeus.rscmdRptRequestMonthlyYearly.RecordCount = 0 Then
        envAmadeus.rscmdRptRequestMonthlyYearly.Close
        MsgBox "Cannot find record."
        txtAnnual.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    strYear = Year(envAmadeus.rscmdRptRequestMonthlyYearly.Fields("Date"))
    rptRequestMonthlyYearly.Title = "Material Withdrawals as of " & strYear
    rptRequestMonthlyYearly.Show
    rptRequestMonthlyYearly.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optFromTo.Value = True Then
    Me.MousePointer = vbHourglass
    If envAmadeus.rscmdRptRequestRange.State = adStateOpen Then
        envAmadeus.rscmdRptRequestRange.Close
        Unload rptRequestRange
    End If

    envAmadeus.cmdRptRequestRange pkrFrom.Value, pkrTo.Value
    
    If envAmadeus.rscmdRptRequestRange.RecordCount = 0 Then
        envAmadeus.rscmdRptRequestRange.Close
        MsgBox "Cannot find record."
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    envAmadeus.rscmdRptRequestRange.MoveFirst
    FromDate = envAmadeus.rscmdRptRequestRange.Fields("Date")
    envAmadeus.rscmdRptRequestRange.MoveLast
    ToDate = envAmadeus.rscmdRptRequestRange.Fields("Date")
    'envAmadeus.rscmdRptGrowerDateFromTo.Resync adAffectCurrent, adResyncAllValues
    rptRequestRange.Title = "Material Withdrawals from " & FromDate & " to " & ToDate
    rptRequestRange.Show
    rptRequestRange.WindowState = vbMaximized
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPreviewWithdrawals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewWithdrawals.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub Form_Load()
Call CenterForm(frmRptWithdrawals, MDIForm1)
Call LoadMonth
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewWithdrawals.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub txtAnnual_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub txtMonthlyYear_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
