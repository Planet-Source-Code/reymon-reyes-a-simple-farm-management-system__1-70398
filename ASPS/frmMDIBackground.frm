VERSION 5.00
Begin VB.Form frmMDIBackground 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9435
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   14910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   9960
      Left            =   0
      Picture         =   "frmMDIBackground.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frmMDIBackground"
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

Private Sub Form_Activate()
Call CenterForm(frmMDIBackground, MDIForm1)
Me.ZOrder (vbSendToBack)
End Sub

Private Sub Form_Load()
Call CenterForm(frmMDIBackground, MDIForm1)
End Sub
