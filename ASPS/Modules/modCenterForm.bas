Attribute VB_Name = "modCenterForm"

Option Explicit

Public Sub CenterForm(frm As Form, mdi As MDIForm)
frm.Left = (mdi.ScaleWidth - frm.Width) / 2
frm.Top = (mdi.ScaleHeight - frm.Height) / 2
End Sub
