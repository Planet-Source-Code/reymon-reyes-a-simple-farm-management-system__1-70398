Attribute VB_Name = "modAdoconnect"
Option Explicit
Public user As String

Public Sub ConnectDB(adoControl As Adodc)
On Error Resume Next
adoControl.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\AmadeusFarm.mdb;Persist Security Info=False;Jet OLEDB:Database Password=a"
adoControl.Refresh
End Sub

