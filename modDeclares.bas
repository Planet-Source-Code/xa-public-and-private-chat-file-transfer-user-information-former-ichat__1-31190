Attribute VB_Name = "modDeclares"
Type person
IP As String
Name As String
active As Boolean
End Type
Public Cone() As person
Public Function FindID(Name As String, Optional lscase As Boolean = False) As Long
If lscase Then
For i = 1 To frmServer.totalPlaces
If LCase(Cone(i).Name) = LCase(Name) Then FindID = i: Exit Function
Next i
Else
For i = 1 To frmServer.totalPlaces
If Cone(i).Name = Name Then FindID = i: Exit Function
Next i
End If
End Function
