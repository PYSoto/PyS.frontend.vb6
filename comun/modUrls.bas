Attribute VB_Name = "modUrls"
Option Explicit

Public Function url_pys() As String

    url_pys = "http://" & gCon.backend_server & ":" & gCon.backend_port & "/"

End Function

Public Function join(paths As Variant) As String
Dim value As Variant

Dim url As String

    url = ""
    For Each value In paths
        url = url & "/" & value
    Next
    
    join = url

End Function
