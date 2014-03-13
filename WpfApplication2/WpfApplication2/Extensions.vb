Imports System.Runtime.CompilerServices

Module Extensions
    <Extension()> _
    Public Function IsComponent(ByVal value As Integer) As Boolean
        Return value Mod 2 <> 0
    End Function
End Module
