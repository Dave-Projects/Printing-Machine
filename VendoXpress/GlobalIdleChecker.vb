Imports System.IO
Imports System.Runtime.InteropServices

Public Class GlobalIdleChecker
    '<DllImport("user32.dll")>
    'Private Shared Function GetLastInputInfo(ByRef plii As LASTINPUTINFO) As Boolean
    'End Function

    'Private Structure LASTINPUTINFO
    '    Public cbSize As UInteger
    '    Public dwTime As UInteger
    'End Structure

    'Private Shared WithEvents idleTimer As New Timer()

    'Public Shared Sub StartIdleCheck()
    '    idleTimer.Interval = 60000
    '    idleTimer.Start()
    'End Sub

    'Private Shared Sub idleTimer_Tick(sender As Object, e As EventArgs) Handles idleTimer.Tick
    '    ' Check for idle time
    '    Dim lastInputInfo As New LASTINPUTINFO()
    '    lastInputInfo.cbSize = CUInt(Marshal.SizeOf(lastInputInfo))
    '    If GetLastInputInfo(lastInputInfo) Then
    '        Dim idleTimeMilliseconds As UInteger = Environment.TickCount - lastInputInfo.dwTime
    '        Dim idleTimeMinutes As Double = idleTimeMilliseconds / (1000 * 60)

    '        If idleTimeMinutes >= 7 Then ' Restart after 8 minutes Idle
    '            mainForm.Dispose()
    '            mainForm.Close()
    '            GC.Collect()
    '            DeleteFilesInFolder() 'Clear Folder
    '            mainForm.Show()
    '            mainForm.ShowFormInPanel(landingForm)
    '            idleTimeMinutes = 0
    '        End If
    '    End If
    'End Sub
    'Public Shared Sub DeleteFilesInFolder()
    '    Dim cacheDirectory As String = "C:\CacheFiles"
    '    Try
    '        Dim files As String() = Directory.GetFiles(cacheDirectory)

    '        For Each filePath As String In files
    '            File.Delete(filePath)
    '        Next
    '    Catch ex As Exception

    '    End Try
    'End Sub
End Class
