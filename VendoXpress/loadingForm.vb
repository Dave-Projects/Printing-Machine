Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class loadingForm

    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    '    Timer1.Stop()
    '    DataFetcher.ApplyAsposeLicense()
    '    Try

    '        formMain.Show()
    '    Catch ex As Exception
    '        'MessageBox.Show("Internal Error occured, Please restart the appllication")
    '        'MessageBox.Show(ex.Message)
    '        'Application.Exit()
    '    End Try
    'End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        DataFetcher.ApplyAsposeLicense()

        Dim maxAttempts As Integer = 5
        Dim currentAttempt As Integer = 1

        While currentAttempt <= maxAttempts
            Try
                If Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1 Then
                    For Each proc As Process In Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName)
                        If proc.Id <> Process.GetCurrentProcess().Id Then
                            proc.CloseMainWindow()
                            proc.WaitForExit(5000)
                            proc.Close()
                        End If
                    Next
                End If

                formMain.Show()
                Exit Sub
            Catch ex As Exception
                'Console.WriteLine($"Attempt {currentAttempt} failed: {ex.Message}")
                currentAttempt += 1
            End Try
            System.Threading.Thread.Sleep(1000)
        End While

        MessageBox.Show("Failed to start the application after multiple attempts. Please restart the application.")
        Application.Exit()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub loadingForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
