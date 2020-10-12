Imports System.IO
Public Class Login

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.RegesteringStage = False
        Form1.Show()
        Me.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Register.Show()
        Me.Close()
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If Directory.Exists("Users") Then 'Checks if the user folder exists if not then it creates one
            Else
                Directory.CreateDirectory("Users") 'Creates the file
            End If
        Catch ex As Exception 'just an error msg for the user if the directroy cant be made
            MsgBox("there was a problem with creating esstential Directories!! ", MsgBoxStyle.Critical)
            Environment.Exit(0)
        End Try
    End Sub
End Class