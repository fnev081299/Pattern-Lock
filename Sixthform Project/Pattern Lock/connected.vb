Imports System.Security.Cryptography
Imports System.ComponentModel
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.Threading

Public Class connected
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'add in a username so that the file read is the current user
        Dim site As String
        Dim pass As String

        site = TextBox1.Text
        pass = TextBox2.Text

        If site = "" And pass = "" Then
            MsgBox("Please enter a website and please enter a password")
            Exit Sub
        ElseIf site = "" Then
            MsgBox("Please enter a website")
            TextBox2.Clear()
            Exit Sub
        ElseIf pass = "" Then
            MsgBox("Please enter a password")
            TextBox1.Clear()
            Exit Sub
        End If

        Dim LocationOfUserFile = GetUserFileLocation() ' Sets the location of the userfile to users
        Dim writer As New StreamWriter(LocationOfUserFile, True)
        writer.WriteLine(TextBox1.Text) 'writes a line for the textbox for the website
        writer.WriteLine(TextBox2.Text) 'writes a line for the textbox for the password
        writer.Close()                  'closes and flushes the data
        MsgBox("saved to your file")
    End Sub
    Function GetUserInfo()
        Dim PasswordStringArray As String = ""
        Dim LocationOfUserFile = GetUserFileLocation()   ' Sets the location of the userfile to users
        Dim LinesinFile() As String = File.ReadAllLines(LocationOfUserFile)     ' Reads all the lines of the file
        Return LinesinFile
    End Function
    Private Sub Formloggedin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Do While True
            Try
                EncryptDecryptClass.Decrpyt_data()
                Exit Sub
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        Loop
        'decrypt textfile using the class when opened
    End Sub

    Private Sub Formloggedin_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Do While True
            Try
                EncryptDecryptClass.Encrypt_data()
                Exit Sub
            Catch ex As Exception
            End Try
        Loop
        ' encrypts textfile using the class when closed
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'will find the password within the file for the certain website
        Dim str2() As String = GetUserInfo() 'content of text file
        Dim s As String 'line in the text file, the content of a line used in the for loop
        Dim website As String
        Dim bFound As Boolean = False
        Dim icount As Integer = 0
        website = TextBox3.Text.ToUpper
        If website = "" Then
            MsgBox("Please enter a website")
            TextBox1.Clear()

            Exit Sub
        End If
        For Each s In str2
            If s.ToUpper.Contains(website) Then
                bFound = True
                'read the line of where the website name is then take out the password below it
                If IEEERemainder(icount, 2) = -1 Then ' gets the remainder
                    If MsgBox("Is this the correcct Password" & vbNewLine & "Website: " & s & vbNewLine & "Password: " & (str2(icount + 1)), MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        TextBox4.Text = (str2(icount + 1)) 'gets the password
                        Exit For
                    End If
                End If
            End If
            icount += 1
        Next
        If bFound = False Then
            MsgBox("Can't find the password for this website")
        End If
        'here i will have a readline function that will read all lines and then find the current website to 
        'then show thw password on a text box
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'will generate a long password that will be strong enough for webistes to aprove of
        Dim Letters As New List(Of Integer)
        'add ASCII codes for numbers
        For i As Integer = 48 To 57
            Letters.Add(i)
        Next
        'lowercase letters
        For i As Integer = 97 To 122
            Letters.Add(i)
        Next
        'uppercase letters
        For i As Integer = 65 To 90
            Letters.Add(i)
        Next
        'selects random integers from a number of items in Letters
        'then convert those random integers to characters and
        'add each to a string and display in the Textbox
        Dim Rnd As New Random
        Dim SB As New System.Text.StringBuilder
        Dim Temp As Integer
        For count As Integer = 1 To 20
            Temp = Rnd.Next(0, Letters.Count)
            SB.Append(Chr(Letters(Temp)))
        Next
        TextBox5.Text = SB.ToString
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Login.Show()
        Me.Close()
    End Sub
End Class
