
Imports System.IO
Public Class Form1
    Public stname As String
    Public NewArray() As String                                     'Stores the password being entered
    Public icount As Integer = 0
    Dim NeedToConfirmPassword As Boolean                            'This boolean is to indicate if the user needs to confirm the password
    'First time they enter the password => NeedToConfirmPassword = True => User needs to re-enter the password then
    'it can acutal be used, this is only needed when regestering the user
    Dim PasswordToConfirm As String                                 ' After they enter the password once it gets saved to when they re-enter the password to check if they are the same just to assure that they entered the right path on the grid of picture boxes
    Public RegesteringStage As Boolean = False                      'If this is true then the user is creating a new accountng
    Dim Image_Filled As Image = My.Resources.Circle_Filled          ' Gets the image of the clicked box and saves it to a variable
    Dim Image_Empty As Image = My.Resources.Circle_Empty            ' Gets the image of the empty box and saves it to a variable
    Public itry As Integer
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try                                                         ' Try will come handy if they haven't entered the password yet
            For Each Entry In NewArray                              ' Cycles through the array of currently entered password
                MsgBox(Entry)
            Next
        Catch
        End Try
    End Sub

    Sub HandleClick(sender As Object, e As EventArgs)
        Dim Currentobject As PictureBox = CType(sender, PictureBox) ' the sender is the object which trigered this form
        'in this is the Event handler of a picture box so we convert the sender to a picturebox so we can Control the picturebox which called this sub
        If Bitmap.Equals(Currentobject.Image, Image_Filled) Then
            ' Checks if the current picturebox is already checked (Already been Clicked)
            MsgBox("That position has been selected already", , "")
            Exit Sub
        End If
        Currentobject.Image = Image_Filled                          ' Fills the picturebox
        Me.Refresh()                                                ' Refreshed it 
        ReDim Preserve NewArray(icount) ' Changes the size of the array by one because icount counts the amount of times pictureboxes have been added to the array while keeping its contents not changed
        NewArray(icount) = Currentobject.Name.ToString.Replace("PictureBox", "") 'Saves the number of the picturebox by getting rid of the PictureBox syntax to the array which stores currently stored password sequence
        icount += 1
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each A As PictureBox In Me.Controls.OfType(Of PictureBox)()
            AddHandler A.Click, AddressOf HandleClick                ' Every picturebox is told that if they are clicked, go to the HandleClick sub
            'This is done by having each pictureboxes click event to go to the HandleClick Sub
        Next
        TextBox1.Text = CurrentUsername                              ' If we are regestering a user then we show/hide appropriate controls. the same happens if the user is trying to log in 
        If RegesteringStage = True Then
            Button1.Visible = True
            Button2.Visible = True
            Button3.Visible = False
            Button4.Visible = True
            TextBox1.ReadOnly = True
            Label1.Visible = False
        Else
            TextBox1.ReadOnly = False
            Label1.Visible = True
            Button1.Visible = False
            Button2.Visible = False
            Button3.Visible = True
            Button4.Visible = True
        End If
    End Sub

    Private Sub ResetEverything()
        For Each A As PictureBox In Me.Controls.OfType(Of PictureBox)()
            A.Image = Image_Empty
        Next                                                        ' Goes through every picturebox and sets it's image back to the unclicked version
        Try
            If NewArray.Count < 1 Then
                MsgBox("Cannot clear the path when it is already cleared")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Cannot clear the path when it is already cleared")
            Exit Sub
        End Try
        Array.Clear(NewArray, 0, NewArray.Length)                   ' Clears the array
        ReDim NewArray(0)                                           ' Resets the array size back to nothing
        icount = 0
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        'Convert array to a single string
        Dim PasswordStringArray As String = ""
        Try
            If NewArray.Count < 1 Then
                MsgBox("Please enter the password for this account")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Please enter the password for this account")
            Exit Sub
        End Try
        For Each entry In NewArray
            PasswordStringArray = PasswordStringArray & entry.ToString 'Cycles through the array and adds its contents to a single string so we can analyse it and hash it
        Next
        If NeedToConfirmPassword = True Then ' If the users is confirming the password do this, if not then do the else
            If PasswordStringArray = PasswordToConfirm Then ' Confirms that the password they just entered matches the password they entered the first time
                Try
                    Dim LocationOfUserFile = GetUserFileLocation() ' Sets the location of the userfile to a variable
                    Dim LinesinFile() As String = File.ReadAllLines(LocationOfUserFile) ' Reads all the lines of the file Sets the End Result Text of the line selected
                    LinesinFile(0) = GetHash(PasswordStringArray) ' Selects line 1 which is represented as the 0 Index of the array and we set its contnets, but it doesnt write to the file yet
                    File.WriteAllLines(LocationOfUserFile, LinesinFile) 'Replaces the selected line with a the variable by writing to the file
                    MsgBox("All done!! You can now log in!!")
                    Login.Show()
                    Me.Close()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Else
                MsgBox("Password did not match please re-try", MsgBoxStyle.Information, "Regiester")
                'If the password they just entered does not match the first time inform the user

                NeedToConfirmPassword = False 'Indicates that the next time they press this button it will be their
                'first time entering the password
                ResetEverything() 'Reset everything
            End If
        Else
            PasswordToConfirm = PasswordStringArray ' Saves the current password so it can be compared later
            NeedToConfirmPassword = True 'Tells the program it will need to check this password with the next password the user inputs
            ResetEverything() ' Clear Everyting
            MsgBox("Please Confirm your password")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        stname = TextBox1.Text
        If TextBox1.Text = "" Then
            MsgBox("Please enter your username")
            Exit Sub
        End If
        Try
            If NewArray.Count < 1 Then
                MsgBox("Please enter the password for this account")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Please enter the password for this account")
            Exit Sub
        End Try
        Try
            Dim PasswordStringArray As String = ""
            'Convert array to a single string
            For Each entry In NewArray
                PasswordStringArray = PasswordStringArray & entry.ToString 'Cycles through the array and adds its contents to a single
                'string so we can analyse it and hash it
            Next
            Dim CheckPassword = File.ReadAllLines("Users\" & GetHash(TextBox1.Text.ToUpper) & ".txt")(0) 'Reads the first line**
            ' This is done by reading all lines and only saving the first line (Index 0) this is done by putting a (0) which
            'indicates that we just want the first index of the array
            If CheckPassword = GetHash(PasswordStringArray) Then ' Does the hash of the password entered match 
                'the hash that is saved in the file
                MsgBox("Password Correct")
                CurrentUsername = TextBox1.Text.ToUpper
                connected.Show()
                Me.Close()
                Exit Sub
            Else
                MsgBox("Password InCorrect")
                itry += 1
                If itry > 2 Then
                    MsgBox("You have entered your password wrong more than three times!!")
                    Me.Close()
                    Exit Sub
                End If
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Please enter a Username and pattern")
            Exit Sub
        End Try
        MsgBox("Password Correct")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        NeedToConfirmPassword = False ' Allows the user to enter the password for the first time, (Not needing to confirm it this time)
        ResetEverything()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Login.Show()
        Me.Close()
    End Sub
End Class
