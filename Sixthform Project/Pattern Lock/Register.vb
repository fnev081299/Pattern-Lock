Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Public Class Register

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For Each cntrl As Control In Me.Controls 'Loops through every control in the form e.g. a label, textbox etc
            If TypeOf cntrl Is TextBox Then 'If the current control is a textbox 
                If cntrl.Text = "" Then ' If the current control is a textbox and is empty then
                    MsgBox("Error one of the textboxes is empty") ' inform the user one of the textboxes is empty
                    Exit Sub
                End If
            End If
        Next
        Dim instance As New Class1 'creates and instance of object that encrytption and decryption occurs
        Try
            'Get all the files from the directory to compare existing usernames 
            For Each file In Directory.GetFiles("Users")
                'Remove Directory syntax E.g. Users\Test1.txt To Test1
                file = file.ToString.Split("\")(1).Split(".")(0)
                'Checks if username already exisits by hashing the username and compararing an already hassed username
                If GetHash(TextBox4.Text.ToUpper).ToUpper = file.ToUpper Then ' Hash the current username and compare it to the file 
                    'that already exists if they do exist then prompt the user that the username has been taken 
                    MsgBox("Username Already Exists" & vbNewLine & "Maybe Consider starting with your Surname,", MsgBoxStyle.Exclamation)
                    Exit Sub
                End If
            Next
            Dim FixedSplitOfPassPara = SplitBy5(TextBox5.Text) 'Calls a funtion that Splits a word larger than 5 into pices of length 5
            FixedSplitOfPassPara = instance.ConvertToASCII(FixedSplitOfPassPara)  'Sorts the sentence 
            'Create a new Variable to store the location of the new User File
            Dim NewFile As String
            'Declare the contents of the Variable to address the new file location and upper cases the username to enable non-sesentive logging in
            NewFile = "Users\" & GetHash(TextBox4.Text.ToUpper) & ".txt"
            'Creates the new File
            'Write Useful information: Name, Age and Password
            Dim writer As New StreamWriter(NewFile)
            writer.WriteLine("")                    ' Top line will be the password so keep it blank
            writer.WriteLine(FixedSplitOfPassPara)  ' this will store the security message 
            writer.WriteLine("Decrypted")           ' This line will be used for verification
            writer.WriteLine(TextBox1.Text)         ' This line will be their name
            writer.WriteLine(TextBox2.Text)         ' This line will be their age
            writer.WriteLine(TextBox3.Text)         ' This line will be their email
            writer.WriteLine("##Information##")     ' Anything past this will be their passwords
            writer.Close()
            MsgBox("All sorted, Please create a path pass now", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("Error Creating User File" & vbNewLine & ex.Message.ToString, MsgBoxStyle.Critical)
            Exit Sub
        End Try
        Form1.RegesteringStage = True           ' Indicate to the form1 that we are regestering a new user
        CurrentUsername = TextBox4.Text.ToUpper ' Tells the form the username of the user 
        ''''
        instance.Encrypt_sentence()
        instance.IgnoreEncryptError = True ' Ignores error checking for unencrypted files
        instance.Encrypt_data()
        instance.IgnoreEncryptError = False
        Form1.Show()
        Me.Close()
        '''''
    End Sub
    ''
    Function SplitBy5(ByVal InpWord As String)
        Dim SplitByNumber = 5
        Dim EachWord() = Strings.Split(InpWord)     ' Split by Space into an array
        Dim CurrentIndex = 0                        ' Know what part of the index we are in for the final thing 
        Dim TempWord                                ' Used as a place holder for the SplitByNumber characters of a word
        Dim TempLength As Integer
        Dim TempIndex As Integer = 0                'what character it is in the word, a pointer
        Dim ProcessedWord As String = ""
        For Each Word In EachWord                   ' No single word can be larger than SplitByNumber characters
            If Len(Word) > SplitByNumber Then                   'If word is is larger than SplitByNumber
                TempLength = Word.Length            'Temp length is the length of the current word
                TempIndex = 0
                ProcessedWord = ""
                Do While True
                    If TempLength > SplitByNumber Then                  'If the length of the current word is greater than the threshold
                        TempWord = Word.Substring(TempIndex, SplitByNumber)         ' Uses a varablie that increments by SplitByNumber to know what part of the string to start at and then gets the next 8 characters 
                        ProcessedWord = ProcessedWord + TempWord + " "  ' Saves the 8 character word into another variable 
                    Else
                        ProcessedWord = ProcessedWord + Strings.Right(Word, TempLength) ' If the end of a word is less than 8 characters then it would have not been added so any left over characters are added at the end
                        Exit Do
                    End If
                    ' to calculate the length left and if it is larger than 5
                    TempIndex += SplitByNumber
                    TempLength -= SplitByNumber
                Loop
                EachWord(CurrentIndex) = ProcessedWord
            End If
            EachWord(CurrentIndex) = EachWord(CurrentIndex).ToUpper
            EachWord(CurrentIndex) = Strings.Replace(EachWord(CurrentIndex), "[^A-Z0-9]!@#$^*()_+=[\]\\,.?:'/%-", "") 'Remove any unsuported characters
            CurrentIndex += 1
        Next
        Return Strings.Join(EachWord, " ") ' Join all the words into one string
    End Function
    ''
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        'makes it so that the user only places numbers in their age textbox
        If Asc(e.KeyChar) < 8 Or Asc(e.KeyChar) > 8 And Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
            e.Handled = True
            MsgBox("please enter an age in numbers.")
        End If
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        ' makes it so that the user can only enter letters for their name
        If Asc(e.KeyChar) < 8 Or Asc(e.KeyChar) > 8 And Asc(e.KeyChar) < 32 Or Asc(e.KeyChar) > 32 And Asc(e.KeyChar) < 65 Or Asc(e.KeyChar) > 90 And Asc(e.KeyChar) < 97 Or Asc(e.KeyChar) > 122 Then
            e.Handled = True
            MsgBox("please enter your name please")
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Login.Show()
        Me.Close()
    End Sub
    Private Sub Register_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class