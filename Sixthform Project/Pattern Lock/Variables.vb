Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Module Variables
    Public CurrentUsername As String
    Public EncryptDecryptClass As New Class1
    Function GetHash(theInput As String) As String
        Using hasher As MD5 = MD5.Create()    ' create hash object
            ' Convert input to byte array and get hash
            Dim dbytes As Byte() =
                 hasher.ComputeHash(Encoding.UTF8.GetBytes(theInput))
            ' declare a stringbuilder to convert bytes to string
            Dim sBuilder As New StringBuilder()
            ' convert byte data to hex string
            For n As Integer = 0 To dbytes.Length - 1 'Loops through every entry of the hash byte and converts it to a string
                sBuilder.Append(dbytes(n).ToString("X2")) ' "X" is an argument for converting the string to
                'hexadecimal, the 2 singals that if the byte is empty it will replace it with 2 zeros
            Next n
            Return sBuilder.ToString() 'outputs the hash as a hexadeciamal string
        End Using
    End Function
    Function GetUserFileLocation()
        Dim FileToRead
        Try
            FileToRead = "Users\" & GetHash(CurrentUsername) & ".txt" ' finds the user file
        Catch ex As Exception
            Return "FAIL"
        End Try
        Return FileToRead
    End Function

End Module
