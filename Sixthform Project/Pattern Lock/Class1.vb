Imports System.ComponentModel
Imports System.IO

Public Class Class1
    ' within this class there will be my encryption, decryption and qucik sort code
    Public ASCIIArray() As String
    Public IgnoreEncryptError As Boolean = False
    Function ConvertToASCII(ByVal SentenceToSort As String)
        ' here the program will convert it into ascii
        Dim stSentence = SentenceToSort
        Dim st_Split_Sentence = stSentence.Split(" ") ' This split the sentence into new array variables
        Dim icount As Integer
        Dim isentence As String = ""
        Dim Arraysize = 0
        Dim LargestLength As Integer = 0

        For Each WordToConvert In st_Split_Sentence ' this for loop is for every word needed to be converted
            isentence = ""
            Dim LengthOfWord = Len(WordToConvert)
            Dim LetterInASCII As String
            icount = 0

            Do Until icount = LengthOfWord
                'values of only 2 ascii numbers for each so that it Is easier for me to enycrypt
                LetterInASCII = Asc(WordToConvert.Substring(icount, 1)) 'converts it to ascii
                isentence = isentence + Str(LetterInASCII)  ' adds to the sentence
                icount = icount + 1
            Loop

            isentence = isentence.Replace(" ", "")          ' take out the spaces
            Arraysize += 1
            ReDim Preserve ASCIIArray(Arraysize - 1)
            If isentence.Length > LargestLength Then        ' checks the longest length
                LargestLength = isentence.Length            ' makes the largest length the larger of the two
            End If
            ASCIIArray(Arraysize - 1) = isentence           'Saves the word
        Next

        icount = 0
        isentence = ""
        For Each Word In ASCIIArray 'add the 0's in place to make all words the same size
            Dim CurrentWord = Word

            If Word.Length < LargestLength Then
                Do Until CurrentWord.Length = LargestLength
                    CurrentWord = CurrentWord + "0"
                Loop
                ASCIIArray(icount) = CurrentWord
            End If
            icount += 1
        Next
        '''''
        'For Each entry In ASCIIArray
        '    MsgBox(entry & vbNewLine & Len(entry))
        'Next
        Dim intArray() = Array.ConvertAll(ASCIIArray, Function(str) Long.Parse(str)) 'makes it into an integer array
        sorter(intArray, 0, intArray.Length - 1)                                     ' calls the quick sort
        'Converts the ascii into string
        'Combine each character of every word 
        'use this for ecnrytpion
        Dim SizeOfArray As Integer = 0
        Dim SortedWordArray() As String
        For Each asciiValue In intArray
            Dim CurrentValue As String = asciiValue
            For i = 0 To Len(CurrentValue) - 2 Step 2   'makes is so that every two numbers is an ascii value
                Dim CurrentCharacterCode = CurrentValue.Substring(i, 2)
                If CurrentCharacterCode = "00" Then     ' deletes all 0's in the word
                    Exit For
                End If
                ReDim Preserve SortedWordArray(SizeOfArray + 1)
                SortedWordArray(SizeOfArray) = SortedWordArray(SizeOfArray) + (Chr(CurrentCharacterCode)) ' Saves
            Next
            SizeOfArray += 1
        Next
        Return Strings.Join(Array.ConvertAll(SortedWordArray, Function(str) CStr(str)), " ") ' Converts the array into a single string that is returned 
    End Function
    ' here is the quick sort:
    Public Sub sorter(ByRef ASCIIArray1() As Long, lft As Long, rt As Long)
        If lft < rt Then
            Dim pos As Long = altering(ASCIIArray1, lft, rt)
            sorter(ASCIIArray1, lft, pos - 1)
            sorter(ASCIIArray1, pos + 1, rt)
        End If
    End Sub

    Function altering(ByRef ASCIIArray As Long(), lfter As Long, rter As Long)
        Dim piv As Long
        piv = ASCIIArray(lfter)

        Do While lfter <> rter
            Do While (ASCIIArray(rter) > piv) And (lfter <> rter)
                rter = rter - 1
            Loop
            ASCIIArray(lfter) = ASCIIArray(rter)
            Do While (ASCIIArray(lfter) < piv) And (lfter <> rter)
                lfter = lfter + 1
            Loop
            ASCIIArray(rter) = ASCIIArray(lfter)
            If ASCIIArray(lfter) = piv And ASCIIArray(rter) = piv Then
                Exit Do
            End If
        Loop
        ASCIIArray(lfter) = piv
        Return lfter
    End Function
    'Variables used by cypher
    Dim Character As String ' string character
    Dim KeyChar As String ' string character of key
    Dim Encrypted As String ' encrypted result of the addition or subtraction of ascii values
    Dim CharCount As Integer ' Length of sentence (Uses the i as the index of the character in a word in the for loop)
    Dim KeyLen As Integer ' length of key
    Dim KeyCount As Integer ' index character of the key
    Dim OverCount ' how far out of range the character changes to
    Dim EncryptedLine ' the final encryption of the line
    Dim IndexOfPass = 0
    Dim IndexOfLine = 0

    Public Function Decrpyt_data()
        'decrypt
        Dim FileToRead = GetUserFileLocation()   ' finds the user file
        If IgnoreEncryptError = False Then
            If File.ReadAllLines(FileToRead)(2).ToUpper = "DECRYPTED" Then ' If a file is decrypted before it was decrypted warn the user that the file was left decrypted 
                MsgBox("The program detected that the file was decrypted when it should have not, please take care closing the program", MsgBoxStyle.Exclamation, "Info")
                Return Nothing
                Exit Function
            End If
        End If
        Dim OriginalPassSenetence() = Decrpyt_Sentence.ToString.Split(" ")
        Dim LenthOfArrayPass = Decrpyt_Sentence.ToString.Split(" ").Count
        'Reverses the order of the words for decryption purposes, e.g. Hello there Bob >>> Bob there Hello
        Dim PassSentence(LenthOfArrayPass - 1)
        For i = 0 To LenthOfArrayPass - 1
            PassSentence((LenthOfArrayPass - 1) - i) = OriginalPassSenetence(i)
        Next
        IndexOfPass = 0
        IndexOfLine = 0
        Dim TempEncryptedLine 'stores encrypted line to then be re encrypted
        For Each TextLine In File.ReadAllLines(FileToRead)
            If IndexOfLine < 2 Then ' Skips First 2 Lines
                GoTo SkipLine ' if the index is under two it will skip the lines 
            End If
            TempEncryptedLine = TextLine
            For Each PassWordInSentence In PassSentence
                EncryptedLine = ""
                KeyCount = 0
                CharCount = Len(TempEncryptedLine) ' Length of current word 
                KeyLen = Len(PassWordInSentence) ' Length of current keyword
                Try
                    For i = 0 To CharCount - 1 ' Do for every character of the current word 
                        If KeyCount = KeyLen Then ' If the keyword already cycled through its length, reset it so that it loops through it again
                            KeyCount = 0
                        End If

                        Character = TempEncryptedLine.Substring(i, 1) ' Gets the currently selected character of the word 
                        If Asc(Character) = 32 Then ' If the character is a space then put a space in 
                            EncryptedLine &= " "
                        Else
                            KeyChar = PassWordInSentence.Substring(KeyCount, 1) ' Gets the currently selected character of the pass 
                            If AscW(KeyChar) = 32 Then ' Skip if it is a space 
                            Else

                                Encrypted = AscW(Character) - AscW(KeyChar) ' subtracts the location of the keyword and pass
                                Encrypted = Encrypted + 33 ' Fix bugs
                                If Encrypted < 33 Then ' If the result is lower than the lowest value, go from the top of the largest value
                                    OverCount = 33 - Encrypted ' See by the amount the value is under
                                    Encrypted = 126 - OverCount 'take away the amount it was under by
                                End If
                            End If
                            EncryptedLine &= Chr(Encrypted) ' Conert the ascii value back into a character
                            KeyCount = KeyCount + 1
                        End If
                    Next
                Catch ex As Exception
                End Try
                If EncryptedLine = "" Then
                Else
                    TempEncryptedLine = EncryptedLine 'changes the value of the temp so that it is the next word
                End If
            Next
            'Re write the line into file decrypted  
            Dim TextInLine() As String = File.ReadAllLines(FileToRead)
            TextInLine(IndexOfLine) = TempEncryptedLine
            File.WriteAllLines(FileToRead, TextInLine)
SkipLine:
            IndexOfLine += 1
        Next
        Return Nothing
    End Function

    Public Function Encrypt_data()
        'Encrypt
        'Notes are the same as the decrypt one
        Dim FileToRead = GetUserFileLocation()   ' finds the user file
        If IgnoreEncryptError = False Then
            If File.ReadAllLines(FileToRead)(2).ToUpper = "DECRYPTED" Then
            Else
                MsgBox("A Possible error occured because the file is already encrypted" & vbNewLine & "Please press OK to attempt to decrypt the message", MsgBoxStyle.Critical, "Error")
                Decrpyt_data()
                If File.ReadAllLines(FileToRead)(2).ToUpper = "DECRYPTED" Then
                    MsgBox("Attempt Was Successful, Contiuning to encrypt the file")
                Else
                    MsgBox("FAILED attempt" & vbNewLine & "An unexpected possible loss of data occured")
                End If
            End If
        End If
        Dim PassSentence() = Decrpyt_Sentence().ToString.Split(" ")     ' finds the sentence or paragraph from the users file
        IndexOfPass = 0
        IndexOfLine = 0
        Dim TempEncryptedLine
        For Each TextLine In File.ReadAllLines(FileToRead)
            If IndexOfLine < 2 Then ' Skips First 2 Lines
                GoTo SkipLine
            End If
            TempEncryptedLine = TextLine
            For Each PassWordInSentence In PassSentence
                EncryptedLine = ""
                KeyCount = 0
                CharCount = Len(TempEncryptedLine)
                KeyLen = Len(PassWordInSentence)
                Try
                    For i = 0 To CharCount - 1
                        If KeyCount = KeyLen Then
                            KeyCount = 0
                        End If
                        Character = TempEncryptedLine.Substring(i, 1)
                        If Asc(Character) = 32 Then
                            EncryptedLine &= " "
                        Else
                            KeyChar = PassWordInSentence.Substring(KeyCount, 1)
                            If AscW(KeyChar) = 32 Then
                            Else
                                Encrypted = Asc(Character) + Asc(KeyChar) ' change to +
                                Encrypted = Encrypted - 33 ' change to - 
                                If Encrypted > 126 Then
                                    OverCount = Encrypted - 126 ' takes away highest value
                                    Encrypted = 33 + OverCount ' change to +
                                End If
                            End If
                            EncryptedLine &= Chr(Encrypted)
                            KeyCount = KeyCount + 1
                        End If
                    Next
                Catch ex As Exception
                End Try
                If EncryptedLine = "" Then
                Else
                    TempEncryptedLine = EncryptedLine
                End If
            Next
            Dim TextInLine() As String = File.ReadAllLines(FileToRead)
            TextInLine(IndexOfLine) = TempEncryptedLine
            File.WriteAllLines(FileToRead, TextInLine)
SkipLine:
            IndexOfLine += 1
        Next
        Return Nothing
    End Function

    Private Function Decrpyt_Sentence()
        Dim sDecrypted_Sentence = ""
        Dim FileToRead = GetUserFileLocation() ' finds the user file
        Dim Key As String
        Dim PassSentence() = File.ReadAllLines(FileToRead)(1).ToString.Split(" ")     ' finds the sentence or paragraph from the users file
        Key = "RandomSentence" 'No Spaces Due to Support
        Dim DecryptedLine = ""
        For Each CurrentWord In PassSentence
            EncryptedLine = ""
            KeyCount = 0
            CharCount = Len(CurrentWord)
            KeyLen = Len(Key)
            Try
                For i = 0 To CharCount - 1

                    Character = CurrentWord.ToString.Substring(i, 1)
                    If Asc(Character) = 32 Then
                        DecryptedLine &= " "
                    Else
                        KeyChar = Key.Substring(KeyCount, 1)
                        If AscW(KeyChar) = 32 Then
                        Else
                            Encrypted = AscW(Character) - AscW(KeyChar)
                            Encrypted = Encrypted + 33
                            If Encrypted < 33 Then
                                OverCount = 33 - Encrypted
                                Encrypted = 126 - OverCount
                            End If
                        End If
                        DecryptedLine &= Chr(Encrypted)
                        KeyCount = KeyCount + 1
                    End If
                Next
            Catch
            End Try
            sDecrypted_Sentence &= DecryptedLine
        Next
        Return sDecrypted_Sentence
    End Function

    Public Function Encrypt_sentence()
        Dim sEncrypted_Sentence = ""
        Dim FileToRead = GetUserFileLocation() ' finds the user file
        Dim Key As String
        Dim PassSentence() = File.ReadAllLines(FileToRead)(1).ToString.Split(" ")     ' finds the sentence or paragraph from the users file
        Key = "RandomSentence" 'No Spaces Due to Support
        For Each CurrentWord In PassSentence
            EncryptedLine = ""
            KeyCount = 0
            CharCount = Len(CurrentWord)
            KeyLen = Len(Key)
            Try
                For i = 0 To CharCount - 1 ' loop is done for the durration of the whole sentence

                    Character = CurrentWord.ToString.Substring(i, 1) ' gets specific character of the sentence
                    If Asc(Character) = 32 Then 'checks is the space and ignore if it is 
                        EncryptedLine &= " "
                    Else
                        KeyChar = Key.Substring(KeyCount, 1) 'gets character of the key
                        If AscW(KeyChar) = 32 Then ' ignore the spaces
                        Else
                            Encrypted = Asc(Character) + Asc(KeyChar) 'adds the ascii values of the key and the character
                            Encrypted = Encrypted - 33 ' takes away to fix the bugs
                            If Encrypted > 126 Then ' checks if it is over the limit and then adds to the smallest values
                                OverCount = Encrypted - 126
                                Encrypted = 33 + OverCount
                            End If
                        End If
                        EncryptedLine &= Chr(Encrypted)
                        KeyCount = KeyCount + 1
                    End If
                Next
            Catch
            End Try
            sEncrypted_Sentence &= EncryptedLine 'saves it into new variable
        Next
        Dim TextInLine() As String = File.ReadAllLines(FileToRead) 'saves it in a new variable
        TextInLine(1) = sEncrypted_Sentence
        File.WriteAllLines(FileToRead, TextInLine)
        Return Nothing
    End Function
End Class
