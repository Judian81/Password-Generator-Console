Imports System.Windows.Forms

Module Module1
    Public sAppPath As String
    Public sDebugMode As Boolean

    'i you give a path name without a file it takes care it ends with a "\"
    'Function DirReal(sZ As String) As Boolean
    '    Dim MyName As String, cPath As String, vA As Integer, vToMake As String
    '    'the path needs to be closed as a "\"
    '    cPath = PathFix(sZ)
    '    'now we remove the "\", in this order it will be right when you close or not close with a "\"
    '    cPath = Left(cPath, Len(cPath) - 1)
    '    vToMake = cPath
    '    For vA = Len(cPath) - 1 To 1 Step -1
    '        If Mid(cPath, vA, 1) = "\" Then
    '            cPath = Left(vToMake, vA)
    '            vToMake = Right(vToMake, Len(vToMake) - vA)
    '            Exit For
    '        End If
    '    Next
    '    MyName = Dir(cPath, vbDirectory)   ' Retrieve the first entry.
    '    Do While MyName <> ""   ' Start the loop.
    '        ' Use bitwise comparison to make sure MyName is a directory.
    '        If (GetAttr(cPath & MyName) And vbDirectory) = vbDirectory Then
    '            ' Display entry only if it's a directory.
    '            If MyName = vToMake Then
    '                Return True
    '            End If
    '        End If
    '        MyName = Dir()   ' Get next entry.
    '    Loop
    '    Return False
    'End Function

    'in some functions you have to have a good path that closes with "\" so you can do "C:\maps\folder" and it ends up like this "c:\maps\folder\"
    'Function PathFix(sPath As String) As String
    '    Dim sPathTemp As String
    '    'is there anything??
    '    If sPath <> "" Then
    '        'this is strange but it should do any good i quess
    '        If Left(sPath, 2) = "//" Then
    '            'if the wrong pahtlike stuff is happing correct it
    '            If Right(sPath, 1) <> "/" Then
    '                sPathTemp = sPath & "\"
    '            Else
    '                sPathTemp = sPath
    '            End If
    '            'see if the end is correct. if there is no "\" make one
    '        ElseIf Right(sPath, 1) <> "\" Then
    '            sPathTemp = sPath & "\"
    '        Else
    '            'it looks like it is good the way it was
    '            sPathTemp = sPath
    '        End If
    '    Else
    '        sPathTemp = ""
    '    End If
    '    Return sPathTemp
    'End Function

    ''this function makes a new directory if it does not exist
    'Function DirMake(sPath As String, Optional bInfo As Boolean = False) As String
    '    'is this path is not already there we could go on and created it
    '    If DirReal(sPath) = False Then
    '        Call MkDir(sPath)
    '        Return "A directory is created"
    '    Else
    '        Return "Directory is already there"
    '    End If
    'End Function

    Sub SHelp()
        'write text to the console. this is the help part so we could no what we can or can not
        Console.WriteLine(vbCrLf & "Uitleg:")
        Console.WriteLine("-l 48  het totale lengte van het wachtwoord.")
        Console.WriteLine("")
        Console.WriteLine("-n *   cijvers '012', * vul een cijver in bijvoorbeeld als -n 3 om het niet meer dan 3 keer te laten voorkomen")
        Console.WriteLine("-m *   het min teken '-', vul je geen cijver in op * dan kan het min teken onbeperkt gebruikt worden")
        Console.WriteLine("-s *   speciale karakters '!@#'")
        Console.WriteLine("-uc *  hoofdletters 'ABC'")
        Console.WriteLine("-lc *  kleine letters 'abc'")
        Console.WriteLine("-u *   het lage streepje '_'")
        Console.WriteLine("-b *   tussen haakjes karakters ']{}<'")
        Console.WriteLine("-sb *  spatie ' '")
        Console.WriteLine("-ha *  high anscii '€¾ƒ'")
        Console.WriteLine("-de    show how password is made")
        Console.WriteLine("")
        'Console.WriteLine("-o     output naar 'clipboard' / 'file'")
        Console.WriteLine("")
        Console.WriteLine("-shell start up in shell mode")
        Console.WriteLine("")
        Console.WriteLine("-exit  to quit the shell")
        Console.WriteLine("-quit  to quit the shell")
        Console.WriteLine("")
        Console.WriteLine("-help2 for more information")
        Console.WriteLine("example: -l 28 -n 3 -s -uc -lc 13 -m 4 -u 1 -b 2") ' -o file")
    End Sub

    Function STotalNumbers(vTemp As String) As Integer
        If Trim(vTemp) = "" Then
            Return 0
        Else
            Return Val(vTemp)
        End If
    End Function

    Sub SChar(aTemp() As String)
        Dim sCompaire(8) As String, sCompaireTotal(8) As Integer, cLen As Integer
        Dim vA As Integer, vB As Integer, value As String, sSaveOption As String, s127 As String
        'set this option so it does not have to be one that must be given
        sSaveOption = ""
        value = ""
        'use this so the password is not the same
        Randomize()
        'here we sort things so we know what to use in the password. and also how many times
        For vA = 1 To aTemp.Length - 1
            Select Case Left(aTemp(vA), 2)
                Case "l "
                    'total length of the password
                    cLen = STotalNumbers(Mid(aTemp(vA), 3))
                Case "n "
                    'numbers are included in the password
                    sCompaireTotal(0) = STotalNumbers(Mid(aTemp(vA), 3))
                    If sCompaireTotal(0) = 0 Then
                        sCompaireTotal(0) = -1
                    End If
                    '10 numbers, max 127 chars so 127 / 10 = 12 * 10 + 7
                    value = "0123456789"
                    s127 = ""
                    For vB = 1 To 12
                        s127 &= value
                    Next vB
                    For vB = 1 To 7
                        s127 &= Mid(value, CInt(Int((Len(value) * Rnd()) + 1)), 1)
                    Next vB
                    sCompaire(0) &= s127
                Case "s "
                    'we want to use special chars in the password
                    sCompaireTotal(1) = STotalNumbers(Mid(aTemp(vA), 3))
                    If sCompaireTotal(1) = 0 Then
                        sCompaireTotal(1) = -1
                    End If
                    '10 special chars, max 127 chars so 127 / 10 = 12 * 10 + 7
                    value = "!@#$%^&*=+"
                    s127 = ""
                    For vB = 1 To 12
                        s127 &= value
                    Next vB
                    For vB = 1 To 7
                        s127 &= Mid(value, CInt(Int((Len(value) * Rnd()) + 1)), 1)
                    Next vB
                    sCompaire(1) &= s127
                Case "uc"
                    'use upper case karakters in the pass string
                    sCompaireTotal(2) = STotalNumbers(Mid(aTemp(vA), 4))
                    If sCompaireTotal(2) = 0 Then
                        sCompaireTotal(2) = -1
                    End If
                    '26 upper letters, max 127 chars so 127 / 26 = 4 * 26 + 23
                    value = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    s127 = ""
                    For vB = 1 To 4
                        s127 &= value
                    Next vB
                    For vB = 1 To 23
                        s127 &= Mid(value, CInt(Int((Len(value) * Rnd()) + 1)), 1)
                    Next vB
                    sCompaire(2) &= s127
                Case "lc"
                    'use lower case karakters in password
                    sCompaireTotal(3) = STotalNumbers(Mid(aTemp(vA), 4))
                    If sCompaireTotal(3) = 0 Then
                        sCompaireTotal(3) = -1
                    End If
                    '26 letters, max 127 chars so 127 / 26 = 4 * 26 + 23
                    value = "abcdefghijklmnopqrstuvwxyz"
                    s127 = ""
                    For vB = 1 To 4
                        s127 &= value
                    Next vB
                    For vB = 1 To 23
                        s127 &= Mid(value, CInt(Int((Len(value) * Rnd()) + 1)), 1)
                    Next vB
                    sCompaire(3) &= s127
                Case "m "
                    'use the minus karakter in string
                    sCompaireTotal(4) = STotalNumbers(Mid(aTemp(vA), 3))
                    If sCompaireTotal(4) = 0 Then
                        sCompaireTotal(4) = -1
                    End If
                    '1 -, max 127 chars so 127 / 1 = 127
                    value = "-"
                    s127 = ""
                    For vB = 1 To 127
                        s127 &= value
                    Next vB
                    sCompaire(4) &= s127
                Case "u "
                    'use the underscore karakter in the password
                    sCompaireTotal(5) = STotalNumbers(Mid(aTemp(vA), 3))
                    If sCompaireTotal(5) = 0 Then
                        sCompaireTotal(5) = -1
                    End If
                    '1 _, max 127 chars so 127 / 1 = 127
                    value = "_"
                    s127 = ""
                    For vB = 1 To 127
                        s127 &= value
                    Next vB
                    sCompaire(5) &= s127
                Case "b "
                    'use brakets in the password
                    sCompaireTotal(6) = STotalNumbers(Mid(aTemp(vA), 3))
                    If sCompaireTotal(6) = 0 Then
                        sCompaireTotal(6) = -1
                    End If
                    '8 []{}<>(), max 127 chars so 127 / 8 = 15 * 8 + 7
                    value = "[]{}<>()"
                    s127 = ""
                    For vB = 1 To 15
                        s127 &= value
                    Next vB
                    For vB = 1 To 7
                        s127 &= Mid(value, CInt(Int((Len(value) * Rnd()) + 1)), 1)
                    Next vB
                    sCompaire(6) &= s127
                Case "sb"
                    'include the space bar
                    sCompaireTotal(7) = STotalNumbers(Mid(aTemp(vA), 4))
                    If sCompaireTotal(7) = 0 Then
                        sCompaireTotal(7) = -1
                    End If
                    '1 spacebar, max 127 chars so 127 / 1 = 127
                    value = " "
                    s127 = ""
                    For vB = 1 To 127
                        s127 &= value
                    Next vB
                    sCompaire(7) &= s127
                Case "ha"
                    'include the high anscii
                    sCompaireTotal(8) = STotalNumbers(Mid(aTemp(vA), 4))
                    If sCompaireTotal(8) = 0 Then
                        sCompaireTotal(8) = -1
                    End If
                    For vB = 128 To 255
                        sCompaire(8) &= Chr(vB)
                    Next vB
                Case "de"
                    sDebugMode = True
                Case "o "
                    'this is to controle the option to save the password
                    sSaveOption = Mid(aTemp(vA), 3)
            End Select
        Next
        'make the string empty to prepare for password
        value = ""
        For vB = 0 To 8
            For vC = 1 To sCompaireTotal(vB)
                'put the karakters that we want in the password
                value &= Mid(sCompaire(vB), CInt(Int((Len(sCompaire(vB)) * Rnd()) + 1)), 1)
            Next
            'if the karkaters could be there without a limit -1 is unlimeted times that are possible in the password
            If sCompaireTotal(vB) <> -1 Then
                'the total amount of the karakters are used in the password for this type
                sCompaireTotal(vB) = 0
            End If
        Next
        'add up chars that have a max numbers. we look later on it the text is to long or need some other stuff to fill the password lenght
        If sDebugMode = True Then Console.WriteLine(vbCrLf & "Limited Chr: '" & value & "'")
        'make string empty. we want to use this to put in the chars we wanted in the password
        Dim sGen As String
        sGen = ""
        'check if the password we created is long enough or we have to add more chars to the password
        If Len(value) < cLen Then
            For vB = 0 To 8
                'one of this char is what we want to be in the password. but whitout a limit
                If sCompaireTotal(vB) = -1 Then
                    'the chars are added to the string. so we could create the rest of the password. we add as exaple '0123456789' and '-' then randomize picks one char from '0123456789-'
                    sGen &= sCompaire(vB)
                End If
            Next
            'this is the part we add to the password if there is room for more chars
            For aC = Len(value) To cLen - 1
                'we randomize what chars we want to have from this string
                value &= Mid(sGen, CInt(Int((Len(sGen) * Rnd()) + 1)), 1)
            Next
            If sDebugMode = True Then Console.WriteLine("Chr left to be added: '" & value & "'")
        ElseIf Len(value) > cLen Then
            If sDebugMode = True Then Console.WriteLine("Password length wanted: " & cLen & " Password lenght possible: " & Len(value))
        End If

        'this are chars enabled but whitout a max total. we could use this if we want to fill the rest of the password
        'If sDebugMode = True Then Console.WriteLine("Chars left to be used like ( -n ): " & sGen)
        'save the password in another string
        sGen = value
        Dim sTemp As String, iTemp As Integer
        'here we huzzel the password. so it is not always for example starting with numbers
        sTemp = ""
        For vA = 0 To Len(value)
            'this is randomized. and chouse a char to be cut and paste in the other string. that will be ad the end the password
            iTemp = CInt(Int((Len(sGen) * Rnd()) + 1))
            sTemp &= Mid(sGen, iTemp, 1)
            sGen = Left(sGen, iTemp - 1) & Mid(sGen, iTemp + 1)
        Next
        'the password is huzzeld. if the password is longer that we want. we just cut a part so it is the right size
        sTemp = Left(sTemp, cLen)
        Console.WriteLine("Password huzzeld: '" & sTemp & "'")
        'this setting is here so we could controle were we want the password be copied to
        'If Trim(sSaveOption) = "file" Then
        'FileOpen(1, sAppPath & "PassGen.txt", OpenMode.Output)
        'write text info to file. this could restore the file when it is deleted
        ' PrintLine(1, sTemp)
        'Close the file.
        'FileClose(1)
        'Console.WriteLine("Pass saved to 'C:\ProgramData\Jur-Gen\PassGen.txt'")
        'ElseIf Trim(sSaveOption) = "clipboard" And sTemp <> "" Then
        '    Clipboard.SetText(sTemp)
        '    Console.WriteLine(vbCrLf & "Copied to clipboard")
        'Else
        '    Console.WriteLine(vbCrLf & "Not Copied to clipboard and not saved in file")
        'End If
        sDebugMode = False
    End Sub

    Sub Main(ByVal sArgs() As String)
        'the path must exist
        'sAppPath = "C:\ProgramData\" & "Password Generator\"
        'If DirReal(sAppPath) = False Then
        '    'create new path. for the aplication itself . the directory is \Password Generator
        '    DirMake(sAppPath)
        'End If
        If sArgs.Length = 0 Then
            'go one like you have too
            Call SHelp()
        ElseIf sArgs(0) = "-shell" Then
            Call SHelp()
            'go one like you have too
        ElseIf sArgs(0) = "-help" Then
            Call SHelp()
            Exit Sub
        ElseIf sArgs(0) = "-help2" Then
            Console.WriteLine("")
            Console.WriteLine("───▄▄▄▄▄▄─────▄▄▄▄▄▄")
            Console.WriteLine("─▄█▓▓▓▓▓▓█▄─▄█▓▓▓▓▓▓█▄")
            Console.WriteLine("▐█▓▓▒▒▒▒▒▓▓█▓▓▒▒▒▒▒▓▓█▌")
            Console.WriteLine("█▓▓▒▒░╔╗╔═╦═╦═╦═╗░▒▒▓▓█")
            Console.WriteLine("█▓▓▒▒░║╠╣╬╠╗║╔╣╩╣░▒▒▓▓█")
            Console.WriteLine("▐█▓▓▒▒╚═╩═╝╚═╝╚═╝▒▒▓▓█▌")
            Console.WriteLine("─▀█▓▓▒▒░░░░░░░░░▒▒▓▓█▀")
            Console.WriteLine("───▀█▓▓▒▒░░░░░▒▒▓▓█▀")
            Console.WriteLine("─────▀█▓▓▒▒░▒▒▓▓█▀")
            Console.WriteLine("──────▀█▓▓▒▓▓█▀")
            Console.WriteLine("────────▀█▓█▀")
            Console.WriteLine("──────────▀")
            Console.WriteLine("")
        ElseIf sArgs.Length > 1 Then
            Dim sTemp As String, vA As Integer
            sTemp = ""
            For vA = 0 To sArgs.Length - 1
                sTemp &= sArgs(vA) & " "
            Next
            SChar(Split(sTemp, "-"))
            Exit Sub
        End If

        'Read value.
        Dim sRunning As Boolean
        sRunning = True
        While sRunning = True
            Dim sConsole As String = Console.ReadLine()
            If sConsole <> "" Then
                If sConsole = "-help" Then
                    Call SHelp()
                ElseIf sConsole = "-help2" Then
                    Console.WriteLine("")
                    Console.WriteLine("───▄▄▄▄▄▄─────▄▄▄▄▄▄")
                    Console.WriteLine("─▄█▓▓▓▓▓▓█▄─▄█▓▓▓▓▓▓█▄")
                    Console.WriteLine("▐█▓▓▒▒▒▒▒▓▓█▓▓▒▒▒▒▒▓▓█▌")
                    Console.WriteLine("█▓▓▒▒░╔╗╔═╦═╦═╦═╗░▒▒▓▓█")
                    Console.WriteLine("█▓▓▒▒░║╠╣╬╠╗║╔╣╩╣░▒▒▓▓█")
                    Console.WriteLine("▐█▓▓▒▒╚═╩═╝╚═╝╚═╝▒▒▓▓█▌")
                    Console.WriteLine("─▀█▓▓▒▒░░░░░░░░░▒▒▓▓█▀")
                    Console.WriteLine("───▀█▓▓▒▒░░░░░▒▒▓▓█▀")
                    Console.WriteLine("─────▀█▓▓▒▒░▒▒▓▓█▀")
                    Console.WriteLine("──────▀█▓▓▒▓▓█▀")
                    Console.WriteLine("────────▀█▓█▀")
                    Console.WriteLine("──────────▀")
                    Console.WriteLine("")
                ElseIf sConsole = "-exit" Or sConsole = "-quit" Or sConsole = "exit" Or sConsole = "quit" Then
                    sRunning = False
                Else
                    If Split(sConsole, "-").Length > 0 Then
                        Call SChar(Split(sConsole, "-"))
                    End If
                End If
            End If
        End While
    End Sub
End Module
