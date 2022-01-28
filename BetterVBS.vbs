Option Explicit
Dim version : version = "v1.0.8"
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim sound : Set sound = CreateObject("WMPlayer.OCX")
Dim nw : nw = vbCrlf
Dim tb : tb = vbTab
Dim qt : qt = """"
Dim nbsp : nbsp = " "
Dim NUL : NUL = empty

Dim debugMode : debugMode = true
Dim debugRunCommand
if debugMode = true then
    debugRunCommand = "$run -test.bvbs"
else
    debugRunCommand = ""
end if

'Sytem Variables
Dim lineNr : lineNr = 0

'Global Variables
Dim filePath
Dim checkExtension
Dim checkRegister
Dim checkOperation

'Registers
Dim eax, ebx, ecx, edx
Dim rax, rbx, rcx, rdx

Call subConsole()
Sub subConsole()
    Dim console : console = InputBox("","BetterVBS " & version,debugRunCommand)
    '$run
    if InStr(console, "$run") > 0 then
        filePath = Replace(console, "$run -","")
        if not InStr(filePath, ".bvbs") > 0 then
            msgBox "Make sure the file extension is .bvbs",0+16,"Unsupported file extension"
            WScript.Quit
        else
            Call compile()
        end if
    end if
End Sub

Sub compile()
    Dim prefix : prefix = "$"
    Dim file : Set file = fso.OpenTextFile(filePath)
    Dim constLine
    Do Until file.AtEndOfStream
        WScript.Sleep(100)
        constLine = file.ReadLine
        lineNr = lineNr + 1

        '$if
        if Left(constLine,3) = prefix & "if" then
            Dim ifCommand : ifCommand = Split(constLine,":")
            'condition section
            ifCommand(0) = Replace(ifCommand(0), prefix & "if","")
            ifCommand(0) = Replace(ifCommand(0), nbsp,"")
            ifCommand(0) = Replace(ifCommand(0), "(","")
            ifCommand(0) = Replace(ifCommand(0), ")","")
            Dim ifCond(2)
            ifCond(0) = Left(ifCommand(0),3)
            ifCommand(0) = Replace(ifCommand(0),ifCond(0),"")
            Dim ifOperator : ifOperator = Left(ifCommand(0),1)
            ifCommand(0) = Replace(ifCommand(0),ifOperator,"")
            ifCond(1) = Left(ifCommand(0),3)
            ifCommand(0) = NUL

            'Filter
            ifCond(0) = Replace(ifCond(0),nbsp,"")
            ifCond(1) = Replace(ifCond(1),nbsp,"")

            Dim ifRes : ifRes = false
            Dim ifLoop : ifLoop = 0
            Do
                if ifCond(ifLoop) = "eax" then
                    ifCond(ifLoop) = eax
                elseif ifCond(ifLoop) = "ebx" then
                    ifCond(ifLoop) = ebx
                elseif ifCond(ifLoop) = "ecx" then
                    ifCond(ifLoop) = ecx
                elseif ifCond(ifLoop) = "edx" then
                    ifCond(ifLoop) = edx
                elseif ifCond(ifLoop) = "rax" then
                    ifCond(ifLoop) = rax
                elseif ifCond(ifLoop) = "rbx" then
                    ifCond(ifLoop) = rbx
                elseif ifCond(ifLoop) = "rcx" then
                    ifCond(ifLoop) = rcx
                elseif ifCond(ifLoop) = "rdx" then
                    ifCond(ifLoop) = rdx
                end if
                ifLoop = ifLoop + 1
            Loop until ifLoop = 2

            'Convert to int
            ifCond(0) = CInt(ifCond(0))
            ifCond(1) = CInt(ifCond(1))

            if ifOperator = "=" then
                if ifCond(0) = ifCond(1) then
                    ifRes = true
                end if
            elseif ifOperator = "!" then
                if ifCond(0) <> ifCond(1) then
                    ifRes = true
                end if
            elseif ifOperator = "<" then
                if ifCond(0) < ifCond(1) then
                    ifRes = true
                end if
            elseif ifOperator = ">" then
                if ifCond(0) > ifCond(1) then
                    ifRes = true
                end if
            end if

            if ifRes = true then
                constLine = ifCommand(1)
            end if
        end if
        '//comment
        if Left(constLine,2) = "//" then
            'Line is a comment
        '$system
        elseif Left(constLine,7) = prefix & "system" then
            Dim sysCommand : sysCommand = Replace(constLine, prefix & "system","")
            sysCommand = Replace(sysCommand, ".","")
            if Right(sysCommand,2) = "x)" then
                sysCommand = Replace(sysCommand, "(","")
                sysCommand = Replace(sysCommand, ")","")
                Dim sysReg : sysReg = Right(sysCommand,3)
                sysCommand = Replace(sysCommand,sysReg,"")
            else
                sysCommand = Replace(sysCommand, "(","")
                sysCommand = Replace(sysCommand, ")","")
            end if
        '$fso
        elseif Left(constLine,4) = prefix & "fso" then
            Dim fsoCopy : fsoCopy = false
            Dim fsoCommand : fsoCommand = Replace(constLine, prefix & "fso","")
            fsoCommand = Replace(fsoCommand, "(","")
            fsoCommand = Replace(fsoCommand, ")","")
            fsoCommand = Replace(fsoCommand, qt, "")
            checkExtension = Left(fsoCommand,5)
            if checkExtension = ".crtx" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            elseif checkExtension = ".crfo" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            elseif checkExtension = ".cytx" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            elseif checkExtension = ".cyfo" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            elseif checkExtension = ".dltx" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            elseif checkExtension = ".dlfo" then
                fsoCommand = Replace(fsoCommand, checkExtension,"")
            end if
            if InStr(fsoCommand,",") then
                fsoCommand = Replace(fsoCommand,nbsp,"")
                Dim fsoSection : fsoSection = Split(fsoCommand,",")
                fsoCopy = true
            end if
            if fsoCommand = "eax" then
                fsoCommand = eax
            elseif fsoCommand = "ebx" then
                fsoCommand = ebx
            elseif fsoCommand = "ecx" then
                fsoCommand = ecx
            elseif fsoCommand = "edx" then
                fsoCommand = edx
            end if
            if checkExtension = ".crtx" then
                if fsoCopy = false then
                    fso.CreateTextFile(fsoCommand)
                else
                    'return error
                end if
            elseif checkExtension = ".crfo" then
                if fsoCopy = false then
                    fso.createFolder(fsoCommand)
                else
                    'return error
                end if
            elseif checkExtension = ".cytx" then
                if fsoCopy = true then
                    fso.CopyFile fsoSection(0),fsoSection(1)
                else
                    'return error
                end if
            elseif checkExtension = ".cyfo" then
                if fsoCopy = true then
                    fso.CopyFolder fsoSection(0),fsoSection(1)
                else
                    'return error
                end if
            elseif checkExtension = ".dltx" then
                if fsoCopy = false then
                    fso.DeleteFile(fsoCommand)
                else
                    'return error
                end if
            elseif checkExtension = ".dlfo" then
                if fsoCopy = false then
                    fso.DeleteFolder(fsoCommand)
                else
                    'return error
                end if
            end if
        '$run
        elseif Left(constLine,4) = prefix & "run" then
            if Left(constLine,12) = prefix & "run.console" then
                shell.Run "cmd.exe"
            else
                Dim runCommand : runCommand = Replace(constLine, prefix & "run","")
                runCommand = Replace(runCommand, qt,"")
                shell.Run(runCommand)
            end if
        '$return
        elseif Left(constLine,7) = prefix & "return" then
            Dim returnCommand : returnCommand = Replace(constLine, prefix & "return","")
            returnCommand = Replace(returnCommand, "(","")
            returnCommand = Replace(returnCommand, ")","")
            returnCommand = Replace(returnCommand, qt,"")
            checkExtension = Left(returnCommand,4)
            if checkExtension = ".txt" then
                returnCommand = Replace(returnCommand,".txt","")
            elseif checkExtension= ".htm" then
                returnCommand = Replace(returnCommand,".htm","")
            elseif checkExtension = ".msg" then
                returnCommand = Replace(returnCommand,".msg","")
            end if
            if returnCommand = "eax" then
                returnCommand = eax
            elseif returnCommand = "ebx" then
                returnCommand = ebx
            elseif returnCommand = "ecx" then
                returnCommand = ecx
            elseif returnCommand = "edx" then
                returnCommand = edx
            elseif returnCommand = "rax" then
                returnCommand = rax
            elseif returnCommand = "rbx" then
                returnCommand = rbx
            elseif returnCommand = "rcx" then
                returnCommand = rcx
            elseif returnCommand = "rdx" then
                returnCommand = rdx
            end if
            'Register Value Filter
            returnCommand = Replace(returnCommand,nbsp,"")
            if checkExtension = ".txt" then
                fso.CreateTextFile "returnCommand.txt"
                Dim ret_txt : Set ret_txt = fso.OpenTextFile("returnCommand.txt",2)
                ret_txt.writeLine returnCommand
                ret_txt.close
                shell.Run "returnCommand.txt"
                WScript.Sleep 1000
                fso.DeleteFile "returnCommand.txt"
            elseif checkExtension= ".htm" then
                fso.CreateTextFile "returnCommand.html"
                Dim ret_htm : Set ret_htm = fso.OpenTextFile("returnCommand.html",2)
                ret_htm.writeLine "<h1>" & returnCommand & "</h1>"
                ret_htm.close
                shell.Run "returnCommand.html"
                WScript.Sleep 1000
                fso.DeleteFile "returnCommand.html"
            elseif checkExtension = ".msg" then
                fso.CreateTextFile "returnCommand.vbs"
                Dim ret_msg : Set ret_msg = fso.OpenTextFile("returnCommand.vbs",2)
                ret_msg.writeLine "Wscript.Echo " & qt & returnCommand & qt
                ret_msg.close
                shell.Run "returnCommand.vbs"
                WScript.Sleep 1000
                fso.DeleteFile "returnCommand.vbs"
            end if
        '$sleep
        elseif Left(constLine,6) = prefix & "sleep" then
            Dim sleepCommand : sleepCommand = Replace(constLine, prefix & "sleep","")
            sleepCommand = Replace(sleepCommand, "(","")
            sleepCommand = Replace(sleepCommand, ")","")
            checkExtension = Left(sleepCommand,4)
            if checkExtension = ".mil" then
                sleepCommand = Replace(sleepCommand, ".mil","")
                if sleepCommand = "eax" then
                    WScript.Sleep(eax)
                elseif sleepCommand = "ebx" then
                    WScript.Sleep(ebx)
                elseif sleepCommand = "ecx" then
                    WScript.Sleep(ecx)
                elseif sleepCommand = "edx" then
                    WScript.Sleep(edx)
                else
                    WScript.Sleep(sleepCommand)
                end if
            elseif checkExtension = ".sec" then
                sleepCommand = Replace(sleepCommand, ".sec","")
                if sleepCommand = "eax" then
                    WScript.Sleep(eax * 1000)
                elseif sleepCommand = "ebx" then
                    WScript.Sleep(ebx * 1000)
                elseif sleepCommand = "ecx" then
                    WScript.Sleep(ecx * 1000)
                elseif sleepCommand = "edx" then
                    WScript.Sleep(edx * 1000)
                else
                    WScript.Sleep(sleepCommand * 1000)
                end if
            end if
        '$play
        elseif Left(constLine,5) = prefix & "play" then
            Dim playCommand : playCommand = Replace(constLine, prefix & "play","")
            playCommand = Replace(playCommand, qt,"")
            sound.URL = playCommand
            sound.controls.play
            While sound.playState <> 1
                WScript.Sleep 100
            Wend
            sound.close
        '$quit
        elseif Left(constLine,5) = prefix & "quit" then
            WScript.Quit
        '$str
        elseif Left(constLine,4) = prefix & "str" then
            Dim str : str = Replace(constLine, prefix & "str" & nbsp,"")
            Dim strReg : strReg = Left(str,3)
            str = Replace(str,strReg,"")
            str = Replace(str,",","")
            str = Replace(str,qt,"")
            str = Replace(str,nbsp,"")
            'msgBox "Register: " & strReg & nw & "Value: " & str 'Debug
            if strReg = "rax" then
                if str = "%NULL" then
                    rax = NUL
                else
                    rax = str
                end if
            elseif strReg = "rbx" then
                if str = "%NULL" then
                    rbx = NUL
                else
                    rbx = str
                end if
            elseif strReg = "rcx" then
                if str = "%NULL" then
                    rcx = NUL
                else
                    rcx = str
                end if
            elseif strReg = "rdx" then
                if str = "%NULL" then
                    rdx = NUL
                else
                    rdx = str
                end if
            end if
        '$mov
        elseif Left(constLine,4) = prefix & "mov" then
            Dim mov : mov = Replace(constLine, prefix & "mov" & nbsp,"")
            Dim movRegister : movRegister = Left(mov,3)
            mov = Replace(mov,movRegister,"")
            mov = Replace(mov,",","")
            'msgBox "Register: " & movRegister & nw & "Value: " & mov 'Debug
            if movRegister = "eax" then
                if mov = " eax" then
                    eax = eax
                elseif mov = " ebx" then
                    eax = ebx
                elseif mov = " ecx" then
                    eax = ecx
                elseif mov = " edx" then
                    eax = edx
                elseif mov = " %NULL" then
                    eax = NUL
                else
                    eax = mov
                end if
            elseif movRegister = "ebx" then
                if mov = " eax" then
                    ebx = eax
                elseif mov = " ebx" then
                    ebx = ebx
                elseif mov = " ecx" then
                    ebx = ecx
                elseif mov = " edx" then
                    ebx = edx
                elseif mov = " %NULL" then
                    ebx = NUL
                else
                    ebx = mov
                end if
            elseif movRegister = "ecx" then
                if mov = " eax" then
                    ecx = eax
                elseif mov = " ebx" then
                    ecx = ebx
                elseif mov = " ecx" then
                    ecx = ecx
                elseif mov = " edx" then
                    ecx = edx
                elseif mov = " %NULL" then
                    ecx = NUL
                else
                    ecx = mov
                end if
            elseif movRegister = "edx" then
                if mov = " eax" then
                    edx = eax
                elseif mov = " ebx" then
                    edx = ebx
                elseif mov = " ecx" then
                    edx = ecx
                elseif mov = " edx" then
                    edx = edx
                elseif mov = " %NULL" then
                    edx = NUL
                else
                    edx = mov
                end if
            end if
        '$op
        elseif Left(constLine,3) = prefix & "op" then
            Dim operation : operation = Replace(constLine, prefix & "op","")
            operation = Replace(operation, "(","")
            operation = Replace(operation, ")","")
            operation = Replace(operation, ",","")
            checkExtension = Left(operation,4)
            if checkExtension = ".add" then
                operation = Replace(operation,".add","")
            elseif checkExtension = ".sub" then
                operation = Replace(operation,".sub","")
            elseif checkExtension = ".mul" then
                operation = Replace(operation,".mul","")
            elseif checkExtension = ".div" then
                operation = Replace(operation,".div","")
            end if
            'Operation Value Filter
            operation = Replace(operation,nbsp,"")
            rem.add
            if checkExtension = ".add" then
                if Left(operation,3) = "eax" then
                    operation = Replace(operation,"eax","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        eax = eax + eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        eax = eax + ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        eax = eax + ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        eax = eax + edx
                    else
                        operation = CInt(operation)
                        eax = eax + operation
                    end if
                elseif Left(operation,3) = "ebx" then
                    operation = Replace(operation,"ebx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        ebx = ebx + eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        ebx = ebx + ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        ebx = ebx + ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        ebx = ebx + edx
                    else
                        operation = CInt(operation)
                        ebx = ebx + operation
                    end if
                elseif Left(operation,3) = "ecx" then
                    operation = Replace(operation,"ecx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        ecx = ecx + eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        ecx = ecx + ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        ecx = ecx + ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        ecx = ecx + edx
                    else
                        operation = CInt(operation)
                        ecx = ecx + operation
                    end if
                elseif Left(operation,3) = "edx" then
                    operation = Replace(operation,"edx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        edx = edx + eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        edx = edx + ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        edx = edx + ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        edx = edx + edx
                    else
                        operation = CInt(operation)
                        edx = edx + operation
                    end if
                end if
            rem.add
            elseif checkExtension = ".sub" then
                if Left(operation,3) = "eax" then
                    operation = Replace(operation,"eax","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        eax = eax - eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        eax = eax - ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        eax = eax - ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        eax = eax - edx
                    else
                        operation = CInt(operation)
                        eax = eax - operation
                    end if
                elseif Left(operation,3) = "ebx" then
                    operation = Replace(operation,"ebx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        ebx = ebx - eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        ebx = ebx - ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        ebx = ebx - ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        ebx = ebx - edx
                    else
                        operation = CInt(operation)
                        ebx = ebx - operation
                    end if
                elseif Left(operation,3) = "ecx" then
                    operation = Replace(operation,"ecx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        ecx = ecx - eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        ecx = ecx - ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        ecx = ecx - ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        ecx = ecx - edx
                    else
                        operation = CInt(operation)
                        ecx = ecx - operation
                    end if
                elseif Left(operation,3) = "edx" then
                    operation = Replace(operation,"edx","")
                    if operation = "eax" then
                        eax = CInt(eax)
                        edx = edx - eax
                    elseif operation = "ebx" then
                        ebx = CInt(ebx)
                        edx = edx - ebx
                    elseif operation = "ecx" then
                        ecx = CInt(ecx)
                        edx = edx - ecx
                    elseif operation = "edx" then
                        edx = CInt(edx)
                        edx = edx - edx
                    else
                        operation = CInt(operation)
                        edx = edx - operation
                    end if
                end if
            end if
        '$sk
        elseif Left(constLine,3) = prefix & "sk" then
            Dim skCommand : skCommand = Replace(constLine, prefix & "sk","")
            skCommand = Replace(skCommand, "(","")
            skCommand = Replace(skCommand, ")","")
            skCommand = Replace(skCommand,qt,"")
            if skCommand = "eax" then
                shell.SendKeys eax
            elseif skCommand = "ebx" then
                shell.SendKeys ebx
            elseif skCommand = "ecx" then
                shell.SendKeys ecx
            elseif skCommand = "edx" then
                shell.SendKeys edx
            elseif skCommand = "rax" then
                shell.SendKeys rax
            elseif skCommand = "rbx" then
                shell.SendKeys rbx
            elseif skCommand = "rcx" then
                shell.SendKeys rcx
            elseif skCommand = "rdx" then
                shell.SendKeys rdx
            else
                shell.SendKeys skCommand
            end if
        end if
    Loop
End Sub
