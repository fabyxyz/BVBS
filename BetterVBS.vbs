Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim nw : nw = vbCrlf
Dim tb : tb = vbTab
Dim qt : qt = """"

Dim debugMode : debugMode = false
Dim debugRunCommand
if debugMode = true then
    debugRunCommand = "$run -test.bvbs"
else
    debugRunCommand = ""
end if
    
'Global Variables
Dim filePath
Dim checkExtension

Call subConsole()
Sub subConsole()
    Dim console : console = InputBox("","BetterVBS Console",debugRunCommand)
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
        constLine = file.ReadLine
        '/cd
        if InStr(constLine, prefix & "cd") > 0 then
            Dim cdCommand : cdCommand = Replace(constLine, prefix & "cd ","")
            cdCommand = Replace(cdCommand, qt,"")
            Dim createFolderPath, createFolderName
            Dim cdSection : cdSection = Split(cdCommand,",")
            createFolderPath = cdSection(0)
            createFolderName = cdSection(1)
            fso.CreateFolder(createFolderPath & "/" & createFolderName)
        '/run
        elseif InStr(constLine, prefix & "run") > 0 then
            Dim runCommand : runCommand = Replace(constLine, prefix & "run ","")
            runCommand = Replace(runCommand, qt,"")
            shell.Run(runCommand)
        '/return
        elseif InStr(constLine, prefix & "return") > 0 then
            Dim returnCommand : returnCommand = Replace(constLine, prefix & "return","")
            returnCommand = Replace(returnCommand, "(","")
            returnCommand = Replace(returnCommand, ")","")
            returnCommand = Replace(returnCommand, qt,"")
            checkExtension = Left(returnCommand,4)
            if checkExtension = ".txt" then
                returnCommand = Replace(returnCommand, ".txt","")
                fso.CreateTextFile "returnCommand.txt"
                Dim ret_txt : Set ret_txt = fso.OpenTextFile("returnCommand.txt",2)
                ret_txt.writeLine returnCommand
                ret_txt.close
                shell.Run "returnCommand.txt"
                WScript.Sleep 1000
                fso.DeleteFile "returnCommand.txt"
            elseif checkExtension= ".htm" then
                returnCommand = Replace(returnCommand, ".htm","")
                fso.CreateTextFile "returnCommand.html"
                Dim ret_htm : Set ret_htm = fso.OpenTextFile("returnCommand.html",2)
                ret_htm.writeLine "<p>" & returnCommand & "</p>"
                ret_htm.close
                shell.Run "returnCommand.html"
                WScript.Sleep 1000
                fso.DeleteFile "returnCommand.html"
            end if
        end if
    Loop
End Sub
