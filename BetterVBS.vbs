Option Explicit
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim shell : Set shell = CreateObject("WScript.Shell")
Dim nw : nw = vbCrlf
Dim qt : qt = """"

'Global Variables
Dim filePath

Call subConsole()
Sub subConsole()
    Dim console : console = InputBox("","Shard Console")
    '$run
    if InStr(console, "$run") > 0 then
        filePath = Replace(console, "$run -","")
        Call compile()
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
        end if
    Loop
End Sub
