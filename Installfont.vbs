FontFile = WScript.Arguments(0)
Const FONTS = &H14&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(FONTS)
objFolder.CopyHere FontFile