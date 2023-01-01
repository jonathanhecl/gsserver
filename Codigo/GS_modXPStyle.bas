Attribute VB_Name = "GS_modXPStyle"
Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Sub cMain()
' NOTE: THE NEXT LINES MUST BE THE FIRST LINES
'       IN YOUR PROJECT.
'       OTHER WISE.. YOU WILL RECIVE AN ERROR
'       MESSAGE FROM WINDOWS AND YOUR APPLICATION
'       WILL NOT RUN.
InitCommonControls
XPStyle

End Sub


Function UnStyle() As Boolean
On Error Resume Next
Dim strManifest As String
strManifest = App.Path & "\" & App.EXEName & ".exe.manifest"
SetAttr strManifest, vbNormal
Kill strManifest
If Err.Number = 0 Then UnStyle = True
End Function

'This function will automatically write a small XML file (manifest) in your program folder
'Which carries some info about your program and let windows draw your application.
'It won't be affected by the exe name.
'
' Written by: Voodoo Attack!!
'     E-Mail: voodooattack@hotmail.com
'

Function XPStyle(Optional AutoRestart As Boolean = True, Optional CreateNew As Boolean) As Boolean
InitCommonControls
On Error Resume Next
Dim XML             As String
Dim ManifestCheck   As String
Dim strManifest     As String
Dim FreeFileNo      As Integer

If AutoRestart = True Then CreateNew = False

'(put the XML in a string)
XML = ("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> " & vbCrLf & "<assembly " & vbCrLf & "   xmlns=""urn:schemas-microsoft-com:asm.v1"" " & vbCrLf & "   manifestVersion=""1.0"">" & vbCrLf & "<assemblyIdentity " & vbCrLf & "    processorArchitecture=""x86"" " & vbCrLf & "    version=""EXEVERSION""" & vbCrLf & "    type=""win32""" & vbCrLf & "    name=""EXENAME""/>" & vbCrLf & "    <description>EXEDESCRIBTION</description>" & vbCrLf & "    <dependency>" & vbCrLf & "    <dependentAssembly>" & vbCrLf & "    <assemblyIdentity" & vbCrLf & "         type=""win32""" & vbCrLf & "         name=""Microsoft.Windows.Common-Controls""" & vbCrLf & "         version=""6.0.0.0""" & vbCrLf & "         publicKeyToken=""6595b64144ccf1df""" & vbCrLf & "         language=""*""" & vbCrLf & "         processorArchitecture=""x86""/>" & vbCrLf & "    </dependentAssembly>" & vbCrLf & "    </dependency>" & vbCrLf & "</assembly>" & vbCrLf & "")

'don't be confused.. i did not write this by hand!
'i made a special text coverter for cases like this. :D

strManifest = App.Path & "\" & App.EXEName & ".exe.manifest"    'set the name of the manifest
ManifestCheck = Dir(strManifest, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive) 'check the app manifest file.
If ManifestCheck = "" Or CreateNew = True Then           'if not found.. make a new one
XML = Replace(XML, "EXENAME", App.EXEName & ".exe")             'Replaces the string "EXENAME" with the program's exe file name.
XML = Replace(XML, "EXEVERSION", App.Major & "." & App.Minor & "." & App.Revision & ".0") 'Replaces the "EXEVERSION" string.
XML = Replace(XML, "EXEDESCRIBTION", App.FileDescription)       'Replaces the app Describtion.
FreeFileNo = FreeFile       'get the next avilabel file
If ManifestCheck <> "" Then
SetAttr strManifest, vbNormal
Kill strManifest
End If
Open strManifest For Binary As #(FreeFileNo) 'open the file
Put #(FreeFileNo), , XML    'uses 'put' to set the file content.. note that 'put' (binary mode) is much faster than 'print'(output mode)
Close #(FreeFileNo)         'close the file.
SetAttr strManifest, vbHidden + vbSystem
If ManifestCheck = "" Then
XPStyle = False             'return false.. this means that the file does not exist
Else
XPStyle = True
End If
If AutoRestart = True Then  'if in automode (default), the program will restart.
Shell App.Path & "\" & App.EXEName & ".exe " & Command$, vbNormalFocus  'restart the program and bypass command line parameters (if any)
End                         'end the session.
End If
Else                'the manifest file exists.
XPStyle = True      'return true.
End If
End Function

