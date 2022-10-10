''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VBScript to run some PCRE2 tests against SRELL.RegExp COM Wrapper
' Refer to https://github.com/PCRE2Project/pcre2/blob/master/testdata
' for suitable input files, esp. testinput1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim RegExpClass
RegExpClass = Replace("SRELL.RegExp", "SRELL", "VBScript", 1, WScript.Arguments.Named("builtin-regexp"))

Dim re
Set re = CreateObject(RegExpClass)

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

Sub RunTest(testinput)
	Dim stm, line, j, c, identified, compilable, uncompilable, ignored, polarity, failed, passed, messed, matched
	Set stm = fs.OpenTextFile(testinput)
	While Not stm.AtEndOfStream
		line = stm.ReadLine
		If InStrRev(line, "/", 1) = 1 Then
			Do
				j = InStrRev(line, "/")
				If j > 1 Then
					' guess if this is the end of the regexp
					If Mid(line, j - 1, 1) <> "\" Then
						c = Mid(line, j + 1, 1)
						If Len(c) = 0 Or UCase(c) <> LCase(c) Then Exit Do
					End If
				End If
				line = line & stm.ReadLine
			Loop Until stm.AtEndOfStream
			WScript.Echo line
			re.pattern = Mid(line, 2, j - 2)
			re.IgnoreCase = InStrRev(line, "i") > j
			re.MultiLine = InStrRev(line, "m") > j
			identified = identified + 1
			If InStrRev(line, "s") > j Or InStrRev(line, "x") > j Then
				' srellcom does not support the s and x modifiers
				ignored = ignored + 1
				polarity = 0
			Else
				On Error Resume Next
				re.Test ""
				If Err.Number = 0 Then
					polarity = 1
					compilable = compilable + 1
				Else
					polarity = 0
					WScript.Echo Err.Source & "-" & Err.Number & ": " & Err.Description
					uncompilable = uncompilable + 1
				End If
				On Error GoTo 0
			End If
		ElseIf InStrRev(line, "\=", 2) = 1 Then
			polarity = -polarity
		ElseIf InStrRev(line, "#", 1) = 0 Then
			line = Trim(line)
			line = Replace(line, "\r", vbCr)
			line = Replace(line, "\n", vbLf)
			line = Replace(line, "\t", vbTab)
			line = Replace(line, "\f", Chr(12))
			line = Replace(line, "\\", "\")
			line = Replace(line, "\'", "'")
			line = Replace(line, "\""", """")
			If polarity <> 0 And line <> "" Then
				WScript.Echo line
				On Error Resume Next
				matched = re.Test(line)
				If Err.Number <> 0 Then
					messed = messed + 1
					WScript.Echo "*** MESSED ***"
				ElseIf polarity > 0 Xor matched Then
					failed = failed + 1
					WScript.Echo "*** FAILED ***"
				Else
					passed = passed + 1
				End If
				On Error GoTo 0
			End If
		End If
	WEnd
	WScript.Echo
	WScript.Echo "[" & testinput & "]"
	WScript.Echo "identified expressions = " & identified
	WScript.Echo "compilable expressions = " & compilable
	WScript.Echo "uncompilable expressions = " & uncompilable
	WScript.Echo "ignored expressions = " & ignored
	WScript.Echo "passed recognitions = " & passed
	WScript.Echo "failed recognitions = " & failed
	WScript.Echo "messed recognitions = " & messed
	WScript.Echo
End Sub

WScript.Echo "RegExpClass = " & RegExpClass
WScript.Echo

Dim testinput
For Each testinput In WScript.Arguments.Unnamed
	RunTest testinput	
Next
