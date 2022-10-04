''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VBScript example for SRELL.RegExp COM Wrapper
' Based on https://github.com/ZimProjects/SRELL's sample01.cpp
' Implements the tests but not the benchmarks
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim RegExpClass
RegExpClass = Replace("SRELL.RegExp", "SRELL", "VBScript", 1, WScript.Arguments.Named("builtin-regexp"))

Function Test(str, exp, max, expected)
	Dim re, mr, sm, i, n, placeholder, matched, msg, num_of_failures
	Set re = CreateObject(RegExpClass)
	re.Pattern = exp
	For i = 1 To max
		Set mr = re.Execute(str)
	Next
	WScript.Echo vbTab & """" & str & """ =~ /" & exp & "/"
	If max > 1 Then WScript.Echo vbTab & max & " times"
	WScript.Echo vbTab & Replace("Not Found", "Not ", "", 1, mr.Count)
	If mr.Count <> 0 Then
		Set sm = mr(0).SubMatches
		n = sm.Count
		For i = 0 To n
			placeholder = Replace("$&", "&", i, 1, i)
			If i = 0 Then
				matched = mr(0)
			Else
				matched = sm(i - 1)
			End If
			msg = vbTab & placeholder & " = """ & matched & """"
			If i < UBound(expected) Then
				if matched = expected(i) Or matched = "" And expected(i) = "(undefined)" Then
					msg = msg & "; passed!"
				Else
					msg = msg & "; failed... (expected: """ & expected(i) & """)"
					num_of_failures = num_of_failures + 1
				End If
			Else
				msg = msg & "; failed..." ' should not exist.
				num_of_failures = num_of_failures + 1
			End If
			If mr(0) <> "" And re.Replace(mr(0), placeholder) <> matched Then
				msg = msg & "; replace failed..." ' should have yielded same result.
				num_of_failures = num_of_failures + 1
			End If
			WScript.Echo msg
		Next
	End If
	If num_of_failures = 0 And UBound(expected) <> n + 1 Then
		num_of_failures = num_of_failures + 1
	End If
	WScript.Echo Replace("Result: passed.", "passed", "failed", 1, num_of_failures)
	WScript.Echo
	Test = 1 - Sgn(num_of_failures)
End Function

Function Main
	Dim str, exp, expected
	Dim num_of_tests
	Dim num_of_tests_passed

	WScript.Echo "RegExpClass = " & RegExpClass
	WScript.Echo

	WScript.Echo "Test 1 (ECMAScript 2021 Language Specification 22.2.2.3, NOTE)"
	str = "abc"
	exp = "((a)|(ab))((c)|(bc))"
	ReDim expected(7)
	expected(0) = "abc"
	expected(1) = "a"
	expected(2) = "a"
	expected(3) = "(undefined)"
	expected(4) = "bc"
	expected(5) = "(undefined)"
	expected(6) = "bc"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 2a (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 2)"
	str = "abcdefghi"
	exp = "a[a-z]{2,4}"
	ReDim expected(1)
	expected(0) = "abcde"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 2b (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 2)"
	str = "abcdefghi"
	exp = "a[a-z]{2,4}?"
	expected(0) = "abc"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 3 (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 2)"
	str = "aabaac"
	exp = "(aa|aabaac|ba|b|c)*"
	ReDim expected(2)
	expected(0) = "aaba"
	expected(1) = "ba"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 4 (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 3)"
	str = "zaacbbbcac"
	exp = "(z)((a+)?(b+)?(c))*"
	ReDim expected(6)
	expected(0) = "zaacbbbcac"
	expected(1) = "z"
	expected(2) = "ac"
	expected(3) = "a"
	expected(4) = "(undefined)"
	expected(5) = "c"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 5a (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 4)"
	str = "b"
	exp = "(a*)*"
	ReDim expected(2)
	expected(0) = ""
	expected(1) = ""
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 5b (ECMAScript 2021 Language Specification 22.2.2.5.1, NOTE 4)"
	str = "baaaac"
	exp = "(a*)b\1+"
	expected(0) = "b"
	expected(1) = ""
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 6a (ECMAScript 2021 Language Specification 22.2.2.8.2, NOTE 2)"
	str = "baaabac"
	exp = "(?=(a+))"
	expected(0) = ""
	expected(1) = "aaa"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 6b (ECMAScript 2021 Language Specification 22.2.2.8.2, NOTE 2)"
	str = "baaabac"
	exp = "(?=(a+))a*b\1"
	expected(0) = "aba"
	expected(1) = "a"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 7 (ECMAScript 2021 Language Specification 22.2.2.8.2, NOTE 3)"
	str = "baaabaac"
	exp = "(.*?)a(?!(a+)b\2c)\2(.*)"
	ReDim expected(4)
	expected(0) = "baaabaac"
	expected(1) = "ba"
	expected(2) = "(undefined)"
	expected(3) = "abaac"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Test 8 (from https://github.com/tc39/test262/tree/master/test/built-ins/RegExp/lookBehind/misc.js)"
	str = "abc"
	exp = "(abc\1)"
	ReDim expected(2)
	expected(0) = "abc"
	expected(1) = "abc"
	num_of_tests_passed = num_of_tests_passed + Test(str, exp, 1, expected)
	num_of_tests = num_of_tests + 1

	WScript.Echo "Results of tests: " & num_of_tests_passed & "/" & num_of_tests & " passed."
	
	Main = num_of_tests - num_of_tests_passed
End Function

wscript.Quit Main()
