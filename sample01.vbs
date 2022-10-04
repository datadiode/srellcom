''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VBScript example for SRELL.RegExp COM Wrapper (tests and benchmarks)
' Based on https://github.com/ZimProjects/SRELL's sample01.cpp
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Dim RegExpClass
RegExpClass = Replace("SRELL.RegExp", "SRELL", "VBScript", 1, WScript.Arguments.Named("builtin-regexp"))

Dim count
count = CInt(WScript.Arguments.Named("benchmark"))

Function Test(str, exp, max, expected)
	Dim re, mr, sm, i, n, placeholder, matched, msg, num_of_failures, st, ed
	Set re = CreateObject(RegExpClass)
	re.Pattern = exp
	st = Timer
	For i = 1 To max
		Set mr = re.Execute(str)
	Next
	ed = Timer
	WScript.Echo vbTab & """" & str & """ =~ /" & exp & "/"
	If max > 1 Then WScript.Echo vbTab & max & " times"
	WScript.Echo vbTab & Replace("Not Found", "Not ", "", 1, mr.Count) & " (" & Int((ed - st) * 1000) & " msec)"
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
	Else
		n = -1
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
	Dim num_of_benches
	Dim num_of_benches_passed

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

	If count <> 0 Then

		WScript.Echo "Benchmark 01"
		      '0123456'
		str = "aaaabaa"
		exp = "^(.*)*b\1$"
		ReDim expected(2)
		expected(0) = "aaaabaa"
		expected(1) = "aa"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 02"
		      '012345678'
		str = "aaaabaaaa"
		exp = "^(.*)*b\1\1$"
		expected(0) = "aaaabaaaa"
		expected(1) = "aa"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 03"
		      '01'
		str = "ab"
		exp = "(.*?)*b\1"
		expected(0) = "b"
		expected(1) = ""
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 10, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 04"
		      '01234567'
		str = "acaaabbb"
		exp = "(a(.)a|\2(.)b){2}"
		ReDim expected(4)
		expected(0) = "aaabb"
		expected(1) = "bb"
		expected(2) = "(undefined)"
		expected(3) = "b"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 10, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 05"
		str = "aabbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbaaaaaa"
		exp = "(a*)(b)*\1\1\1"
		ReDim expected(3)
		expected(0) = "aabbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbaaaaaa"
		expected(1) = "aa"
		expected(2) = "b"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 06a"
		str = "aaaaaaaaaab"
		exp = "(.*)*b"
		ReDim expected(2)
		expected(0) = "aaaaaaaaaab"
		expected(1) = "aaaaaaaaaa"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 10, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 06b"
		str = "aaaaaaaaaab"
		exp = "(.*)+b"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 10, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 06c"
		str = "aaaaaaaaaab"
		exp = "(.*){2,}b"
		expected(1) = ""
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 10, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 07"
		str = "aaaaaaaaaabc"
		exp = "(?=(a+))(abc)"
		ReDim expected(3)
		expected(0) = "abc"
		expected(1) = "a"
		expected(2) = "abc"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 08"
		str = "1234-5678-1234-456"
		exp = "(\d{4}[-]){3}\d{3,4}"
		ReDim expected(2)
		expected(0) = "1234-5678-1234-456"
		expected(1) = "1234-"
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, count * 5, expected)
		num_of_benches = num_of_benches + 1

		WScript.Echo "Benchmark 09"
		str = "aaaaaaaaaaaaaaaaaaaaa"
		exp = "(.*)*b"
		ReDim expected(0)
		num_of_benches_passed = num_of_benches_passed + Test(str, exp, 1, expected)
		num_of_benches = num_of_benches + 1

	End If

	WScript.Echo "Results of tests: " & num_of_tests_passed & "/" & num_of_tests & " passed."
	WScript.Echo "Results of benchmarks: " & num_of_benches_passed & "/" & num_of_benches & " passed."

	Main = num_of_tests - num_of_tests_passed
End Function

wscript.Quit Main()
