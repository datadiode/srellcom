# SRELL Regular Expression COM Wrapper (srellcom)

The srellcom project aims at providing an [SRELL](https://www.akenotsuki.com/misc/srell/) based VBScript.RegExp replacement.

Here is a small VBScript example which uses a Unicode property escape to know that &#960; is a Greek letter:
```
Dim re
Set re = CreateObject("SRELL.RegExp")
re.Pattern = "\p{Script=Greek}"
wscript.Echo "It is " & re.Test(ChrW(960)) & " that " & ChrW(960) & " is a Greek letter." 
```
