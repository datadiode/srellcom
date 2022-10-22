# SRELL Regular Expression COM Wrapper (srellcom)
[![StandWithUkraine](https://raw.githubusercontent.com/vshymanskyy/StandWithUkraine/main/badges/StandWithUkraine.svg)](https://github.com/vshymanskyy/StandWithUkraine/blob/main/docs/README.md)
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Fdatadiode%2Fsrellcom.svg?type=shield)](https://app.fossa.com/projects/git%2Bgithub.com%2Fdatadiode%2Fsrellcom?ref=badge_shield)

The srellcom project aims at providing an [SRELL](https://www.akenotsuki.com/misc/srell/en/) based VBScript.RegExp replacement.

Here is a small VBScript example which uses a Unicode property escape to know that &#960; is a Greek letter:
```
Dim re
Set re = CreateObject("SRELL.RegExp")
re.Pattern = "\p{Script=Greek}"
wscript.Echo "It is " & re.Test(ChrW(960)) & " that " & ChrW(960) & " is a Greek letter." 
```


## License
[![FOSSA Status](https://app.fossa.com/api/projects/git%2Bgithub.com%2Fdatadiode%2Fsrellcom.svg?type=large)](https://app.fossa.com/projects/git%2Bgithub.com%2Fdatadiode%2Fsrellcom?ref=badge_large)