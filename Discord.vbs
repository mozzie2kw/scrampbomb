Set WshShell = CreateObject("WScript.Shell")
StringsToType = Array(":3", ">w<", "uwu", "owo", "X3")
do
wscript.sleep.100
WshShell.Sendkeys "{CAPSLOCK}"
WScript.Sleep 100
WshShell.SendKeys "{NUMLOCK}"
WScript.Sleep 100

RandomIndex = Int(Rnd * UBound(StringsToType) + 1)
    StringToSend = StringsToType(RandomIndex)

WshShell.SendKeys StringToSend
loop
