Dim url
url = Array("https://forms.gle/n21jJycDQC6p9fuD8", "https://forms.gle/b8qZwvFTyq5Tiop96", "https://forms.gle/PfdioEXhuGQew4DV8", "https://forms.gle/R8KdsqkAYvgdbbZKA")
'all urls that you want to pick from

Randomize
Dim randoNumber
randoNumber = Int((3 + 1) * Rnd)
'generate a random number from 0 to size url of array - 1 ---> Int((<size of url array - 1> + 1) * Rnd)

Dim rando
rando = url(randoNumber)
'gets the random url from the url array

Dim shell
Set shell = WScript.CreateObject("WScript.Shell")
shell.Run "msedge " & rando & " --hide-scrollbars --content-shell-hide-toolbar"
'opens microsoft edge with the random url