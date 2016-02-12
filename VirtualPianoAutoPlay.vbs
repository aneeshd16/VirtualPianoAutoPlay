'Put music sheet here. Remove all new lines
str = "[et] | [o0] | [wr] | 9  [et] | [8a] | [wr] | 9  [etpa] | [0uoh] | [wroa] | [9y] [etpa] | [8tak] | [wroa] | [9y] [6ts] t r e [3oh] 0 Q 0  [5ra] r e w 2 9 0 9  [6ts] t r e [1ak] 8 e w  [5ra] r e w 2 9 u [yd] [6ts]e [6s]e [6s]e [6p]e [0oah] u I u  [5ra]w [5a]w [5p]w [5o]w [9p] y u y  [6ts]e [6s]e [6a]e [6p]e [8afk] t p o  [5ra]w [5a]w [5p]w [5p]w [5o]w [yp] [ypd] [9uaf] [2yd] Alternative  [6et] 6 6 6 [3ro] 3 3 3 [5wr] 5 5 5 [2e] 2 2 2  [6et] 6 6 6 [10wr] 1 1 1 [59wr] 5 5 5 [29] 2 2 2  [680t] 6 6 6 [30wro] 3 3 3 [57wr] 5 5 5 [29] 2 2 2  [et680] 6 6 6 [180wr] 1 1 1 [59wr] 5 5 5 [29] 2 2 2  [68e] 7 6 [350w] Q w [57wr] 6 7 [269] 0 Q  [6etps] r e [30wuk] I o [5wrya] e r [29p] u I  [6ets] r e 0 [30wro] [5wra] e r y [29pd] [6etps] r e 0 [18etuoa] e w 0 [5wra] e r y [29d] [60esl] [ra] [ep] [0u] [30wh] [QI] [wo] [ep] [5wak] [ep] [9y] [0u] [29] [0u] [QI] [ep] [60esl] [ra] [ep] [0uoa] [1afhk] [ep] [wo] [ep] [5wrya] [ep] [wo] [odkz] [29z] [0u] [9y] [0u] [6esfl] [ak] [uf] [37ofh] [IG] [ra] [5wak] [pj] [yd] [269] [uf] [ep] [60sfl] [sl] [uf] [15ahk] [oh] [ts] [59rya] [pj] [yd] [69] [uf] [ep] [6ps] [jl] [30h] [af] [5oa] [dk] [2y] [fhk] [psjl] [tafk] [oafk] [ypd] [ps] [uh] [oa] [y] [ps] [tk] [oa] [ydf] GGGGGGGG"  

'Time interval between notes'
timeInterval = 500

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate("chrome")
WScript.Sleep 2000
interval = timeInterval
For i=1 To Len(str)
	If Mid(str,i,1) = "[" Then
		interval = 0
	ElseIf Mid(str,i,1) = "]" Then
		interval = timeInterval
		WScript.Sleep interval
	Else
		WshShell.SendKeys Mid(str,i,1)
		WScript.Sleep interval
	End If
Next