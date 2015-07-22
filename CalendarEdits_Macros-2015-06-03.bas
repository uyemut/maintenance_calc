Sub CalTimeEdits()
'
' CalEdits Macro
'

'
    Range("A1:A2").Select
    Cells.Replace What:=".5-", Replacement:=":30-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=".5pm", Replacement:=":30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=".5am", Replacement:=":30am", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    ActiveCell.Replace What:="-0:30pm", Replacement:="12:30pm", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=" 0:30pm", Replacement:=" 12:30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="-0:30pm", Replacement:="-12:30pm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:=" 0:30am", Replacement:=" 12:30am", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub CalDescriptionEdits()
'
' CalEdits Macro
'

'
    Range("A1:A2").Select
    Cells.Replace What:="E 12-4pm - Red Hats", _
	Replacement:="E 1-4pm - Red Hats", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D 4-6pm - Gospel Singers", _
	Replacement:="D 5-6pm - Gospel Singers", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9-11am - Green Thumb Club", _
	Replacement:="C 10-11am - Green Thumb Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 4-10pm - Polish Club Monthly Meeting", _
	Replacement:="A 5-10pm - Polish Club Monthly Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 1-4pm - Senior Singles", _
	Replacement:="A 2-4pm - Senior Singles", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="9:30-12pm - Violations Comm", _
	Replacement:="10am - Violations Comm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="9:30-12am - Violations Comm", _
	Replacement:="10am - Violations Comm", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="12-5pm - Board of Trustee", _
	Replacement:="1pm - Board of Trustee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7:30-9:30pm - Yacht Club", _
	Replacement:="7:30-9:30pm-Boat & Fishing Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="A 6:30-11pm - HOA Board Meeting", _
	Replacement:="A 7pm - HOA Board Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9:30-12pm - VETS Council", _
	Replacement:="C 10-12pm - VETS Council", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="C 9:30-12am - VETS Council", _
	Replacement:="C 10-12pm - VETS Council", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="1-2:30pm - RV CLUB", _
	Replacement:="1:30-2:30pm - RV CLUB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="E 1-4pm - Recreation Committee", _
	Replacement:="E 2-4pm - Recreation Committee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D 2-4:30pm - BFB Conservative Club", _
	Replacement:="D 2:30-4:30pm - BFB Conservative Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 6:30-9pm - Democratic Club", _
	Replacement:="D&E 7-9pm - Democratic Club", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 5-10pm - Computer Club General Meeting", _
	Replacement:="D&E 6-10pm - Computer Club General Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 5-10pm - Computer Club Round Table Meeting", _
        Replacement:="D&E 6-10pm - Computer Club Round Table Mtg.", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="E 12:30-3pm - HOA - Orientation", _
	Replacement:="E 1-3pm - HOA - Orientation", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="7-10pm Board of Trustee", _
	Replacement:="7pm - Board of Trustee", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="R 9-11am - ARCC Meeting", _
	Replacement:="RR 9am - ARCC Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Replace What:="D&E 11-3pm - Men's Golf - General Meeting", _
	Replacement:="D&E 11:30-3pm - Men's Golf - General Meeting", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'    IF  Range("A1").Cells == "JUNE" or "JULY" or "AUGUST"
'        Cells.Replace What:="Computer Club General Meeting", _
'	Replacement:="Computer Club Round Table Mtg", L'ookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False
'    END-IF

End Sub


'7  *** DRAFT *** *** DRAFT ***  FOR INTERNAL DISTRIBUTION ONLY
'D&E 11:30-3pm - Men's Golf - General Meeting
