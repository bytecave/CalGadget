OpenWindow(0, 0, 0, 0, 0, "", #PB_Window_Invisible)
OpenWindow(1, 0, 0, 500, 300, "", #PB_Window_ScreenCentered | #PB_Window_BorderLess, WindowID(0))

XIncludeFile "iCalModule.pbi"

InitNetwork()

Filename$ = "C:\Users\bytecave\Desktop\msft.ics"

  If ReceiveHTTPFile("https://outlook.office365.com/owa/calendar/aff5ddd2d53c4603a3bc9d3f8ac38e7a@microsoft.com/20a729b46729493da38146d0da69061e1216144310575959539/calendar.ics", Filename$)
    Debug "Success"
  Else
    Debug "Failed"
  EndIf
  
  
    NewList Events.iCal::Event_Structure()
    
    If iCal::Create(500)
      
    iCal::ImportFile(500, "C:\Users\bytecave\Desktop\msft.ics")
    
        iCal::GetEvents(500, Events())
        ForEach Events()
          
          ;Events()\StartDate   = iCal()\Event()\StartDate
          ;Events()\EndDate     = iCal()\Event()\EndDate
          ;Events()\Summary     = iCal()\Event()\Summary
          ;Events()\Description = iCal()\Event()\Description
          ;Events()\Location    = iCal()\Event()\Location
          ;Events()\Class       = iCal()\Event()\Class
          
          With Events()
            Debug "Event: " + \Summary + ": " + \Location
            Debug "Start: " + FormatDate("%mm/%dd %hh:%ii", \StartDate)
            Debug "End: " + FormatDate("%mm/%dd %hh:%ii", \EndDate)
            Debug "Class: " + \Class
            Debug "Description: " + \Description
          EndWith
          
    Next
  EndIf
  
Repeat
  Event = WaitWindowEvent()
  If Event = #PB_Event_RightClick
    End
  ElseIf Event = #WM_LBUTTONDOWN
    SendMessage_(WindowID(1),#WM_NCLBUTTONDOWN, #HTCAPTION,0)
  EndIf
ForEver

