OpenWindow(0, 0, 0, 0, 0, "", #PB_Window_Invisible)
OpenWindow(1, 0, 0, 600, 400, "", #PB_Window_ScreenCentered | #PB_Window_BorderLess, WindowID(0))
cvsCalendar = CanvasGadget(#PB_Any, 0, 0, 600, 400, #PB_Canvas_Border)

XIncludeFile "iCalModule.pbi"

Filename$ = "C:\Users\bytecave\Desktop\msft.ics"

Procedure DownloadCalendar(*value)
zFilename$ = "C:\Users\bytecave\Desktop\msft.ics"

InitNetwork()

  iCalDownload = ReceiveHTTPFile("https://outlook.office365.com/owa/calendar/aff5ddd2d53c4603a3bc9d3f8ac38e7a@microsoft.com/20a729b46729493da38146d0da69061e1216144310575959539/calendar.ics", zFilename$, #PB_HTTP_Asynchronous)
  
  If iCalDownload
    Repeat
      Progress = HTTPProgress(iCalDownload)
      
      Select Progress
        Case #PB_HTTP_Success
          *Buffer = FinishHTTP(iCalDownload)
          Debug "Download finished (size: " + MemorySize(*Buffer) + ")"
          FreeMemory(*Buffer)
          End

        Case #PB_HTTP_Failed
          Debug "Download failed"
          FinishHTTP(iCalDownload)
          End

        Case #PB_HTTP_Aborted
          Debug "Download aborted"
          FinishHTTP(iCalDownload)
          End
          
        Default
          Debug "Current download: " + Progress
       
      EndSelect
      
      Delay(50) ; Don't stole the whole CPU
    ForEver
  EndIf
EndProcedure
  
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
            ;Debug "Event: " + \Summary + ": " + \Location
            ;Debug "Start: " + FormatDate("%mm/%dd %hh:%ii", \StartDate)
            ;Debug "End: " + FormatDate("%mm/%dd %hh:%ii", \EndDate)
            ;Debug "Class: " + \Class
            ;Debug "Description: " + \Description
          EndWith
          
    Next
  EndIf
  
    CreateThread(@DownloadCalendar(), 0)
  

Repeat
  Event = WaitWindowEvent()
  If Event = #PB_Event_RightClick
    End
  ElseIf Event = #WM_LBUTTONDOWN
    SendMessage_(WindowID(1),#WM_NCLBUTTONDOWN, #HTCAPTION,0)
  EndIf
ForEver


; IDE Options = PureBasic 5.71 LTS (Windows - x64)
; CursorPosition = 42
; FirstLine = 4
; Folding = -
; EnableThread
; EnableXP
; DPIAware
; EnableCompileCount = 17
; EnableBuildCount = 0