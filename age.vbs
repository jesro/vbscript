'Create a basic continue/cancel message.
Set win = New Window
win.SetTitle = "Age Warning!"
win.Icon = "%WINDIR%\System32\DFDWiz.exe"
win.ContextMenu = "no"
win.Scroll = "no"
win.SetHeight = 180

'Add CSS styling
win.AddStyles = "body{text-align:center;margin-top:5px;padding:25px;font-size:11pt;}div{padding:25px;font-size:11pt;}"
win.AddStyles = "input{border:0;border-radius:2px;padding:5px 10px;margin:5px;background-color:#e2e2e2;}"

'CLose already opened window
Set objSWbemServices = GetObject ("WinMgmts:Root\Cimv2") 
Set colProcess = objSWbemServices.ExecQuery _ 
("Select * From Win32_Process where name = 'wscript.exe'")
Dim strReport
For Each objProcess in colProcess
    If InStr (objProcess.CommandLine, "age.vbs") > 0 Then
        'strReport = strReport & vbNewLine & vbNewLine & _
         '   "ProcessId: " & objProcess.ProcessId & vbNewLine & _
          '  "ParentProcessId: " & objProcess.ParentProcessId & _
           ' vbNewLine & "CommandLine: " & objProcess.CommandLine & _
           ' vbNewLine & "Caption: " & objProcess.Caption & _
           ' vbNewLine & "ExecutablePath: " & objProcess.ExecutablePath
            Set oShell = CreateObject("WScript.Shell") 
            If oShell.AppActivate("Age Warning!") Then
                WScript.Sleep 500
                 oShell.SendKeys "%{F4}"
           End If
    End If
Next
'WScript.Echo strReport

'Add HTML content to the body tag
win.AddContent = "<div>Lets Use this div later </div>"
win.create()
'-----------------------------------------------------------------------------------------------------------------------------------------------
Class Window
'@description: Create a custom window with MSHTA.
  Private title, style, body, options, width, height, xpos, ypos  
  Private Sub Class_Initialize()
    title = "&nbsp;" : width = 350 : height = 250
    xpos = "(screen.width - " & width & ")"
    ypos = "-(screen.height -" & height & ")/"& height
    style = "html{display:table;}body{display:table-cell;font-family:Arial;background-color:#30f8ff;}html,body{width:100%;height:100%;margin:0;}"
  End Sub  
  Public Property Let SetTitle(str)     : title = str         : End Property   
  Public Property Let SetWidth(num)     : width = num         : End Property 
  Public Property Let SetHeight(num)    : height = num        : End Property
  Public Property Let SetXPosition(num) : xpos = num          : End Property 
  Public Property Let SetYPosition(num) : ypos = num          : End Property 
  Public Property Let AddStyles(css)    : style = style & css : End Property 
  Public Property Let AddContent(html)  : body = body & html  : End Property    
  Public Property Let ApplicationName(str)    : options = options & "applicationName='" & str & "' "    : End Property
  Public Property Let Border(thick_thin_none) : options = options & "border='" & thick_thin_none & "' " : End Property
  Public Property Let Caption(yes_no)         : options = options & "caption='" & yes_no & "' "         : End Property
  Public Property Let ContextMenu(yes_no)     : options = options & "contextMenu='" & yes_no & "' "     : End Property
  Public Property Let Icon(path)              : options = options & "icon='" & path & "' "              : End Property
  Public Property Let MaximizeButton(yes_no)  : options = options & "maximizeButton='" & yes_no & "' "  : End Property
  Public Property Let MinimizeButton(yes_no)  : options = options & "minimizeButton='" & yes_no & "' "  : End Property
  Public Property Let Scroll(yes_no)          : options = options & "scroll='" & yes_no & "' "          : End Property
  Public Property Let Selection(yes_no)       : options = options & "selection='" & yes_no & "' "       : End Property
  Public Property Let ShowInTaskBar(yes_no)   : options = options & "showInTaskBar='" & yes_no & "' "   : End Property
  Public Property Let SingleInstance(yes_no)  : options = options & "singleInstance='" & yes_no & "' "  : End Property
  Public Property Let SysMenu(yes_no)         : options = options & "sysMenu='" & yes_no & "' "         : End Property
  Public Property Let WindowState(normal_minimize_maximize) : options = options & "windowState='" & normal_minimize_maximize & "' " : End Property  
  Public Function Create()
    Create = CreateObject("WScript.Shell").Exec( _
      "mshta ""about:<!DOCTYPE html><html><head><meta http-equiv='X-UA-Compatible' content='IE=9'><title>" & title & _
      "</title><hta:application " & options & "/><style>" & style & "</style>" & _
      "<script>var c=true;function send(s){c=false;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(s);close();}" & _
      "window.onbeforeunload=function(){if(c)send(0);};resizeTo(" & width & "," & height & ");moveTo(" & xpos & "," & ypos & ");</script>" & _
      "</head><body>" & Age("16-02-0019") & "<p>College :</p>" & College("05-03-006") & "</body></html>""" _
    ).StdOut.ReadLine()
  End Function 
Function Age(DateOfBirth)
    Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
    CurrentDate = CDate(DateOfBirth)
    Years = DateDiff("yyyy", CurrentDate, Date)
    ThisYear = DateAdd("yyyy", Years, CurrentDate)
    Months = DateDiff("m", ThisYear, Date)
    ThisMonth = DateAdd("m", Months, ThisYear)
    Days = DateDiff("d", ThisMonth, Date)

    Do While (Days < 0) Or (Months < 0)
        If Days < 0 Then
            Months = Months - 1
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
        If Months < 0 Then
            Years = Years - 1
            ThisYear = DateAdd("yyyy", Years, CurrentDate)
            Months = DateDiff("m", ThisYear, Date)
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
    Loop
    Age = Years & " years " & Months & " months " & Days+1 & " days"
End Function
Function College(BirthOfLeave)
    Dim CurrentDate, Years, ThisYear, Months, ThisMonth, Days
    CurrentDate = CDate(BirthOfLeave)
    Years = DateDiff("yyyy", CurrentDate, Date)
    ThisYear = DateAdd("yyyy", Years, CurrentDate)
    Months = DateDiff("m", ThisYear, Date)
    ThisMonth = DateAdd("m", Months, ThisYear)
    Days = DateDiff("d", ThisMonth, Date)

    Do While (Days < 0) Or (Months < 0)
        If Days < 0 Then
            Months = Months - 1
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
        If Months < 0 Then
            Years = Years - 1
            ThisYear = DateAdd("yyyy", Years, CurrentDate)
            Months = DateDiff("m", ThisYear, Date)
            ThisMonth = DateAdd("m", Months, ThisYear)
            Days = DateDiff("d", ThisMonth, Date)
        End If
    Loop
    College = Years & " years " & Months & " months " & Days+1 & " days"
End Function
End Class