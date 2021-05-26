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
            '---------------------------------------------------------
            ' This sample enumerates through the tasks on the local computer and
            ' displays their name and state.
            '---------------------------------------------------------


            ' Create the TaskService object.
            Set serviceRoot = CreateObject("Schedule.Service")
            call serviceRoot.Connect()

            ' Get the task folder that contains the tasks. 
            Dim rootFolders
            Set rootFolders = serviceRoot.GetFolder("\")

            Dim taskCollection
            Set taskCollection = rootFolders.GetTasks(0)

            Dim numberOfTasks
            numberOfTasks = taskCollection.Count

            If numberOfTasks = 0 Then 
                Wscript.Echo "No tasks are registered."
            Else
                WScript.Echo "Number of tasks registered: " & numberOfTasks
                Dim flag
                Dim registeredTask
                For Each registeredTask In taskCollection
                    If registeredTask.Name = "Test Daily Trigger" then
                        WScript.Echo "Task Name: " & registeredTask.Name 
                        Dim taskState 
                        Select Case registeredTask.State 
                            Case "0"
                                taskState = "Unknown"
                            Case "1"
                                taskState = "Disabled"
                            Case "2"
                                taskState = "Queued"
                            Case "3"
                                taskState = "Ready"
                            Case "4"
                                taskState = "Running"
                        End Select

                        WScript.Echo "    Task State: " & taskState
                        flag = false
                    Else
                        flag = true
                    End If
                Next
            End If
            If flag = true then
                          '------------------------------------------------------------------
                            ' This sample schedules a task to start on a daily basis.
                            '------------------------------------------------------------------

                            ' A constant that specifies a daily trigger.
                            const TriggerTypeDaily = 2
                            ' A constant that specifies an executable action.
                            const ActionTypeExec = 0

                            '********************************************************
                            ' Create the TaskService object.
                            Set service = CreateObject("Schedule.Service")
                            call service.Connect()

                            '********************************************************
                            ' Get a folder to create a task definition in. 
                            Dim rootFolder
                            Set rootFolder = service.GetFolder("\")

                            ' The taskDefinition variable is the TaskDefinition object.
                            Dim taskDefinition
                            ' The flags parameter is 0 because it is not supported.
                            Set taskDefinition = service.NewTask(0) 

                            '********************************************************
                            ' Define information about the task.

                            ' Set the registration info for the task by 
                            ' creating the RegistrationInfo object.
                            Dim regInfo
                            Set regInfo = taskDefinition.RegistrationInfo
                            regInfo.Description = "Start notepad at 8:00AM daily"
                            regInfo.Author = "Administrator"

                            ' Set the task setting info for the Task Scheduler by
                            ' creating a TaskSettings object.
                            Dim settings
                            Set settings = taskDefinition.Settings
                            settings.Enabled = True
                            settings.StartWhenAvailable = True
                            settings.Hidden = False

                            '********************************************************
                            ' Create a daily trigger. Note that the start boundary 
                            ' specifies the time of day that the task starts and the 
                            ' interval specifies what days the task is run.
                            Dim triggers
                            Set triggers = taskDefinition.Triggers

                            Dim trigger
                            Set trigger = triggers.Create(TriggerTypeDaily)

                            ' Trigger variables that define when the trigger is active 
                            ' and the time of day that the task is run. The format of 
                            ' this time is YYYY-MM-DDTHH:MM:SS
                            Dim startTime, endTime

                            Dim time
                            startTime = "2021-05-25T08:00:00"  'Task runs at 8:00 AM
                            endTime = "2023-05-18T08:00:00"

                            WScript.Echo "startTime :" & startTime
                            WScript.Echo "endTime :" & endTime

                            trigger.StartBoundary = startTime
                            trigger.EndBoundary = endTime
                            trigger.DaysInterval = 1    'Task runs every day.
                            trigger.Id = "DailyTriggerId"
                            trigger.Enabled = True

                            ' Set the task repetition pattern for the task.
                            ' This will repeat the task 5 times.
                            'Dim repetitionPattern
                            'Set repetitionPattern = trigger.Repetition
                            'repetitionPattern.Duration = "PT1M"
                            'repetitionPattern.Interval = "PT1M"

                            '***********************************************************
                            ' Create the action for the task to execute.

                            ' Add an action to the task to run notepad.exe.
                            Dim Action
                            Set Action = taskDefinition.Actions.Create( ActionTypeExec )
                            Action.Path = replace(Split(objProcess.CommandLine, " """)(1), chr(34), "")

                            WScript.Echo "Task definition created. About to submit the task..."

                            '***********************************************************
                            ' Register (create) the task.

                            call rootFolder.RegisterTaskDefinition( _
                                "Test Daily Trigger", taskDefinition, 6, , , 3)

                            WScript.Echo "Task submitted."
            End If
    End If
Next
