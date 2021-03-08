Option Strict On
Option Explicit On

Imports System.Configuration
Imports System.Collections.Specialized
Imports System.IO

Public Class CareHubFIleTidy

    Private mstrImportFilesLocation As String = ""
    Private mstrExportFilesLocation As String = ""

    Private IntervalTimer As System.Threading.Timer
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        ' Start the timer
        Dim tsInterval As TimeSpan = New TimeSpan(0, 0, 5)
        IntervalTimer = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf IntervalTimer_Elapsed), Nothing, tsInterval, tsInterval)

        Try

            ' Event log that the service is starting
            Dim myLog As New EventLog()
            If Not myLog.SourceExists("CareHubFileArchive") Then
                myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
            End If

            myLog.Source = "CareHubFileArchive"


            myLog.WriteEntry("CareHub File Archive Log", "Service Started on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Information)


            ' Read the values from the app.config
            mstrImportFilesLocation = ConfigurationManager.AppSettings.Get("ImportFileLocation")
            mstrExportFilesLocation = ConfigurationManager.AppSettings.Get("ExportFileLocation")


            ' Event log the values that are read

        Catch ex As Exception
            Dim myLog As New EventLog()
            If Not myLog.SourceExists("CareHubFileArchive") Then
                myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
            End If

            myLog.Source = "CareHubFileArchive"


            myLog.WriteEntry("CareHub File Archive Log", "Error occurred on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay) & vbCrLf & vbCrLf & ex.Message.ToString,
                    EventLogEntryType.Error)

        End Try

    End Sub

    Protected Overrides Sub OnStop()


        ' Disable the timer
        IntervalTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)
        IntervalTimer.Dispose()
        IntervalTimer = Nothing

        ' Event log that the service is shutting down.
        Try
            Dim myLog As New EventLog()
            If Not myLog.SourceExists("CareHubFileArchive") Then
                myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
            End If

            myLog.Source = "CareHubFileArchive"


            myLog.WriteEntry("CareHub File Archive Log", "Service Stopped on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay),
                    EventLogEntryType.Information)

        Catch ex As Exception

        End Try



    End Sub

    Private Sub IntervalTimer_Elapsed(ByVal state As Object)


        Try



            Dim colLocations As New Collection

            If mstrExportFilesLocation.Trim <> "" Then
                colLocations.Add(mstrExportFilesLocation)
            End If

            If mstrImportFilesLocation.Trim <> "" Then
                colLocations.Add(mstrImportFilesLocation)
            End If

            For intLoopCounter As Integer = 1 To colLocations.Count




                Dim intFilesMoved As Integer = 0

                Dim folderName As String = colLocations(intLoopCounter).ToString
                If Strings.Right(folderName, 1) <> "\" Then folderName += "\"

                'Dim myLog As New EventLog()
                'If Not myLog.SourceExists("CareHubFileArchive") Then
                '    myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
                'End If

                'myLog.Source = "CareHubFileArchive"
                'myLog.WriteEntry("CareHub File Archive Log", "Folder: " & folderName,
                '                EventLogEntryType.Information)


                Dim di As New DirectoryInfo(folderName)

                ' Get a reference to each file in that directory.
                Dim fiArr As FileInfo() = di.GetFiles()

                ' Display the names of the files.
                Dim fri As FileInfo

                If fiArr.Count > 0 Then
                    ' Event log that we are archiving the folder
                    Dim myLog As New EventLog()
                    'myLog = New EventLog()
                    If Not myLog.SourceExists("CareHubFileArchive") Then
                        myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
                    End If

                    myLog.Source = "CareHubFileArchive"
                    myLog.WriteEntry("CareHub File Archive Log", "Archiving Folder " & folderName,
                                EventLogEntryType.Information)

                End If

                For Each fri In fiArr

                    ' Get first six chars of filename
                    Dim branchCode As String = Strings.Left(fri.Name, 6)
                    If Strings.Left(branchCode.ToUpper, 3) = "EXP" Or Strings.Left(branchCode.ToUpper, 3) = "IMP" Then

                        intFilesMoved += 1

                        branchCode = branchCode.ToUpper
                        branchCode = branchCode.Replace("EXP", "")
                        branchCode = branchCode.Replace("IMP", "")

                        ' Check if the folder exists and if not create one.
                        Dim newFolderPath As String = folderName & branchCode
                        If (Not Directory.Exists(newFolderPath)) Then
                            System.IO.Directory.CreateDirectory(newFolderPath)
                        End If

                        ' Get the Year, Month and Date of the file name and create a folder.
                        ' E.g. 2020\Aug\01\
                        'Dim fileYear As String = fri.CreationTime.ToString
                        Dim creationDate As Date = fri.CreationTime

                        Dim fileYear As String = creationDate.Year.ToString
                        Dim fileMonth As String = creationDate.ToString("MMM")
                        Dim fileDay As String = creationDate.ToString("dd")

                        newFolderPath += "\" & fileYear & "\" & fileMonth & "\" & fileDay
                        If (Not Directory.Exists(newFolderPath)) Then
                            System.IO.Directory.CreateDirectory(newFolderPath)
                        End If

                        ' Move the file to the folder

                        If System.IO.File.Exists(newFolderPath & "\" & fri.Name) Then
                            'The file exists rename
                            Dim strNewFileName As String = Replace(fri.Name, fri.Extension, "") & "_1" & fri.Extension


                            Dim myLog As New EventLog()
                            'myLog = New EventLog()
                            If Not myLog.SourceExists("CareHubFileArchive") Then
                                myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
                            End If

                            myLog.Source = "CareHubFileArchive"
                            myLog.WriteEntry("CareHub File Archive Log", "Renaming file to : " & strNewFileName,
                                EventLogEntryType.Information)



                            fri.MoveTo(newFolderPath & "\" & strNewFileName)
                        Else
                            'the file doesn't exist
                            fri.MoveTo(newFolderPath & "\" & fri.Name)
                        End If


                    End If




                Next fri

                ' Event log the number of files archived
                If intFilesMoved > 0 Then


                    Dim myLog As New EventLog()
                    'myLog = New EventLog()
                    If Not myLog.SourceExists("CareHubFileArchive") Then
                        myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
                    End If

                    myLog.Source = "CareHubFileArchive"
                    myLog.WriteEntry("CareHub File Archive Log", "Number of files moved to Folder: " & folderName & " = " & intFilesMoved.ToString,
                                EventLogEntryType.Information)
                End If
            Next

        Catch ex As Exception
            Dim myLog As New EventLog()
            If Not myLog.SourceExists("CareHubFileArchive") Then
                myLog.CreateEventSource("CareHubFileArchive", "CareHub File Archive Log")
            End If

            myLog.Source = "CareHubFileArchive"


            myLog.WriteEntry("CareHub File Archive Log", "Error occurred on  " &
                    Date.Today.ToShortDateString & " " &
                    CStr(TimeOfDay) & vbCrLf & vbCrLf & ex.Message.ToString,
                    EventLogEntryType.Error)

        End Try

    End Sub


End Class
