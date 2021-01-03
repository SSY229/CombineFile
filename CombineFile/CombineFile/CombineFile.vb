Imports System.IO

Module CombineFile
#Region "Declaration"
    Dim ModuleDeveloperLog As String = My.Settings.ModuleDeveloperLog 'Log file path
    Dim WatchFolder As String = My.Settings.WatchFolder 'Log file path
    Dim DropFolder As String = My.Settings.DropFolder 'Log file path
    Dim NumberOfDays As String = My.Settings.NumberOfDays
    Dim ArchiveFolder As String = My.Settings.ArchiveFolder

#End Region

    Sub Main()
        Try
            If Directory.Exists(WatchFolder) AndAlso Directory.Exists(DropFolder) Then
                Message("Information", "Start Combine File Process")
                Dim FileObjList As New List(Of CombineFileObj)
                Dim FileToBeArchived As New List(Of CombineFileObj)
                Dim oTimeSpan As New System.TimeSpan(NumberOfDays - 1, 0, 0, 0)
                Dim StartPoint As Date = DateTime.Now.Subtract(oTimeSpan)
                Dim EndPoint As Date = DateTime.Now
                Dim CurrentDate As Date = StartPoint
                Dim Archived As Boolean = False

                Do While (CurrentDate <= EndPoint)
                    FileObjList = GetCurrentDateFile(WatchFolder, "*" & CurrentDate.ToString("yy") & CurrentDate.ToString("MM") & CurrentDate.ToString("dd") & "?.dat")
                    For Each FileToBeCopied In FileObjList
                        FileToBeArchived.Add(FileToBeCopied)
                    Next
                    CurrentDate = CurrentDate.AddDays(1)
                Loop
                Dim DistinctFilePrefix = (From f In FileToBeArchived Select Path.GetFileNameWithoutExtension(f.File).Substring(0, 3)).Distinct()

                For Each Prefix In DistinctFilePrefix
                    Dim FileObjList2 As New List(Of CombineFileObj)
                    Dim FileToBeProcessed As New List(Of CombineFileObj)
                    For Each File2 In FileToBeArchived
                        Dim FileWithoutExtension As String = Path.GetFileNameWithoutExtension(File2.File)
                        If FileWithoutExtension.Contains(Prefix) Then
                            FileObjList2.Add(File2)
                        End If
                    Next

                    CurrentDate = StartPoint
                    Do While (CurrentDate <= EndPoint)
                        Dim LatestVersionList As New List(Of CombineFileObj)
                        For Each File2 In FileObjList2
                            Dim FileWithoutExtension As String = Path.GetFileNameWithoutExtension(File2.File)
                            If FileWithoutExtension.Contains(CurrentDate.ToString("yy") & CurrentDate.ToString("MM") & CurrentDate.ToString("dd")) Then
                                LatestVersionList.Add(File2)
                            End If
                        Next
                        Dim LatestVersionFile = GetLatestVersion(LatestVersionList)
                        If Not IsNothing(LatestVersionFile) Then
                            Message("Information", "Add " & LatestVersionFile.File & " to the list.")
                            FileToBeProcessed.Add(LatestVersionFile)
                        End If
                        CurrentDate = CurrentDate.AddDays(1)
                    Loop

                    If FileToBeProcessed.Count > 0 Then
                        Dim fileName As String = Path.GetFileNameWithoutExtension(FileToBeProcessed(0).File)
                        Dim DropFolderPath As String = DropFolder & "\" & "CombineFile_" & fileName & ".DAT"
                        Dim NextStep As Boolean = False
                        Message("Information", "Start combining " & FileToBeProcessed.Count & " files into one DAT file(" & DropFolderPath & ").")

                        If File.Exists(DropFolderPath) Then
                            If FileAccessCheck(DropFolderPath) Then
                                File.Delete(DropFolderPath)
                                NextStep = True
                            Else
                                Message("Error", "Failed to remove existing file: " & DropFolderPath)
                            End If
                        Else
                            NextStep = True
                        End If

                        If NextStep Then
                            For Each DATFile In FileToBeProcessed
                                Dim fileReader As String = ""
                                fileReader = My.Computer.FileSystem.ReadAllText(DATFile.File.ToString)
                                ' Write the text.
                                File.AppendAllText(DropFolderPath, fileReader)
                            Next
                            Archived = True
                        End If
                        Message("Information", "End Combine File Process")
                    Else
                        Message("Information", "Not found any today file.")
                    End If
                Next
                If Archived Then
                    For Each ToBeArchivedFile In FileToBeArchived
                        Dim ArchivedPath As String = Path.Combine(ArchiveFolder, Path.GetFileName(ToBeArchivedFile.File))
                        File.Copy(ToBeArchivedFile.File, ArchivedPath, True)
                        If FileAccessCheck(ToBeArchivedFile.File) Then
                            File.Delete(ToBeArchivedFile.File)
                        End If
                    Next
                End If
            Else
                Message("Error", "Unable to locate watch folder / drop folder.")
            End If

            DeleteOldLog()

        Catch ex As Exception
            Message("Exception Error", "Main Function() " & ex.Message)
        End Try
    End Sub
#Region "Misc Function"
    Private Sub Message(ByVal level As String, ByVal msg As String)
        Console.WriteLine(level & ": " & msg) 'Developer Log
        WriteToLog(level & ": " & msg) 'Developer Log
    End Sub

    Public Sub WriteToLog(ByVal new_text As String)
        If Directory.Exists(ModuleDeveloperLog) Then
            Dim filename As String = ModuleDeveloperLog & "\" & "CombineFile_" & DateTime.Now.ToString("yyyy_MM_dd") & ".txt"
            Dim text = DateTime.Now.ToString("HH:mm:ss.fff:") & vbTab & new_text & vbCrLf
            ' Write the text.
            File.AppendAllText(filename, text)
        End If
    End Sub

    Public Sub DeleteOldLog()
        If Directory.Exists(ModuleDeveloperLog) Then
            For Each logFilePath As String In Directory.GetFiles(ModuleDeveloperLog)
                If FileAccessCheck(logFilePath) AndAlso Path.GetFileName(logFilePath).ToLower.Contains(".txt") AndAlso Path.GetFileName(logFilePath).ToLower.Contains("combinefile") Then
                    Dim fileLastWrite As DateTime = File.GetLastWriteTime(logFilePath)
                    If DateDiff("d", fileLastWrite.Date, Now.Date) > 31 Then File.Delete(logFilePath)
                End If
            Next
        End If
    End Sub

    Private Function FileAccessCheck(ByVal filePath As String) As Boolean
        Try
            Using inputstreamreader As New StreamReader(filePath)
                inputstreamreader.Close()
            End Using
            Using inputStream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                inputStream.Close()
                Return True
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function GetCurrentDateFile(ByVal WatchFolder As String, SearchPattern As String) As List(Of CombineFileObj)
        Dim FileObjList As New List(Of CombineFileObj)
        For Each File As String In Directory.GetFiles(WatchFolder, SearchPattern)
            Dim FileObj As New CombineFileObj
            Dim FileSize As Long = FileLen(File)
            If FileSize > 0 Then
                FileObj.File = File
            End If
            FileObjList.Add(FileObj)
        Next

        Return FileObjList
    End Function

    Private Function GetLatestVersion(ByVal FileObjList As List(Of CombineFileObj)) As CombineFileObj
        Dim latestVersion = (From F In FileObjList
                             Let Version = Path.GetFileNameWithoutExtension(F.File).Substring(10, 1)
                             Order By Version Descending
                             Select F).FirstOrDefault()
        Return latestVersion
    End Function
#End Region
End Module
