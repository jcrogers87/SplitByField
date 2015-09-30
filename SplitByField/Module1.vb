'\\Plataine 2015
'\\searches a CSV for a user specified column inside it and creates csv files with the separated output. 
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Module Module1
    'readconfig globals
    Dim inputJobSheet() As String, outputFolder As String, fieldInt As Integer, headers As Boolean, folderSplit As Boolean, globalMaxjobs As Integer
    'output files globals
    Dim JobsFilename As String, tempDir As String
    'Arrays from input file
    Dim columnHeaders() As String, fieldRows() As String
    Dim columnCount As Integer
    Sub Main()
        Try
            ReadConfig()

            'test to see if we have write access
            tempDir = "C:\ProgramData\Plataine\temp\"
            If (Not Directory.Exists(tempDir)) Then
                Directory.CreateDirectory(tempDir)
            End If

            CleanUp(tempDir)
            ReadInput()
            SplitFiles()
            CleanUp(tempDir)

        Catch ex As Exception
            CleanUp(tempDir)
            MsgBox("An exception with SplitByField has occurred. Contact Plataine" & Chr(13) & Chr(13) & ex.ToString)
            Exit Sub
        End Try
    End Sub
    Public Sub ReadInput()
        For Each element As String In inputJobSheet
            Dim SR As StreamReader = New StreamReader(element)
            Dim line As String = SR.ReadLine()
            Dim strArray As String() = line.Split(",")
            Dim data As DataTable = New DataTable()
            Dim row As DataRow

            For Each s As String In strArray
                data.Columns.Add(New DataColumn())
            Next

            Do
                If Not line = String.Empty Then
                    row = data.NewRow()
                    row.ItemArray = line.Split(",")
                    data.Rows.Add(row)
                Else
                    Exit Do
                End If
                line = SR.ReadLine
            Loop

            Dim i As Integer, x As Integer, j As Integer, tempHeaders As Boolean = headers
            If tempHeaders = True Then
                'For i = 0 To data.Columns.Count - 1
                ' ReDim Preserve columnHeaders(x)
                ' columnHeaders(i) = data.Rows(0).Item(i).ToString
                ' x = x + 1
                ' Next
                data.Rows(0).Delete()
                GetHeaders(element)
            Else
                columnHeaders = Nothing
            End If

            Dim view As New DataView(data)
            view.Sort = data.Columns(fieldInt).ToString
            data = view.ToTable()

            ' Write the row data from the datatable to the new csv file
            Dim count As Integer = 1, fileName As String, split As Integer
            Dim rowData() As String
            For i = 0 To data.Rows.Count - 1
                x = 0
                For j = 0 To data.Columns.Count - 1
                    ReDim Preserve rowData(x)
                    rowData(x) = data.Rows(i).Item(j).ToString
                    x = x + 1
                Next
                fileName = Path.GetFileNameWithoutExtension(element) & "_" & rowData(fieldInt - 1)
                Writer(rowData, fileName, tempDir)
            Next
            SR.Close()
        Next
    End Sub
    Public Sub GetHeaders(input As String)
        If Path.GetExtension(input) = ".csv" Then
            Dim headerString() As String = File.ReadAllLines(input)
            columnHeaders = Split(headerString(0), ",")
        End If
    End Sub
    Public Sub Writer(line() As String, filename As String, outDirectory As String)
        If (Not Directory.Exists(outDirectory)) Then
            Directory.CreateDirectory(outDirectory)
        End If
        JobsFilename = outDirectory & filename & ".csv"
        Dim sw As New StreamWriter(JobsFilename, True)
        sw.WriteLine(String.Join(",", line))
        sw.Close()
    End Sub
    Public Sub SplitFiles()
        For Each element As String In Directory.GetFiles(tempDir)
            If Path.GetExtension(element) = ".xlsx" Or Path.GetExtension(element) = ".xls" Or Path.GetExtension(element) = ".csv" Then
                Dim SR As StreamReader = New StreamReader(element)
                Dim line As String = SR.ReadLine()
                Dim strArray As String() = line.Split(",")
                Dim data As DataTable = New DataTable()
                Dim row As DataRow

                For Each s As String In strArray
                    data.Columns.Add(New DataColumn())
                Next

                Do
                    If Not line = String.Empty Then
                        row = data.NewRow()
                        row.ItemArray = line.Split(",")
                        data.Rows.Add(row)
                    Else
                        Exit Do
                    End If
                    line = SR.ReadLine
                Loop

                Dim place As String = Nothing, filename As String, tempheaders As Boolean
                tempheaders = headers

                Dim maxJobs As Integer = globalMaxjobs
                If maxJobs = Nothing Or 0 Then maxJobs = data.Rows.Count
                If data.Rows.Count / maxJobs > CInt(data.Rows.Count / maxJobs) And data.Rows.Count > maxJobs Then
                    maxJobs = CInt(Math.Ceiling(data.Rows.Count / Math.Ceiling(data.Rows.Count / maxJobs)))
                End If

                Dim rowData() As String, count As Integer = 0, iteration As Integer = 0, x As Integer
                For i = 0 To data.Rows.Count - 1
                    x = 0
                    For j = 0 To data.Columns.Count - 1
                        ReDim Preserve rowData(x)
                        rowData(x) = data.Rows(i).Item(j).ToString
                        x = x + 1
                    Next
                    If folderSplit = True Then
                        place = rowData(fieldInt - 1) & "\"
                    End If
                    filename = Path.GetFileNameWithoutExtension(element) & "-" & iteration + 1
                    If count < maxJobs Then
                        If tempheaders = True And Not File.Exists(filename) Then
                            Writer(columnHeaders, filename, outputFolder & place)
                            tempheaders = False
                        End If
                        Writer(rowData, filename, outputFolder & place)
                        count = count + 1
                    Else
                        tempheaders = True
                        count = 0
                        iteration = iteration + 1
                        i = i - 1
                    End If
                Next
                SR.Close()
            End If
        Next
    End Sub
    Public Sub CleanUp(path As String)
        If Directory.Exists(path) Then
            For Each filepath As String In Directory.GetFiles(tempDir)
                File.Delete(filepath)
            Next
            For Each Dir As String In Directory.GetDirectories(tempDir)
                CleanUp(Dir)
            Next
            Directory.Delete(path, True)
        End If
    End Sub
    Public Sub ReadConfig()
        If Not File.Exists("C:\ProgramData\Plataine\SplitByField.config") Then
            BuildConfig()
        End If
        globalMaxjobs = Nothing
        Try
            headers = False
            Dim configFile() As String = File.ReadAllLines("C:\ProgramData\Plataine\SplitByField.config")
            For Each line As String In configFile
                Dim setting() As String = Split(line, "=")
                If UCase(setting(0)) = "INPUTFOLDER" Then
                    inputJobSheet = Directory.GetFiles(setting(1))
                ElseIf UCase(setting(0)) = "OUTPUTFOLDER" Then
                    outputFolder = setting(1).ToString
                    If Not Right(outputFolder, 1) = "\" Then outputFolder = outputFolder & "\"
                    'column mappings:
                ElseIf UCase(setting(0)) = "HEADERS" Then
                    If UCase(setting(1).ToString) = "TRUE" Then headers = True Else headers = False
                ElseIf UCase(setting(0)) = "FIELD" Then
                    fieldInt = CInt(setting(1).ToString)
                ElseIf UCase(setting(0)) = "SPLITTOFOLDERS" Then
                    If UCase(setting(1).ToString) = "TRUE" Then folderSplit = True Else folderSplit = False
                ElseIf UCase(setting(0)) = "MAXJOBS" Then
                    globalMaxjobs = (setting(1).ToString)
                End If
            Next
            If IsNothing(inputJobSheet) Or IsNothing(outputFolder) Or IsNothing(fieldInt) Then
                Call MsgBox("Your config file is invalid. Must be of the form:" _
                           & Chr(13) & "inputfile=pathtojobs\" _
                           & Chr(13) & "outputFolder=path" _
                           & Chr(13) & "Config location must be: C:\ProgramData\Plataine\SplitByField.config")
                End
            End If
        Catch ex As Exception
            Call MsgBox("Your config file is missing, or missing required column mappings." _
                               & Chr(13) & "Config location must be: C:\ProgramData\Plataine\SplitByField.config")
            End
        End Try
    End Sub
    Public Sub BuildConfig()
        If (Not Directory.Exists("C:\ProgramData\Plataine\")) Then
            Directory.CreateDirectory("C:\ProgramData\Plataine\")
        End If
        Dim sw As New StreamWriter("C:\ProgramData\Plataine\SplitByField.config", False)
        sw.WriteLine("##Auto Config##")
        sw.WriteLine("InputFolder=C:\InputFiles")
        sw.WriteLine("OutputFolder=C:\OutputFiles")
        sw.WriteLine("##Column Mappings")
        sw.WriteLine("field=10")
        sw.WriteLine("headers=true")
        sw.WriteLine("splittofolders=false")
        sw.WriteLine("maxjobs=500")
        sw.Close()
    End Sub
End Module

