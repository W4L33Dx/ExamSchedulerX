Imports System.ComponentModel
Imports System.Globalization
Imports System.IO

Public Class ExamSchedulerX

#Region "TSV HELPERS"

    Sub ReadTSV(ByVal Datagrid As DataGridView, ByVal FilePath As String)
        Dim TextLine As String
        Dim SplitLine() As String

        If System.IO.File.Exists(FilePath) = True Then
            Dim objReader As New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
            Do While objReader.Peek() <> -1
                TextLine = objReader.ReadLine()
                SplitLine = Split(TextLine, ControlChars.Tab)
                Datagrid.Rows.Add(SplitLine)
            Loop
            objReader.Close()
        Else
            'MsgBox("File Does Not Exist")
        End If

    End Sub

    Public Sub WriteDataGridViewTSV(ByVal Datagrid As DataGridView, ByVal FilePath As String)
        'Build the TSV file data as a Tab separated string.

        Dim tsv As String = String.Empty

        'Adding the Rows
        For Each row As DataGridViewRow In Datagrid.Rows
            For Each cell As DataGridViewCell In row.Cells
                tsv = tsv & CType(cell.Value, String) & ControlChars.Tab 'Add the Data rows
            Next

            'trim the last Tab and Add new line
            tsv = tsv.TrimEnd(ControlChars.Tab) & vbCrLf

        Next

        If tsv.Length > 4 Then  'Remove the extra lines at the end of the tsv file
            tsv = tsv.Substring(0, tsv.Length - 4)
        Else
            tsv = ""            'Delete extra row if empty
        End If

        IO.File.WriteAllText(FilePath, tsv)

    End Sub
#End Region

#Region "Schedule Generation"


    ' Recursive function to generate exam schedules
    Sub GenerateExamSchedules(ByVal examCourses As List(Of String),
                              ByVal examsStudentsToCourses As Dictionary(Of String, List(Of String)),
                              ByVal examDurationsToCourses As Dictionary(Of String, Double),
                              ByVal examAvailableDays As List(Of DateTime),
                              ByVal maxExamsPerDayForStudent As Integer,
                              ByRef scheduleNumber As Integer,
                              ByVal schedule As List(Of KeyValuePair(Of String, DateTime)),
                              ByVal howManySchedules As Integer, ByVal startTime As TimeSpan, ByVal endTime As TimeSpan)

        ' Base case: If the schedule is complete (contains all exam courses)
        If schedule.Count = examCourses.Count Then
            ' Print the schedule
            If scheduleNumber <= howManySchedules Then
                PrintSchedule(scheduleNumber, schedule, examsStudentsToCourses, examDurationsToCourses)
                scheduleNumber += 1
                Return
            Else
                Exit Sub
            End If
        End If
        If scheduleNumber >= howManySchedules Then
            Exit Sub
        End If
        ' For each exam course
        For Each course As String In examCourses
            ' Get the list of students for the current course
            Dim courseStudents As List(Of String) = examsStudentsToCourses(course)

            ' Check if the course is already scheduled
            If schedule.Any(Function(e) e.Key = course) Then
                Continue For
            End If

            ' For each available day
            For Each day As DateTime In examAvailableDays

                ' Count the number of exams scheduled on the same day
                Dim examsInSameDay As Integer = schedule.Where(Function(e) e.Value.Date = day.Date).Count

                ' Check if adding the current course exceeds the maximum exams per day for any student
                If examsInSameDay >= maxExamsPerDayForStudent AndAlso
                    courseStudents.Any(Function(student)
                                           Dim tempSchedule = schedule.ToList()
                                           Return tempSchedule.Any(Function(e) e.Value.Date = day.Date AndAlso examsStudentsToCourses(e.Key).Contains(student))
                                       End Function) Then
                    Continue For
                End If




                ' Get the duration of the current exam
                Dim examDuration As Double = examDurationsToCourses(course)

                ' Iterate over possible start times within the available day
                Dim currentTime As TimeSpan = startTime
                While currentTime.Add(TimeSpan.FromHours(examDuration)).CompareTo(endTime) <= 0
                    Dim examTime As DateTime = day.Date.Add(currentTime)


                    Dim conflictingExam As Boolean = False

                    For Each e In schedule
                        If e.Value.Date = day.Date Then
                            Dim examEndTime As DateTime = e.Value.AddHours(examDurationsToCourses(e.Key))
                            Dim newExamStartTime As DateTime = examTime.AddHours(0)
                            If examsStudentsToCourses(e.Key).Intersect(courseStudents).Any() Then
                                newExamStartTime = examTime.AddHours(examDuration)
                            End If

                            If Not (examEndTime <= examTime OrElse newExamStartTime <= e.Value) Then
                                conflictingExam = True
                                Exit For
                            End If
                        End If
                    Next e

                    If conflictingExam Then
                        currentTime = currentTime.Add(TimeSpan.FromHours(1)) ' Increment by 1 hour for the next possible start time
                        Continue While
                    End If


                    ' Schedule the current exam
                    schedule.Add(New KeyValuePair(Of String, DateTime)(course, examTime))

                    ' Recursive call to generate schedules for remaining courses
                    GenerateExamSchedules(examCourses, examsStudentsToCourses, examDurationsToCourses, examAvailableDays,
                                      maxExamsPerDayForStudent, scheduleNumber, schedule, howManySchedules, startTime, endTime)

                    ' Backtrack by removing the last scheduled exam
                    schedule.RemoveAt(schedule.Count - 1)

                    currentTime = currentTime.Add(TimeSpan.FromHours(1)) ' Increment by 1 hour for the next possible start time
                End While
            Next
        Next
    End Sub

    ' Function to check if a student is busy during the given exam schedule
    Function IsStudentBusy(ByVal student As String, ByVal schedule As List(Of KeyValuePair(Of String, DateTime)), ByVal examsStudentsToCourses As Dictionary(Of String, List(Of String))) As Boolean
        For Each exam In schedule
            ' Check if the current exam is taken by the given student
            ' And if the exam time is the same as the last scheduled exam time
            If examsStudentsToCourses(exam.Key).Contains(student) AndAlso exam.Value = schedule.Last().Value Then
                ' If the conditions are met, it means the student is busy with an exam at the given schedule
                Return True
            End If
        Next

        ' If no busy exam is found for the student in the schedule, return False
        Return False
    End Function


    ' Function to print the exam schedule
    Sub PrintSchedule(ByVal scheduleNumber As Integer,
                      ByVal schedule As List(Of KeyValuePair(Of String, DateTime)),
                      ByVal examsStudentsToCourses As Dictionary(Of String, List(Of String)),
                      ByVal examDurationsToCourses As Dictionary(Of String, Double))

        ' Print the schedule number
        RichTextBox1.Text += ($"SCHEDULE NUMBER {scheduleNumber}")

        ' Print the header for the schedule table
        RichTextBox1.Text += vbNewLine & ("COURSE NAME - Date - Start Time - End Time - Students")

        ' Print a line for visual separation
        RichTextBox1.Text += vbNewLine & ("=====================")

        ' Iterate over each exam in the schedule
        For Each exam In schedule
            ' Get the course name from the exam
            Dim courseName = exam.Key

            ' Get the date of the exam in the format "yyyy-MM-dd"
            Dim dateValue = exam.Value.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)

            ' Get the start time of the exam in the format "HH:mm"
            Dim startTime = exam.Value.ToString("HH:mm")

            ' Calculate the end time of the exam by adding the exam duration to the start time,
            ' and format it as "HH:mm"
            Dim endTime = exam.Value.AddHours(examDurationsToCourses(courseName)).ToString("HH:mm")

            ' Get the list of students taking the current course and join them into a string separated by commas
            Dim students = String.Join(", ", examsStudentsToCourses(courseName))

            ' Print the exam details in the format "courseName - dateValue - startTime - endTime - students"
            RichTextBox1.Text += vbNewLine & ($"{courseName} - {dateValue} - {startTime} - {endTime} - {students}")
        Next

        ' Print a line for visual separation
        RichTextBox1.Text += vbNewLine & ("=====================") & vbNewLine

        ' Print an empty line for separation between schedules
        RichTextBox1.Text += vbNewLine
    End Sub


#End Region

    Private Sub ExamSchedulerX_Load(sender As Object, e As EventArgs) Handles Me.Load
        If (File.Exists(Application.StartupPath & "\d1.tsv")) Then
            ReadTSV(DataGridView1, Application.StartupPath & "\d1.tsv")
        End If
        If (File.Exists(Application.StartupPath & "\d2.tsv")) Then
            ReadTSV(DataGridView2, Application.StartupPath & "\d2.tsv")
        End If
    End Sub
    Private Sub FlatButton1_Click(sender As Object, e As EventArgs) Handles FlatButton1.Click
        WriteDataGridViewTSV(DataGridView1, Application.StartupPath & "\d1.tsv")
    End Sub
    Private Sub FlatButton2_Click(sender As Object, e As EventArgs) Handles FlatButton2.Click
        WriteDataGridViewTSV(DataGridView2, Application.StartupPath & "\d2.tsv")
    End Sub

    Private Sub FlatButton3_Click(sender As Object, e As EventArgs) Handles FlatButton3.Click
        RichTextBox1.Text = ""
        Dim examDurationsToCourses As New Dictionary(Of String, Double)

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim course As String = row.Cells(0).Value.ToString()
                Dim duration As Double

                If Double.TryParse(row.Cells(1).Value.ToString(), duration) Then
                    examDurationsToCourses.Add(course, duration)
                End If
            End If
        Next

        Dim examsStudentsToCourses As New Dictionary(Of String, List(Of String))

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim course As String = row.Cells(0).Value.ToString()
                Dim students As List(Of String) = New List(Of String)()

                Dim studentsData As String = row.Cells(2).Value.ToString()
                Dim studentNames As String() = studentsData.Split(","c)

                For Each studentName As String In studentNames
                    students.Add(studentName.Trim())
                Next

                examsStudentsToCourses.Add(course, students)
            End If
        Next

        Dim examAvailableDays As New List(Of DateTime)()

        For Each row As DataGridViewRow In DataGridView2.Rows
            If Not row.IsNewRow Then
                Dim dateStr As String = row.Cells(0).Value.ToString()
                Dim dateParts As String() = dateStr.Split("\"c)

                If dateParts.Length = 3 Then
                    Dim day As Integer
                    Dim month As Integer
                    Dim year As Integer

                    If Integer.TryParse(dateParts(0), day) AndAlso
               Integer.TryParse(dateParts(1), month) AndAlso
               Integer.TryParse(dateParts(2), year) Then

                        Dim examDate As New DateTime(year, month, day)
                        examAvailableDays.Add(examDate)
                    End If
                End If
            End If
        Next
        Dim maxExamsPerDayForStudent As Integer = FlatNumeric1.Value
        Dim startTime As TimeSpan = New TimeSpan(FlatNumeric2.Value, 0, 0)
        Dim endTime As TimeSpan = New TimeSpan(FlatNumeric3.Value, 0, 0)
        Dim howManySchedules As Integer = FlatNumeric4.Value



        Dim examCourses As List(Of String) = examsStudentsToCourses.Keys.ToList()

        Dim schedule As New List(Of KeyValuePair(Of String, DateTime))

        GenerateExamSchedules(examCourses, examsStudentsToCourses, examDurationsToCourses, examAvailableDays, maxExamsPerDayForStudent, 1, schedule, howManySchedules, startTime, endTime)
    End Sub


    Private Sub FlatButton4_Click(sender As Object, e As EventArgs) Handles FlatButton4.Click
        Dim ofd As New OpenFileDialog
        With ofd
            .Filter = "*.tsv|*.tsv"
            .ShowDialog()
        End With
        If (ofd.FileName.Length > 1) Then
            DataGridView1.Rows.Clear()
            ReadTSV(DataGridView1, ofd.FileName)
        End If
    End Sub

    Private Sub FlatButton5_Click(sender As Object, e As EventArgs) Handles FlatButton5.Click
        Dim sfd As New SaveFileDialog
        With sfd
            .Filter = "*.tsv|*.tsv"
            .ShowDialog()
        End With
        If (sfd.FileName.Length > 1) Then
            WriteDataGridViewTSV(DataGridView1, sfd.FileName)
        End If
    End Sub

    Private Sub FlatButton7_Click(sender As Object, e As EventArgs) Handles FlatButton7.Click
        Dim ofd As New OpenFileDialog
        With ofd
            .Filter = "*.tsv|*.tsv"
            .ShowDialog()
        End With
        If (ofd.FileName.Length > 1) Then
            DataGridView2.Rows.Clear()
            ReadTSV(DataGridView2, ofd.FileName)
        End If
    End Sub

    Private Sub FlatButton6_Click(sender As Object, e As EventArgs) Handles FlatButton6.Click
        Dim sfd As New SaveFileDialog
        With sfd
            .Filter = "*.tsv|*.tsv"
            .ShowDialog()
        End With
        If (sfd.FileName.Length > 1) Then
            WriteDataGridViewTSV(DataGridView2, sfd.FileName)
        End If
    End Sub

    Private Sub FlatButton8_Click(sender As Object, e As EventArgs) Handles FlatButton8.Click
        Dim sfd As New SaveFileDialog
        With sfd
            .Filter = "*.txt|*.txt"
            .ShowDialog()
        End With
        If (sfd.FileName.Length > 1) Then
            IO.File.WriteAllText(sfd.FileName, RichTextBox1.Text)
        End If
    End Sub
    Private Sub FlatButton9_Click(sender As Object, e As EventArgs) Handles FlatButton9.Click
        Dim savedb As DialogResult = MsgBox("Do You Want To Save The Data ?", vbYesNo)
        If (savedb = DialogResult.Yes) Then
            WriteDataGridViewTSV(DataGridView2, Application.StartupPath & "\d2.tsv")
            WriteDataGridViewTSV(DataGridView1, Application.StartupPath & "\d1.tsv")
            Application.Exit()
        Else
            Application.Exit()
        End If
    End Sub

    Private FillingTextIndex As Integer = 0
    Private Sub FlatTabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FlatTabControl1.SelectedIndexChanged
        If FlatTabControl1.SelectedIndex = 4 Then
            ' Start filling the RichTextBox with delay
            FillingTextIndex = 0
            RichTextBox2.Text = ""
            Me.Timer1.Interval = 10 ' Adjust the interval as needed
            Me.Timer1.Start()
        Else
            ' Stop filling the RichTextBox
            Me.Timer1.Stop()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If FillingTextIndex < My.Resources.About.Length Then
            ' Append the next character to RichTextBox2.Text
            RichTextBox2.AppendText(My.Resources.About(FillingTextIndex))
            FillingTextIndex += 1
        Else
            ' Stop the timer when the text has been fully filled
            Timer1.Stop()
        End If

    End Sub

End Class