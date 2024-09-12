
Imports System.Data.Odbc
Imports System.Data.OleDb
Imports System.Globalization

Public Class Form1

    Public koneski As OdbcConnection
    Dim PathFile As String
    Dim PasswordMdb As String
    Dim DSNname As String



    Private Sub MainWindow_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mdbsFilePath.Text = "C:\Program Files (x86)\KTT_FPReader\dbFolder\hds_fpsystem.mdb"
        mdbPassword.Text = "tjdgustltmxpa"
        odbsDsnText.Text = "ErpTech"


        If IntervalSelected.Items.Count >= 4 Then
            IntervalSelected.SelectedIndex = 3
        End If
        Button1_Click(sender, e)
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button_mdb_read.Click
        PathFile = mdbsFilePath.Text
        PasswordMdb = mdbPassword.Text
        DSNname = odbsDsnText.Text
        Dim Path As String = mdbsFilePath.Text
        Dim password As String = mdbPassword.Text
        Dim odbcDsn As String = odbsDsnText.Text

        If String.IsNullOrEmpty(PathFile) Then
            MessageBox.Show("MDB 파일 경로 입력해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(PasswordMdb) Then
            MessageBox.Show("MDB 암호키 입력해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf String.IsNullOrEmpty(DSNname) Then
            MessageBox.Show("ODBS DSN 입력해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf IntervalSelected.SelectedIndex = -1 Then
            MessageBox.Show("주기 시간 선택해주세요.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            ConnectToMariadb()
            Dim selectedItem As String = IntervalSelected.SelectedItem.ToString()
            Dim selecteditemInt As Integer
            Dim selectDate As Date = DateTimePicker.Value
            Dim formattedDateSelected As String = selectDate.ToString("yyyy-MM-dd")

            If Integer.TryParse(selectedItem, selecteditemInt) Then
                ' Conversion successful, selectedItemInt now holds the integer value
                Me.Timer1.Interval = TimeSpan.FromSeconds(selecteditemInt).TotalMilliseconds
                Me.Timer1.Start()
                testText.AppendText("주기 시간 " & selecteditemInt.ToString() & "초로 설정되었습니다." + vbCrLf)
            Else
                ' Conversion failed, handle the error or provide a default value
                testText.AppendText("Conversion failed. Default value set to 0.")
                selecteditemInt = 0 ' Set a default value or handle the error as needed
            End If
        End If

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles button_end_work.Click
        Me.Timer1.Stop()
        testText.AppendText("작업 끝났습니다 " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & Environment.NewLine)
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        RetrieveDataAndDisplay()
    End Sub




    Private Sub ConnectToMariadb()
        Dim connectionString As String = "DSN=" & DSNname

        Try
            koneski = New OdbcConnection(connectionString)

            If koneski.State = ConnectionState.Closed Then
                koneski.Open()
                testText.Text = "Connected to MariaDB!" & vbCrLf
            Else
                testText.Text = "Connection is already open." & vbCrLf
            End If
        Catch ex As Exception
            testText.Text = "Error: " & ex.Message
        End Try
    End Sub


    Private Sub RetrieveDataAndDisplay()
        'Maria DB 연결이 종료되었으면 다시 연결 
        If koneski.State = ConnectionState.Closed Then
            ConnectToMariadb()
        End If
        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PathFile & ";Jet OLEDB:Database Password=" & PasswordMdb
        'Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Spring/vbtest.mdb;Jet OLEDB:Database Password=tjdgustltmxpa"
        Dim connection As New OleDbConnection(connectionString)

        Try
            ' Open the connection
            connection.Open()
            testText.AppendText("작업 시작되었습니다 " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & Environment.NewLine)
            testText.ScrollToCaret()


            Dim currentDate As Date = Date.Today
            Dim formattedDate As String = currentDate.ToString("yyyy-MM-dd")
            Dim selectDate As Date = DateTimePicker.Value
            Dim formattedDateSelected As String = selectDate.ToString("yyyy-MM-dd")
            'Dim query As String = "SELECT * FROM tb_workresult WHERE LEFT(date_Attestation, 10) >= #" & formattedDateSelected & "# AND (str_ValidationStatus = '성공') ORDER BY date_Attestation ASC"'

            Dim query As String = "SELECT * FROM tb_workresult WHERE LEFT(date_Attestation, 10) >= #" & formattedDateSelected & "# AND (str_ValidationStatus = '성공') AND ((str_Mode = '출입' AND str_accTerminalPlace IN ('정문','연구1층출입문내부', '테크1층출입문내부', '테크1층차고지내부', '테크지하1층출입문내부', '연구1층출입문외부', '테크1층출입문외부' , '테크1층차고지외부', '테크지하1층출입문외부')) OR (str_Mode IN ('출근', '퇴근'))) ORDER BY date_Attestation ASC"

            Dim command As New OleDbCommand(query, connection)


            ' Execute the query and read the data
            Dim reader As OleDbDataReader = command.ExecuteReader()

            ' Check if there is data to read
            If reader.HasRows Then
                ' Initialize an empty string to hold the concatenated result
                Dim resultText As String = ""

                ' Read the data and append it to the resultText
                While reader.Read()
                    ' Assuming you have a TextBox named testText on your form
                    Dim koreanCulture As New CultureInfo("ko-KR")
                    Dim inputDateTime As DateTime = DateTime.ParseExact(reader("date_Attestation").ToString(), "yyyy-MM-dd tt h:mm:ss", koreanCulture)
                    Dim outputDateString As String = inputDateTime.ToString("yyyy-MM-dd HH:mm:ss")


                    Dim checkQuery As String = "SELECT COUNT(*) FROM TB_WORKER_ATTENDANCE WHERE str_workermpNum = ? AND str_accTerminalPlace = ? AND date_Attestation = ? AND str_Mode = ? AND str_workempName = ?"
                    Dim checkCommand As New OdbcCommand(checkQuery, koneski)
                    checkCommand.Parameters.AddWithValue("@str_workermpNum", Convert.ToInt32(reader("str_workempNum")))
                    checkCommand.Parameters.AddWithValue("@str_accTerminalPlace", reader("str_accTerminalPlace").ToString())
                    checkCommand.Parameters.AddWithValue("@date_Attestation", inputDateTime)
                    checkCommand.Parameters.AddWithValue("@str_Mode", reader("str_Mode").ToString())
                    checkCommand.Parameters.AddWithValue("@str_workempName", reader("str_workempName").ToString())
                    Dim existingRecordsCount As Integer = CInt(checkCommand.ExecuteScalar())

                    Dim originalWorkempNum As String = reader("str_workempNum").ToString()
                    Dim strPlacemnet As String = reader("str_accTerminalPlace").ToString()
                    Dim paddedWorkempNum As String
                    Dim workerNo As String = ""

                    Dim existingRecordsCountWorker As Integer

                    If originalWorkempNum.Length = 3 Then
                        paddedWorkempNum = originalWorkempNum.PadLeft(4, "0")
                    Else
                        paddedWorkempNum = originalWorkempNum
                    End If



                    If strPlacemnet = "정문" Then
                        ' Conditional SQL queries with additional condition for biz_no
                        Dim checkQueryWorker As String = "SELECT COUNT(*) FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%') AND biz_no = '399-81-01591'"
                        Dim selectWorkerNo As String = "SELECT worker_no FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%') AND biz_no = '399-81-01591'"

                        Dim checkCommandWorkerCode As New OdbcCommand(checkQueryWorker, koneski)
                        Dim selectCommandWorkerNumber As New OdbcCommand(selectWorkerNo, koneski)

                        checkCommandWorkerCode.Parameters.AddWithValue("@worker_code", paddedWorkempNum)
                        selectCommandWorkerNumber.Parameters.AddWithValue("@worker_code", paddedWorkempNum)

                        existingRecordsCountWorker = CInt(checkCommandWorkerCode.ExecuteScalar())

                        Dim readerWorker_no As OdbcDataReader = selectCommandWorkerNumber.ExecuteReader()

                        ' Process the readerWorker_no as needed
                        While readerWorker_no.Read()
                            workerNo = readerWorker_no("worker_no").ToString()
                            ' Do something with workerNo
                        End While

                        readerWorker_no.Close()
                    Else
                        ' SQL queries with the condition for biz_no != 399-81-01591
                        Dim checkQueryWorker As String = "SELECT COUNT(*) FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%') AND biz_no != '399-81-01591'"
                        Dim selectWorkerNo As String = "SELECT worker_no FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%') AND biz_no != '399-81-01591'"

                        Dim checkCommandWorkerCode As New OdbcCommand(checkQueryWorker, koneski)
                        Dim selectCommandWorkerNumber As New OdbcCommand(selectWorkerNo, koneski)

                        checkCommandWorkerCode.Parameters.AddWithValue("@worker_code", paddedWorkempNum)
                        selectCommandWorkerNumber.Parameters.AddWithValue("@worker_code", paddedWorkempNum)

                        existingRecordsCountWorker = CInt(checkCommandWorkerCode.ExecuteScalar())

                        Dim readerWorker_no As OdbcDataReader = selectCommandWorkerNumber.ExecuteReader()

                        ' Process the readerWorker_no as needed
                        While readerWorker_no.Read()
                            workerNo = readerWorker_no("worker_no").ToString()
                            ' Do something with workerNo
                        End While

                        readerWorker_no.Close()
                    End If



                    'Dim checkQueryWorker As String = "SELECT COUNT(*) FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%')"
                    'Dim selectWorkerNo As String = "SELECT worker_no FROM TB_WORKER WHERE CONCAT(',', worker_code, ',') LIKE CONCAT('%,', ?, ',%')"
                    'Dim checkCommandWorkerCode As New OdbcCommand(checkQueryWorker, koneski)
                    'Dim selectCommandWorkerNumber As New OdbcCommand(selectWorkerNo, koneski)
                    'checkCommandWorkerCode.Parameters.AddWithValue("@worker_code", paddedWorkempNum)
                    'selectCommandWorkerNumber.Parameters.AddWithValue("@worker_code", paddedWorkempNum)
                    'Dim existingRecordsCountWorker As Integer = CInt(checkCommandWorkerCode.ExecuteScalar())
                    'Dim readerWorker_no As OdbcDataReader = selectCommandWorkerNumber.ExecuteReader()
                    'If readerWorker_no.Read Then
                    'workerNo = readerWorker_no("worker_no").ToString()
                    'End If


                    Dim insertQuery As String = "INSERT INTO TB_WORKER_ATTENDANCE(str_workermpNum, str_accTerminalPlace, date_Attestation, str_Mode, str_workempName) " &
                                    "VALUES (?, ?, ?, ?, ?)"
                    Dim commandInsert As New OdbcCommand(insertQuery, koneski)
                    commandInsert.Parameters.AddWithValue("@str_workermpNum", Convert.ToInt32(reader("str_workempNum")))
                    commandInsert.Parameters.AddWithValue("@str_accTerminalPlace", reader("str_accTerminalPlace").ToString())
                    commandInsert.Parameters.AddWithValue("@date_Attestation", inputDateTime)
                    commandInsert.Parameters.AddWithValue("@str_Mode", reader("str_Mode").ToString())
                    commandInsert.Parameters.AddWithValue("@str_workempName", reader("str_workempName").ToString())



                    ' Insert data only if no matching record exists
                    If existingRecordsCount = 0 AndAlso existingRecordsCountWorker <> 0 Then

                        Dim isOverTimeNight As Boolean = False

                        Dim attendanceDate As String = inputDateTime.ToString("yyyy-MM-dd HH:mm:sss")
                        Dim inputDateTimeParse As DateTime = DateTime.Parse(attendanceDate)
                        If (reader("str_Mode").ToString() = "퇴근" OrElse (reader("str_Mode").ToString() = "출입" AndAlso {"연구1층출입문내부", "테크1층출입문내부", "테크1층차고지내부", "테크지하1층출입문내부"}.Contains(reader("str_accTerminalPlace")))) AndAlso (inputDateTimeParse.Hour >= 0 AndAlso inputDateTimeParse.Hour < 3) Then
                            testText.Text &= "its overtiime." & attendanceDate & vbCrLf
                            isOverTimeNight = True
                        End If



                        Dim checkQueryEntrance As String = "SELECT COUNT(*) FROM TB_WORKER_ENTRANCE WHERE date = ? and worker_no = ?"
                        Dim checkCommandEntrance As New OdbcCommand(checkQueryEntrance, koneski)
                        checkCommandEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                        checkCommandEntrance.Parameters.AddWithValue("@worker_no", workerNo)
                        Dim existingRecordsCountEntrance As Integer = CInt(checkCommandEntrance.ExecuteScalar())


                        Dim insertQueryEntrance As String = ""
                        Dim status As String = "0"
                        Dim startTimeWorker As String = ""
                        Dim time1 As TimeSpan
                        Dim updateStatus As String = ""
                        Dim existingKtGoToWork As String = ""
                        Dim setNewDate As Boolean = False


                        If existingRecordsCountEntrance > 0 Then
                            Dim getStatus As String = "SELECT update_status, kt_go_to_work FROM TB_WORKER_ENTRANCE WHERE date = ? and worker_no = ?"
                            Dim getStatusCommand As New OdbcCommand(getStatus, koneski)
                            getStatusCommand.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                            getStatusCommand.Parameters.AddWithValue("@worker_no", workerNo)
                            Dim readerExistingData As OdbcDataReader = getStatusCommand.ExecuteReader()
                            If readerExistingData.Read Then
                                updateStatus = readerExistingData("update_status").ToString()
                                Dim inputDateTimeExisting As DateTime = DateTime.ParseExact(readerExistingData("kt_go_to_work").ToString(), "yyyy-MM-dd tt h:mm:ss", koreanCulture)
                                existingKtGoToWork = inputDateTimeExisting.ToString("HH:mm:ss")
                                Dim newDate As String = inputDateTime.ToString("HH:mm:ss")
                                Dim timeExistKtGoToWork As TimeSpan = TimeSpan.Parse(existingKtGoToWork)
                                Dim timeNewGoToWork As TimeSpan = TimeSpan.Parse(newDate)
                                If timeExistKtGoToWork > timeNewGoToWork Then
                                    setNewDate = True
                                    testText.Text &= "Data existing data." & vbCrLf
                                End If

                            End If
                        End If



                        If existingRecordsCountEntrance < 1 AndAlso Not isOverTimeNight Then
                            Dim checkWorkStatus As String = "SELECT start_time FROM TB_WORKER_WORK_TIME WHERE worker_no = ? AND status = 1 AND start_date <= ?" '수정 2024-05-16
                            Dim checkCountStatus As String = "SELECT COUNT(*) FROM TB_WORKER_WORK_TIME WHERE worker_no = ? AND status = 1 AND start_date <= ?"  '수정 2024-05-16
                            Dim countGrant As String = "SELECT COUNT(*) FROM techerp_approval.TB_GRANT WHERE worker_no = ? AND start_date <= ? AND end_date >= ? "
                            Dim checkGrantStatus As String = "SELECT attendance_no, am_pm FROM techerp_approval.TB_GRANT WHERE worker_no = ? AND start_date <= ? AND end_date >= ?"

                            Dim checkCommandGrantStatus As New OdbcCommand(checkGrantStatus, koneski)
                            Dim checkGrantCountStatus As New OdbcCommand(countGrant, koneski)
                            checkCommandGrantStatus.Parameters.AddWithValue("@worker_no", workerNo)
                            checkCommandGrantStatus.Parameters.AddWithValue("@start_date", inputDateTime.ToString("yyyy-MM-dd"))
                            checkCommandGrantStatus.Parameters.AddWithValue("@end_date", inputDateTime.ToString("yyyy-MM-dd"))
                            Dim attendanceNo As String = checkCommandGrantStatus.ExecuteScalar()?.ToString()
                            Dim am_pm As String = ""

                            checkGrantCountStatus.Parameters.AddWithValue("@worker_no", workerNo)
                            checkGrantCountStatus.Parameters.AddWithValue("@start_date", inputDateTime.ToString("yyyy-MM-dd"))
                            checkGrantCountStatus.Parameters.AddWithValue("@end_date", inputDateTime.ToString("yyyy-MM-dd"))
                            Dim countGrantStatus As Integer = CInt(checkGrantCountStatus.ExecuteScalar())

                            If countGrantStatus > 0 Then
                                Dim readerGrant As OdbcDataReader = checkCommandGrantStatus.ExecuteReader()
                                If readerGrant.Read Then
                                    am_pm = readerGrant("am_pm").ToString()
                                End If
                            End If


                            Dim checkCommandWorkStatus As New OdbcCommand(checkWorkStatus, koneski)
                            Dim checkCommnadCountStatus As New OdbcCommand(checkCountStatus, koneski)
                            checkCommandWorkStatus.Parameters.AddWithValue("@worker_no", workerNo)
                            checkCommandWorkStatus.Parameters.AddWithValue("@start_date", inputDateTime.ToString("yyyy-MM-dd"))  '추가 2024-05-16
                            checkCommnadCountStatus.Parameters.AddWithValue("@worker_no", workerNo)
                            checkCommnadCountStatus.Parameters.AddWithValue("@start_date", inputDateTime.ToString("yyyy-MM-dd")) '추가 2024-05-16

                            Dim existingRecordsCountWorkStart As Integer = CInt(checkCommnadCountStatus.ExecuteScalar())
                            ' Execute the query and store the result directly in startTimeWorker

                            If existingRecordsCountWorkStart <> 0 Then
                                startTimeWorker = checkCommandWorkStatus.ExecuteScalar()?.ToString()
                                time1 = TimeSpan.Parse(startTimeWorker)
                                time1 = time1.Add(New TimeSpan(0, 0, 59))
                            End If


                            Dim timeFrom As String = inputDateTime.ToString("HH:mm:ss")
                            Dim time2 As TimeSpan = TimeSpan.Parse(timeFrom)
                            Dim time3 As TimeSpan = TimeSpan.Parse("09:00:59")

                            Dim time3Grant As TimeSpan = TimeSpan.Parse("13:30:59")


                            Dim dateString As String = inputDateTime.ToString("yyyy-MM-dd")


                            Dim dateValue As DateTime



                            If DateTime.TryParse(dateString, dateValue) Then
                                If dateValue.DayOfWeek = DayOfWeek.Saturday OrElse dateValue.DayOfWeek = DayOfWeek.Sunday Then   '주말에 경우에 지각 처리 안함
                                    status = "0"
                                Else
                                    If countGrantStatus > 0 Then   '해당일에 신청 내역이 있는지 확인'

                                        If attendanceNo = "1000000010" Then
                                            If am_pm = "0" Then
                                                If existingRecordsCountWorkStart <> 0 Then  '시차출퇴근 내역이 있는지 확인
                                                    If time1 > time2 Then
                                                        status = "0"
                                                    Else
                                                        status = "1"
                                                    End If
                                                Else
                                                    If time2 > time3 Then
                                                        status = "1"
                                                    Else
                                                        status = "0"
                                                    End If
                                                End If
                                            Else
                                                If existingRecordsCountWorkStart <> 0 Then   '시차출퇴근 내역이 있는지 확인
                                                    time1 = time1.Add(New TimeSpan(4, 30, 0))
                                                    If time1 > time2 Then
                                                        status = "0"
                                                    Else
                                                        status = "1"
                                                    End If
                                                Else
                                                    If time2 > time3Grant Then
                                                        status = "1"
                                                    Else
                                                        status = "0"
                                                    End If
                                                End If
                                            End If
                                        Else
                                            status = "3"
                                        End If

                                    Else
                                        If existingRecordsCountWorkStart <> 0 Then  '시차출퇴근 내역이 있는지 확인
                                            If time1 > time2 Then
                                                status = "0"
                                            Else
                                                status = "1"
                                            End If
                                        Else
                                            If time2 > time3 Then
                                                status = "1"
                                            Else
                                                status = "0"
                                            End If
                                        End If
                                    End If

                                End If
                            Else

                            End If
                            insertQueryEntrance = "INSERT INTO TB_WORKER_ENTRANCE(date, worker_code, kt_go_to_work, worker_no, status) VALUES (?,?,?,?,?)"
                        ElseIf (reader("str_Mode").ToString() = "퇴근" OrElse (reader("str_Mode").ToString() = "출입" AndAlso {"연구1층출입문내부", "테크1층출입문내부", "테크1층차고지내부", "테크지하1층출입문내부"}.Contains(reader("str_accTerminalPlace")))) AndAlso (updateStatus <> "3" AndAlso updateStatus <> "2") Then
                            insertQueryEntrance = "UPDATE TB_WORKER_ENTRANCE SET kt_leave_work = ? WHERE date = ?  AND worker_no = ?"
                            testText.Text &= "Updated status equals." & updateStatus & vbCrLf
                        ElseIf (reader("str_Mode").ToString() = "출입" AndAlso {"정문", "연구1층출입문외부", "테크1층출입문외부", "테크1층차고지외부", "테크지하1층출입문외부"}.Contains(reader("str_accTerminalPlace"))) AndAlso Not setNewDate Then
                            insertQueryEntrance = "UPDATE TB_WORKER_ENTRANCE SET kt_go_inside = ? WHERE date = ?  AND worker_no = ?"
                        ElseIf (reader("str_Mode").ToString() = "출근" OrElse (reader("str_Mode").ToString() = "출입" AndAlso {"정문", "연구1층출입문외부", "테크1층출입문외부", "테크1층차고지외부", "테크지하1층출입문외부"}.Contains(reader("str_accTerminalPlace")))) AndAlso setNewDate Then
                            insertQueryEntrance = "UPDATE TB_WORKER_ENTRANCE SET kt_go_to_work = ?, status = ? WHERE date = ?  AND worker_no = ?"
                        End If


                        Dim commandInsertEntrance As New OdbcCommand(insertQueryEntrance, koneski)

                        If existingRecordsCountEntrance < 1 AndAlso Not isOverTimeNight Then
                            commandInsertEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                            commandInsertEntrance.Parameters.AddWithValue("@worker_code", Convert.ToInt32(reader("str_workempNum")))
                            commandInsertEntrance.Parameters.AddWithValue("@kt_go_to_work", inputDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            commandInsertEntrance.Parameters.AddWithValue("@worker_no", workerNo)
                            commandInsertEntrance.Parameters.AddWithValue("@status", status)
                        ElseIf (reader("str_Mode").ToString() = "퇴근" OrElse (reader("str_Mode").ToString() = "출입" AndAlso {"연구1층출입문내부", "테크1층출입문내부", "테크1층차고지내부", "테크지하1층출입문내부"}.Contains(reader("str_accTerminalPlace")))) AndAlso (updateStatus <> "3" AndAlso updateStatus <> "2") Then
                            commandInsertEntrance.Parameters.AddWithValue("@kt_leave_work", inputDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            If isOverTimeNight Then
                                inputDateTime = inputDateTimeParse.AddDays(-1)
                                commandInsertEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                                isOverTimeNight = False
                            Else
                                commandInsertEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                            End If

                            commandInsertEntrance.Parameters.AddWithValue("@worker_no", workerNo)
                        ElseIf (reader("str_Mode").ToString() = "출입" AndAlso {"정문", "연구1층출입문외부", "테크1층출입문외부", "테크1층차고지외부", "테크지하1층출입문외부"}.Contains(reader("str_accTerminalPlace"))) AndAlso Not setNewDate Then
                            commandInsertEntrance.Parameters.AddWithValue("@kt_go_inside", inputDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            commandInsertEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                            commandInsertEntrance.Parameters.AddWithValue("@worker_no", workerNo)
                        ElseIf (reader("str_Mode").ToString() = "출근" OrElse (reader("str_Mode").ToString() = "출입" AndAlso {"정문", "연구1층출입문외부", "테크1층출입문외부", "테크1층차고지외부", "테크지하1층출입문외부"}.Contains(reader("str_accTerminalPlace")))) AndAlso setNewDate Then
                            commandInsertEntrance.Parameters.AddWithValue("@kt_go_to_work", inputDateTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            commandInsertEntrance.Parameters.AddWithValue("@status", status)
                            commandInsertEntrance.Parameters.AddWithValue("@date", inputDateTime.ToString("yyyy-MM-dd 00:00:00"))
                            commandInsertEntrance.Parameters.AddWithValue("@worker_no", workerNo)
                            testText.Text &= "출근 시간 변경되었습니다." & vbCrLf
                        End If


                        ' Create a command with parameters


                        Try
                            ' Execute the INSERT query
                            commandInsert.ExecuteNonQuery()
                            If insertQueryEntrance <> "" Then
                                commandInsertEntrance.ExecuteNonQuery()
                            End If
                            ' Display a message indicating successful insertion
                            testText.Text &= "Data inserted successfully." & vbCrLf
                        Catch ex As Exception
                            ' Handle any exceptions during data insertion
                            testText.Text &= "Error inserting data: " & ex.Message & vbCrLf
                        End Try
                        ' Display a message indicating successful insertion
                        testText.Text &= outputDateString & " " & reader("str_workempName").ToString() & " " & reader("str_accTerminalPlace").ToString() & " " & reader("str_Mode").ToString() & " " & reader("str_ValidationStatus").ToString() & vbCrLf
                        testText.Text &= "Data inserted successfully." & vbCrLf
                    Else
                        'testText.Text &= "Data already exists. Skipping insertion." & vbCrLf
                    End If


                End While

            Else
                testText.AppendText("No data found." & vbCrLf)
            End If

        Catch ex As Exception
            testText.AppendText("Error: " & ex.Message & vbCrLf)
        Finally
            ' Close the connection after retrieving the data
            connection.Close()
        End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub mdbsFilePath_TextChanged(sender As Object, e As EventArgs) Handles mdbsFilePath.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles mdbPassword.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles odbsDsnText.TextChanged

    End Sub

    Private Sub Label4_Click_1(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub IntervalSelected_SelectedIndexChanged(sender As Object, e As EventArgs) Handles IntervalSelected.SelectedIndexChanged

    End Sub
End Class
