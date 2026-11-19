Imports System.Reflection.MethodBase
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports System.Net.Http
Imports System.Threading.Thread
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.CompilerServices
Imports DOC360API
Imports System.Configuration

Imports System.Text
Imports RestSharp
Imports System.Reflection

Public Module GlobalInstances
    Public MyMainForm As frmMain
End Module
Module mdlCorrectedUBProcessing

    Public strICN As String = String.Empty
    Public ManProcFlag As Boolean = False
    Public boolOffshore As Boolean, newDT As DateTime, MachineName As String = ""
    Public PathLog As String, PrevICN As String, FlagAbort As Boolean
    Public EDSVCLine As String
    Dim clientCon As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("clientDB").ConnectionString)
    Dim wiproCon As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("wiproDB").ConnectionString)
    Public Cmd As SqlCommand
    Public dt As DataTable
    Public da As SqlDataAdapter
    Dim wsShell = CreateObject("WScript.Shell")
    Dim citrixUsername As String = wsShell.ExpandEnvironmentStrings("%USERNAME%")
    Dim dt_ListOfHxMatchingANPI As New DataTable
    Dim locality As String
    Public dt_LinesWith50Modifier As DataTable
    Dim dt_iCESCheck As DataTable
    Dim dt_icesMPCScreen As DataTable
    Dim dtClaimInfo As DataTable
    Dim dt_denialRC As DataTable
    Private dtNonServiceRemarkCodes As DataTable
    Private dtMatchedNonServiceCodes As DataTable

    <Extension()>
    Public Function ConvertNothingStringToNullString(ByVal Str As String) As String
        If Str Is Nothing Then Return ""
        Return Str
    End Function
    Public Sub ExitAllAndReturnWithException()
        Dim a As Integer = 0
        Dim b As Integer = 5
        Dim c As Integer
        c = b / a
    End Sub
    Public Sub PutE9InClosureField()
        Dim PageCount As Integer = 1
        Try
            For II As Integer = 6 To 18 Step 3 'Infinite Loop exception is handled
                If GetText(II, 1, 80, True) = "" Then Exit For
                If InStr(InvalidProcCode, GetText(II, 11, 5, True)) = 0 AndAlso GetText(II, 38, 2) <> "30" Then
                    SendKeyes("E9", II, 69)
                End If

                If II = 18 Then
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 OrElse PageCount > 2 Then
                        Exit For
                    Else
                        II = 3
                        PageCount += 1
                        If PageCount > 20 Then Exit For
                    End If
                End If

            Next
            SendEnter()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Private Sub InitializeMatchedNonServiceCodesTable()

        dtMatchedNonServiceCodes = New DataTable("MatchedCodes")

        dtMatchedNonServiceCodes.Columns.Add("RevCode", GetType(String))
        dtMatchedNonServiceCodes.Columns.Add("Charge", GetType(String))
        dtMatchedNonServiceCodes.Columns.Add("RemarkCode", GetType(String))

    End Sub
    Private Sub InitializeNonServiceRemarkCodes()

        dtNonServiceRemarkCodes = New DataTable("RevenueRemarkCodes")

        dtNonServiceRemarkCodes.Columns.Add("RevenueCode", GetType(String))
        dtNonServiceRemarkCodes.Columns.Add("RemarkCode", GetType(String))

        ' Patient convenience items -> Remark Code 13
        dtNonServiceRemarkCodes.Rows.Add("0990", "13")
        dtNonServiceRemarkCodes.Rows.Add("0991", "13")
        dtNonServiceRemarkCodes.Rows.Add("0992", "13")
        dtNonServiceRemarkCodes.Rows.Add("0993", "13")
        dtNonServiceRemarkCodes.Rows.Add("0994", "13")
        dtNonServiceRemarkCodes.Rows.Add("0995", "13")
        dtNonServiceRemarkCodes.Rows.Add("0998", "13")

        ' Experimental drugs -> Remark Code YG
        dtNonServiceRemarkCodes.Rows.Add("0256", "YG")

    End Sub
    Private Function GetRemarkCode(ByVal revCode As String) As String

        Dim rows() As DataRow = dtNonServiceRemarkCodes.Select(
        "RevenueCode = '" & revCode.Trim() & "'")

        If rows.Length > 0 Then
            Return rows(0)("RemarkCode").ToString()
        End If

        Return String.Empty

    End Function
    Function GetManualInput(ByVal input As String)
        Dim val As String = Trim(InputBox("Enter " + input))
        Return val
    End Function
    Public Sub SendOverrideInScreen(ByVal Override As String)
        Try
            If GetText(2, 2, 2) = "MP" Then
                For I As Integer = 10 To 22 Step 2
                    If InStr(PISLLINE, GetText(I, 4, 6, True)) = 0 OrElse InStr("PR, IC, NB, SP, WD", GetText(I, 4, 6, True)) > 0 And GetText(I, 1, 80, True).Replace("-", "").Trim <> "" Then
                        If GetText(I, 29, 1) = "P" And InStr("PM,CM,CV,AD", GetText(I, 36, 2)) > 0 Then
                            autECLPSObj.SendKeys("09", I, 29)
                            apiChk()
                        Else
                            If GetText(I, 29, 2) <> "30" Then autECLPSObj.SendKeys(Override, I, 29)
                            apiChk()
                        End If
                        If GetColorScheme(I, 36) = 3 Then
                            autECLPSObj.SendKeys("--", I, 36)
                            apiChk()
                        End If

                    End If
                Next
                autECLPSObj.sendkeys("[ENTER]")
                apiChk()
            ElseIf GetText(3, 2, 5) = "EDS 6" Then
                Dim count As Integer = 0
                For I As Integer = 6 To 18 Step 3
                    If GetText(I, 11, 6, True) = "" Then
                        Exit For
                    End If

                    If InStr(PISLLINE, GetText(I, 11, 6, True)) = 0 Then
                        autECLPSObj.SendKeys(Override, I, 38)
                        apiChk()
                    End If
                    If I = 18 Then
                        autECLPSObj.sendkeys("[ENTER]")
                        apiChk()
                        SendF8()
                        count += 1
                        If InStr(GetText(24, 1, 80), "NO MORE RECORDS") > 0 Then
                            Exit For
                        Else
                            If count > 50 Then Exit For
                            I = 3
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function ConvertToFormat(ByVal Str As String, ByVal Num As Integer, Optional ByVal Pre As Boolean = False, Optional ByVal text As String = "") As String
        Try
            If IsNumeric(Str) Then
                If Microsoft.VisualBasic.Left(Str, 1) = "-" Then Str = Microsoft.VisualBasic.Right(Str, Len(Str) - 1)
            End If
            While Len(Str) < Num
                If Pre Then
                    If text = "" Then
                        Str = " " & Str
                    Else
                        Str = text & Str
                    End If
                Else
                    If text = "" Then
                        Str = Str & " "
                    Else
                        Str = Str & text
                    End If
                End If
            End While
            If Len(Str) > Num Then
                If Pre Then
                    Str = Microsoft.VisualBasic.Right(Str, Num)
                Else
                    Str = Microsoft.VisualBasic.Left(Str, Num)
                End If
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        ConvertToFormat = Str
    End Function

    Public Function IsAlphaNumeric(ByVal strInput As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(strInput, "^[a-zA-Z]{2}[0-9]{8}$")
    End Function
    Public Function CheckIfAlphaCharacterIsPresent(ByVal StringToCheck As String) As Boolean
        For i = 0 To StringToCheck.Length - 1
            If Char.IsLetter(StringToCheck.Chars(i)) Then
                Return True
            End If
        Next

        Return False 'Return false if any elements is not a character
    End Function
    Public Sub AbortUpdate(ByVal icn As String)
        Dim con As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("clientDB").ConnectionString)
        Dim conEWD As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("hhhhDB").ConnectionString)

        Dim cmd1, cmd2 As SqlCommand
        Dim query, query2 As String
        'cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\unpiox53pn\Kolkata_Claim_Drive\Sourav Maji\CorrectedAutomation\Corrected.accdb; Persist Security Info=False;"

        Try
            con.Open()
            If conEWD.State = ConnectionState.Closed Then
                conEWD.Open()
            End If
            If boolOffshore = True Then
                query = "Update WCRPaymentLog UBStatus='Aborted' WHERE ICN ='" + icn + "' "
                query2 = "Update CliffsKFCWaffles_Api MSID=@msid WHERE ICN ='" + icn + "' "
            Else
                query = "Update WCRPaymentLog UBStatus='Aborted' WHERE ICN ='" + icn + "' "
                query2 = "Update CliffsKFCWaffles_Api MSID=@msid WHERE ICN ='" + icn + "' "
            End If
            WriteLog("Query1= " + query)
            cmd1 = New SqlCommand(query, con)
            Dim a As Integer = cmd1.ExecuteNonQuery()
            WriteLog("Query1 execution returned= " + a.ToString())
            cmd2 = New SqlCommand(query2, con)
            cmd2.Parameters.AddWithValue("msid", DBNull.Value)
            WriteLog("Query2= " + query2)
            Dim b As Integer = cmd2.ExecuteNonQuery()
            WriteLog("Query2 execution returned= " + b.ToString())
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

            Console.WriteLine(ex.Message.ToString)
            WriteLog(ex.Message.ToString)
        Finally
            con.Close()
            conEWD.Close()
        End Try
    End Sub
    Public Sub AbortUpdate(ByVal icn As String, ByVal comment As String)
        Dim con As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("clientDB").ConnectionString)
        Dim conEWD As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("hhhhDB").ConnectionString)

        Dim cmd1, cmd2 As SqlCommand
        Dim query, query2 As String
        'cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\unpiox53pn\Kolkata_Claim_Drive\Sourav Maji\CorrectedAutomation\Corrected.accdb; Persist Security Info=False;"

        Try
            con.Open()
            If conEWD.State = ConnectionState.Closed Then
                conEWD.Open()
            End If
            If boolOffshore = True Then
                query = $"Update WCRPaymentLog UBStatus='Aborted', Result='{comment}' WHERE ICN ='" + icn + "' "
                query2 = "Update CliffsKFCWaffles_Api MSID=@msid WHERE ICN ='" + icn + "' "
            Else
                query = $"Update WCRPaymentLog UBStatus='Aborted', Result='{comment}' WHERE ICN ='" + icn + "' "
                query2 = "Update CliffsKFCWaffles_Api MSID=@msid WHERE ICN ='" + icn + "' "
            End If
            WriteLog("Query1= " + query)
            cmd1 = New SqlCommand(query, con)
            Dim a As Integer = cmd1.ExecuteNonQuery()
            WriteLog("Query1 execution returned= " + a.ToString())
            cmd2 = New SqlCommand(query2, con)
            cmd2.Parameters.AddWithValue("msid", DBNull.Value)
            WriteLog("Query2= " + query2)
            Dim b As Integer = cmd2.ExecuteNonQuery()
            WriteLog("Query2 execution returned= " + b.ToString())
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

            Console.WriteLine(ex.Message.ToString)
            WriteLog(ex.Message.ToString)
        Finally
            con.Close()
            conEWD.Close()
        End Try
    End Sub
    Public Function FetchUserID() As String
        Dim id As String
        Dim loc As String
        locality = ""
        If boolOffshore Then
            loc = "Global"
            locality = "Global"
        Else
            loc = "Domestic"
            locality = "Domestic"
        End If

        Try
            wiproCon.Open()
            Cmd = New SqlCommand("Select [ID] FROM [Wipro_Corrected_Droid].[dbo].[DroidCred] where Locality='" + loc + "'", wiproCon)
            da = New SqlDataAdapter(Cmd)
            dt = New DataTable
            da.Fill(dt)

            id = CStr(dt.Rows(0)("ID"))

            wiproCon.Close()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

            'MessageBox.Show(ex.Message.ToString)
        Finally
            If wiproCon.State = ConnectionState.Open Then
                wiproCon.Close()
            End If
        End Try
        FetchUserID = id

    End Function

    Public Function FetchUserPassword() As String
        Dim pass As String
        Dim loc As String
        If boolOffshore Then
            loc = "Global"
        Else
            loc = "Domestic"
        End If

        Try
            wiproCon.Open()
            Cmd = New SqlCommand("SELECT [Pass] FROM [Wipro_Corrected_Droid].[dbo].[DroidCred] where Locality='" + loc + "'", wiproCon)
            da = New SqlDataAdapter(Cmd)
            dt = New DataTable
            da.Fill(dt)

            pass = CStr(dt.Rows(0)("Pass"))

            wiproCon.Close()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        Finally
            If wiproCon.State = ConnectionState.Open Then
                wiproCon.Close()
            End If
        End Try
        FetchUserPassword = pass

    End Function
    Public Function callOgaAPI(office As String, claim_desc As String, reportingPool As String, icn As String, engine As String, claimAge As String, locality As String,
                    msid As String) As Boolean
        Dim success As Boolean = False
        Try
            WriteLog("==========================Inside " + GetCurrentMethod.Name() + "==========================")
            Dim client As New HttpClient
            Dim url As String = "https://eni-claims-dev.optum.com/load_data"
            WriteLog("URL: " + url)
            Dim requestData As New With {
                .office = office,
                .claim_spec_desc = claim_desc,
                .reporting_pool = reportingPool,
                .icn = icn,
                .pay_eng = engine,
                .claim_age = claimAge,
                .sites = locality,
                .msid = msid
            }

            Dim json As String = JsonConvert.SerializeObject(requestData)

            Dim content As New StringContent(json, Encoding.UTF8, "application/json")
            Dim response As HttpResponseMessage = client.PostAsync(url, content).Result
            Dim responseBody As String = response.Content.ReadAsStringAsync().Result
            Dim statusCode As String = response.StatusCode.ToString()
            If statusCode = "OK" Or responseBody.Contains("Data already present in table") Then success = True

            WriteLog("Status code: " + statusCode + vbNewLine + "Response body: " + responseBody.ToString())

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + " exception as " + ex.Message)
        End Try

        Return success
    End Function

    Public Function CallOptumAPI(icnnn As String, Optional Priority As String = "") As String
        WriteLog("====================ENTERING " + GetCurrentMethod.Name + "================")
        'Dim icnnn As String = "AC01483175"
        'Dim Priority As String = "2"
        Dim client = New RestClient("http://ewdrtng-srvcs-mprd1.optum.com/determineEligibilityAndSendToEWDIncrementPriority/" + icnnn + "/" + Priority + "")
        client.Timeout = -1
        WriteLog("URL=" + client.BaseUrl.ToString)
        Dim request = New RestRequest(Method.GET)
        request.AddHeader("Accept", "text/plain")
        Dim response As IRestResponse = client.Execute(request)
        Dim responsecode = response.StatusCode
        WriteLog("ResponseCode:" + responsecode.ToString)
        Dim responsecontent = response.Content
        WriteLog("Response content: " + responsecontent)

        If responsecode.ToString <> "OK" Then
            Dim client1 = New RestClient("http://ewdrtng-srvcs-mprd2.optum.com/determineEligibilityAndSendToEWDIncrementPriority/" + icnnn + "/" + Priority + "")
            client1.Timeout = -1
            WriteLog("URL=" + client1.BaseUrl.ToString)
            Dim request1 = New RestRequest(Method.GET)
            request1.AddHeader("Accept", "text/plain")
            response = client.Execute(request)
            responsecode = response.StatusCode
            WriteLog("ResponseCode:" + responsecode.ToString)
            responsecontent = response.Content
            WriteLog("Response content: " + responsecontent)
        End If
        Return responsecontent
        WriteLog("====================EXIT " + GetCurrentMethod.Name + "================")
    End Function
    Private Function IsPaylocInfrastructureError(ByVal Status As String) As Boolean

        If String.IsNullOrEmpty(Status) Then Return False

        Status = Status.ToUpper.Trim

        Return Status.Contains("PAYLOC_TIMEOUT_RETRY") _
        Or Status.Contains("INVALID COMMAND SYNTAX") _
        Or Status.Contains("COMMAND UNRECOGNIZED") _
        Or Status.Contains("HOST NOT AVAILABLE") _
        Or Status.Contains("SESSION FAILURE")

    End Function
    Private Sub ReQueueICN(ByVal ICN As String, ByVal comment As String)

        Try

            WriteLog("Infrastructure issue detected. Releasing ICN : " & ICN)

            AbortUpdate(ICN, comment)

            autECLPSObj.StopCommunication()

            Threading.Thread.Sleep(2000)

            autECLPSObj.StartCommunication()

            Threading.Thread.Sleep(2000)

        Catch ex As Exception

            WriteLog("ReQueueICN Error : " & ex.Message)

        End Try

    End Sub
    Public Sub Start(myForm As frmMain)
        Dim ID As String = String.Empty
        Dim ICN As String = String.Empty
        Dim Engine As String = String.Empty
        Dim Payloc As String = String.Empty
        Dim UserID As String = String.Empty
        Dim Password As String = String.Empty
        Dim session As String = String.Empty
        Dim ndbSession As String = String.Empty
        Dim status As String = String.Empty
        Dim TempStatus As String = String.Empty
        Dim TPvalue As String
        Dim IsPGPolicy As Integer
        Dim OrgID As String = String.Empty
        UserID = FetchUserID()
        Password = FetchUserPassword()
        UnetID = UserID
        session = Trim(myForm.txtSession.Text)
        ndbSession = Trim(myForm.txtNdbSession.Text)
        Dim dt As DataTable
        dt_PG_Policyno = New DataTable
        dt_PG_Policyno = GetPgpolicylist()
        InitializeNonServiceRemarkCodes()

        Try
            Do
                myForm.updateStatus("Searching ICN")

DBFetch:        dt = New DataTable()
                StartDateMonth = DateTime.Today.ToString("MMM")
                StartDateLong = DateTime.Today.ToString("MMddyyyy")
                Dim query1 = ""
                Dim query2 = ""
                strResearchComments = String.Empty
                status = String.Empty
                TempStatus = String.Empty
                'SQLWorkDataRead(myForm, dt, query1, query2)
                AddTestRows(dt)

                For row As Integer = 0 To dt.Rows.Count - 1
                    Dim SuffixCount As Integer
                    dt_ListOfHxMatchingANPI = New DataTable
                    dt_LinesWith50Modifier = New DataTable
                    dt_iCESCheck = New DataTable
                    dt_icesMPCScreen = New DataTable
                    dtClaimInfo = New DataTable
                    dt_denialRC = New DataTable
                    'dtMatchedNonServiceCodes.Rows.Clear()
                    InitializeMatchedNonServiceCodesTable()

                    Dim realICN As String = String.Empty
                    TPvalue = String.Empty
                    IsPGPolicy = 0
                    ICN = Trim(dt.Rows(row)(1).ToString())
                    Payloc = Trim(dt.Rows(row)(3).ToString())
                    If Payloc.Length = 1 Then
                        Payloc = "00" + Payloc
                    ElseIf Payloc.Length = 2 Then
                        Payloc = "0" + Payloc
                    Else
                        Payloc = Payloc
                    End If
                    Engine = Trim(dt.Rows(row)(2).ToString())
                    ID = Trim(dt.Rows(row)(0).ToString())

                    myForm.updateCurrentICN(ICN)

                    CurrICN = ICN
                    Office = Payloc
                    Engine1 = Engine
                    WriteLog("================= Entering " + GetCurrentMethod.Name + "====================")
                    WriteLog("Query2= " + query2)
                    WriteLog("Query1= " + query1)
                    realICN = ICN
                    If ICN = "" Then
                        myForm.updateCurrentICN(ICN)
                        For waitTime As Integer = 10 To 1 Step -1
                            myForm.updateStatus("Next ICN Search after " + CStr(waitTime) + " Seconds...")
                            myForm.updateCurrentICN("")
                            System.Threading.Thread.Sleep(1000)
                        Next
                    Else
                        InitializeEmulator(session)
                        autECLPSObj.StopCommunication()
                        Do

                        Loop While autECLPSObj.CommStarted
                        autECLPSObj.StartCommunication()
                        Do
                            apiChk()
                        Loop Until autECLPSObj.CommStarted
                        Threading.Thread.Sleep(3000)

                        myForm.updateStatus("Processing...")
                        status = UBClaimResearchStart(myForm, ICN, Engine, Payloc, UserID, Password, session, TPvalue, IsPGPolicy, dt_PG_Policyno, ndbSession, ID, SuffixCount)
                        WriteLog("STATUS:===> " + status)
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If IsPaylocInfrastructureError(status) Then

                            WriteLog("PAYLOC / ENGINE DOWN DETECTED")

                            ReQueueICN(ICN, "PAYLOC / ENGINE DOWN DETECTED")

                            Continue For

                        End If
                        ''''''''''''''        ^^^^^^^^^^^    ''''''''''''''''''''''''''''''
                        If status = "ICN picked by different user" And myForm.ThreadRun <> False Then
                            GoTo DBFetch
                        ElseIf status = "ICN picked by different user" And myForm.ThreadRun = False Then
                            Exit Do
                        End If
                        If InStr(status.ToLower, "exception") > 0 And (status.Contains(Chr(34)) Or status.Contains(Chr(39))) And status.Trim.Contains("MANUAL INTERVENTION REQUIRED") = False Then
                            status = status.Replace(Chr(39), String.Empty)
                            status = status.Replace(Chr(34), String.Empty)
                            status = status.Replace(Chr(46), String.Empty)
                            status = "Exception"
                        End If
                        If InStr(status.ToLower, "conversion") > 0 Then
                            status = "Conversion failed"
                        End If
                        If IsHDClaim = True Then
                            status = "HD claim: " + status
                        End If

                        If MPCScreenCount > 0 And InStr(status, "Claim process by tool") = 0 AndAlso status <> "Already Processed Claim" AndAlso status <> "EDS1 NYS Line present" AndAlso
                            status <> "COVID Claim" AndAlso status <> "Suffix greater than one" AndAlso status <> "E028NOT AUTHORIZED," AndAlso
                            status <> "CLAIM AWAITING AUTO ADJUDICATION" AndAlso status <> "Claim is not Corrected Claim" AndAlso status <> "Claim is not an electronic/paper submission" AndAlso
                            status <> "PROCESSED; Paid amount>5000" AndAlso status <> "Sent to MARS EDITING" Then
                            'RevertBack(realICN)
                            If GetText(2, 2, 2) = "MP" Then
                                If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                                    SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                                    If status <> "Claim process by tool" Then
                                        SendOverrideInScreen("--")
                                    Else
                                        SendOverrideInScreen("01")
                                    End If
                                    CaptureScreenLog()
                                End If
                            End If
                        End If
                        WriteLog("Number of times MPC screen is hit =" + MPCScreenCount.ToString)


                        'If MPCScreenCount > 0 And status <> "Claim process by tool" Then
                        If status <> "E030OFFICE INVALID" And InStr(status, "Claim process by tool") = 0 And status <> "Already Processed Claim" AndAlso
                            status <> "Claim is not an electronic/paper submission" AndAlso status <> "Suffix greater than one" AndAlso status <> "E028NOT AUTHORIZED," AndAlso
                            status <> "CLAIM AWAITING AUTO ADJUDICATION" AndAlso status <> "Claim is not Corrected Claim" Then

                            Dim retVal As String = RemoveDummyDraftandICN(ICN, Payloc)
                            If flagHistoryVoid = True Then
                                'GenerateHardEdit()
                            End If
                            If retVal <> "" Then status = status + "," + retVal
                            TempStatus = REMOVE01IFSKIPPED(ICN, SuffixCount)
                        End If

                        If TempStatus = "" Or IsNothing(TempStatus) Then
                            UpdateData(ICN, status, TPvalue, IsPGPolicy, ID, Engine, Payloc, locality, OrgID, SuffixCount)
                        Else
                            status = status + " //" + TempStatus
                            UpdateData(ICN, status, TPvalue, IsPGPolicy, ID, Engine, Payloc, locality, OrgID, SuffixCount)
                        End If
                        'If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                        '    SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                        '    CaptureScreenLog()
                        'End If
                        If clientCon.State = ConnectionState.Open Then clientCon.Close()
                        If wiproCon.State = ConnectionState.Open Then wiproCon.Close()

                    End If
                    If myForm.ThreadRun = False Then Exit Do
                    If (row = dt.Rows.Count - 1) Then
                        GoTo DBFetch
                    End If
                Next

            Loop While myForm.ThreadRun = True
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

            WriteLog("Exception: " + ex.Message.ToString, True)
        Finally
            If clientCon.State = ConnectionState.Open Then clientCon.Close()
            If wiproCon.State = ConnectionState.Open Then wiproCon.Close()
            Console.WriteLine("EXIT " + GetCurrentMethod.Name)
        End Try
    End Sub
    Public Sub RevertBack(ByVal ICN As String)
        Try
            SendClear()
            SendClear()
            SendKeyes("mei,i" + ICN + ",9", 2, 2)
            apiChk()
            SendEnter()
            apiChk()
            SendKeyes("A", 1, 76)
            SendEnter()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            apiChk()
            If Trim(GetText(2, 2, 2)) = "AR" And InStr(strResearchComments, "ARI screen") > 0 Then
                Exit Sub
            End If
            'PutOriginalValue()
            If InStr(strResearchComments, "exception") > 0 Then
                Exit Sub
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    'Public Sub UpdateData1(ByVal icn As String, ByVal status As String)
    '    Try
    '        If status = "Claim process by tool" Then
    '            UpdateSQLParametrizedDataReader({"PaymentComment", "Result", "Da62teProcess", "UBStatus"}, {PaymentComment, status, "getdate()", status}, icn, "WCRPaymentLog")
    '            UpdateSQLParametrizedDataReader({"MSID", "Filter_Droid", "ZeroPaid", "InapCheckOut"}, {"UBPdroid", "Ran", "Processed", MachineName}, icn, "CliffsKFCWaffles")
    '        Else
    '            UpdateSQLParametrizedDataReader({"ResearchComment", "Result", "UBStatus"}, {status, "Skip", "Skip"}, icn, "WCRPaymentLog")

    '            If isHistoryICNVoided = True Then
    '                UpdateSQLParametrizedDataReader({"MSID", "Filter_Droid", "ZeroPaid", "Priority_Flag", "InapCheckOut"}, {"", "Ran", "Skip", "4", MachineName}, icn, "CliffsKFCWaffles")
    '            Else
    '                UpdateSQLParametrizedDataReader({"MSID", "Filter_Droid", "ZeroPaid", "Priority_Flag", "InapCheckOut"}, {"", "Ran", "Skip", "", MachineName}, icn, "CliffsKFCWaffles")
    '            End If
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
    Public Sub UpdateData(ByVal icn As String, ByVal status As String, ByVal TPvalue As String, ByVal IsPGPolicy As Integer,
                          ByVal ID As String, Engine As String, Payloc As String, ByVal locality As String, ByVal orgID As String, ByVal SuffixCount As Integer)

        Dim clientCon As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("clientDB").ConnectionString)
        Dim clientConEWD As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("hhhhDB").ConnectionString)

        Try
UpdData:    WriteLog("================= Entering " + GetCurrentMethod.Name + "====================")
            Dim cmd1, cmd2 As SqlCommand
            Dim query1, query2, query4 As String
            Dim response As String = ""
            If clientCon.State = ConnectionState.Closed Then clientCon.Open()
            If clientConEWD.State = ConnectionState.Closed Then clientConEWD.Open()

            If InStr(status, "Claim process by tool") > 0 Then

                'Check for AV pend
                '''''''''''Code''''''''''''''



                'Check for overpayment
                If InStr(strResearchStatus, "Overpayment adjustment performed") > 0 Then
                    query1 = "Update CliffsKFCWaffles_Api set MSID='UBPdroid', Filter_Droid='Ran', ZeroPaid='Processed', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                    query2 = "Update WCRPaymentLog set SuffixCount='" + SuffixCount.ToString() + "' ,PickStatus= 'Done',IsPgpolicy='" + IsPGPolicy.ToString + "', Isbilltype='" + TPvalue + "' ,PaymentComment='" + strResearchStatus + "', Result='" + status + "', DateProcess=getdate(), UBStatus='" + status + "',HistoryICNList='" + FinalHistoryICNList + "',PCCmd='OVPT004 need to be sent' where ICN='" + icn + "' and ID='" + ID + "'"

                    WriteLog(query1)
                    cmd1 = New SqlCommand(query1, clientConEWD)
                    Dim c As Integer = cmd1.ExecuteNonQuery()

                    WriteLog("Query1 execution returned= " + c.ToString())
                    cmd2 = New SqlCommand(query2, clientCon)
                    WriteLog(query2)
                    Dim d As Integer = cmd2.ExecuteNonQuery()
                    'While b <> 1
                    '    b = cmd2.ExecuteNonQuery
                    'End While
                    WriteLog("Query2 execution returned= " + d.ToString())
                Else
                    query1 = "Update CliffsKFCWaffles_Api set MSID='UBPdroid', Filter_Droid='Ran', ZeroPaid='Processed', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                    query2 = "Update WCRPaymentLog set SuffixCount='" + SuffixCount.ToString() + "' ,PickStatus= 'Done',IsPgpolicy='" + IsPGPolicy.ToString() + "', Isbilltype='" + TPvalue + "' ,PaymentComment='" + PaymentComment + "', Result='" + status + "', DateProcess=getdate(), UBStatus='" + status + "',HistoryICNList='" + FinalHistoryICNList + "' where ICN='" + icn + "' and ID='" + ID + "'"

                    WriteLog(query1)
                    cmd1 = New SqlCommand(query1, clientConEWD)
                    Dim a As Integer = cmd1.ExecuteNonQuery()
                    'While a <> 1
                    '    a = cmd1.ExecuteNonQuery()
                    'End While
                    WriteLog("Query1 execution returned= " + a.ToString())
                    cmd2 = New SqlCommand(query2, clientCon)
                    WriteLog(query2)
                    Dim b As Integer = cmd2.ExecuteNonQuery()
                    'While b <> 1
                    '    b = cmd2.ExecuteNonQuery
                    'End While
                    WriteLog("Query2 execution returned= " + b.ToString())
                End If


                '>>>>>>>>>>>>>>>>>>>> Commented out exception related update query @6th Dec,2024 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                'ElseIf InStr(status, "ARITHMETIC OPERATION RESULTED IN AN OVERFLOW") > 0 OrElse InStr(status, "HIT CLEAR TO RETURN TO TOPS") > 0 OrElse InStr(status, "API object is NOTHING") Or InStr(status, "exception") > 0 Then
                '    query1 = "Update CliffsKFCWaffles_Api set MSID= @msid WHERE ICN ='" + icn + "' "
                '    query2 = "Update WCRPaymentLog set PickStatus= @pick, UBStatus=@ubstat, IsPgpolicy='" + IsPGPolicy.ToString + "', Isbilltype='" + TPvalue + "' where ICN='" + icn + "' and ID='" + ID + "'"

                '    cmd1 = New SqlCommand(query1, clientConEWD)
                '    cmd1.Parameters.AddWithValue("msid", DBNull.Value)
                '    WriteLog(query1)
                '    Dim a As Integer = cmd1.ExecuteNonQuery()
                '    'While a <> 1
                '    '    a = cmd1.ExecuteNonQuery()
                '    'End While
                '    WriteLog("Query1 execution returned= " + a.ToString())

                '    cmd2 = New SqlCommand(query2, clientCon)
                '    cmd2.Parameters.AddWithValue("pick", DBNull.Value)
                '    cmd2.Parameters.AddWithValue("ubstat", DBNull.Value)
                '    WriteLog(query2)
                '    Dim b As Integer = cmd2.ExecuteNonQuery()
                '    'While b <> 1
                '    '    b = cmd2.ExecuteNonQuery
                '    'End While
                '    WriteLog("Query2 execution returned= " + b.ToString())
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Else
                If status = "NDB login failed" Then
                    query4 = "Update WCRPaymentLog set SuffixCount='" + SuffixCount.ToString() + "' ,PickStatus= @null, IsPgpolicy='" + IsPGPolicy.ToString + "', Isbilltype='" + TPvalue + "' , ResearchComment='Facility claim', Result=@null, UBStatus=@null,HistoryICNList='" + FinalHistoryICNList + "' where ICN='" + icn + "' and ID='" + ID + "'"
                    'query3 = "Update CliffsKFCWaffles_Api [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='0', Priority_Flag='2', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "

                    'cmd1 = New SqlCommand(query3, clientConEWD)
                    'cmd1.Parameters.AddWithValue("stat", DBNull.Value)
                    'cmd1.Parameters.AddWithValue("msid", DBNull.Value)
                    'WriteLog(query3)
                    'Dim a As Integer = cmd1.ExecuteNonQuery()
                    'While a <> 1
                    '    a = cmd1.ExecuteNonQuery()
                    'End While
                    'WriteLog("Query1 execution returned= " + a.ToString())

                    cmd2 = New SqlCommand(query4, clientCon)
                    cmd2.Parameters.AddWithValue("null", DBNull.Value)
                    WriteLog(query4)
                    Dim b As Integer = cmd2.ExecuteNonQuery()
                    'While b <> 1
                    '    b = cmd2.ExecuteNonQuery
                    'End While
                    WriteLog("Query2 execution returned= " + b.ToString())
                Else
                    query2 = "Update WCRPaymentLog set SuffixCount='" + SuffixCount.ToString() + "' ,PickStatus= 'Done', IsPgpolicy='" + IsPGPolicy.ToString + "', Isbilltype='" + TPvalue + "' , ResearchComment='" + status + "', Result='Skip', UBStatus='Skip',HistoryICNList='" + FinalHistoryICNList + "' where ICN='" + icn + "' and ID='" + ID + "'"
                    If isHistoryICNVoided = True Then
                        response = CallOptumAPI(icn, "2")
                        If InStr(response.ToLower, "ewd manual routing event successfully dispatched") > 0 Then
                            query1 = "Update CliffsKFCWaffles_Api set [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='1', Priority_Flag='2', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                            query2 = $"Update WCRPaymentLog set  SuffixCount='{SuffixCount.ToString()}' ,PickStatus= 'Done', IsPgpolicy='{IsPGPolicy.ToString()}', Isbilltype='{ TPvalue}' , ResearchComment='{status }', Result='Skip', UBStatus='Skip',HistoryICNList='{FinalHistoryICNList }' where ICN='{ icn }' and ID='{ ID }'"
                        Else
                            query1 = "Update CliffsKFCWaffles_Api set [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='0', Priority_Flag='2', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                        End If

                    Else
                        If status = "Sent to MARS EDITING" Or InStr(status, "EOC DATA DIFFERENT") > 0 Or InStr(status, "Bill Type 8 skipped") > 0 Then
                            query1 = "Update CliffsKFCWaffles_Api set [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', Priority_Flag='4', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                        Else
                            response = CallOptumAPI(icn, "4")
                            If InStr(response.ToLower, "ewd manual routing event successfully dispatched") > 0 Then
                                query1 = "Update CliffsKFCWaffles_Api SET [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='1', Priority_Flag='4', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                            Else
                                query1 = "Update CliffsKFCWaffles_Api SET [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='0', Priority_Flag='4', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                            End If
                            'Dim APICallSuccess As Boolean = callOgaAPI(Payloc, orgID, "Hello", icn, Engine, "6", locality, "holmes_aut_u1")
                            'If APICallSuccess = True Then
                            '    query1 = "Update CliffsKFCWaffles_Api set [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='1', Priority_Flag='4', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                            'Else
                            '    query1 = "Update CliffsKFCWaffles_Api set [Status]=@stat, MSID= @msid, Filter_Droid='Ran', ZeroPaid='Skip', ORS='0', Priority_Flag='4', InapCheckOut ='" + MachineName + "' WHERE ICN ='" + icn + "' "
                            'End If
                        End If

                    End If

                    cmd1 = New SqlCommand(query1, clientConEWD)
                    cmd1.Parameters.AddWithValue("stat", DBNull.Value)
                    cmd1.Parameters.AddWithValue("msid", DBNull.Value)
                    WriteLog(query1)
                    Dim a As Integer = cmd1.ExecuteNonQuery()
                    'While a <> 1
                    '    a = cmd1.ExecuteNonQuery()
                    'End While
                    WriteLog("Query1 execution returned= " + a.ToString())

                    cmd2 = New SqlCommand(query2, clientCon)
                    WriteLog(query2)
                    Dim b As Integer = cmd2.ExecuteNonQuery()
                    'While b <> 1
                    '    b = cmd2.ExecuteNonQuery
                    'End While
                    WriteLog("Query2 execution returned= " + b.ToString())
                End If

            End If

        Catch ex As Exception
            WriteLog("EXCEPTION=>" + ex.Message)
            GoTo UpdData
        Finally
            If clientCon.State = ConnectionState.Open Then clientCon.Close()
            If clientConEWD.State = ConnectionState.Open Then clientConEWD.Close()
            WriteLog("================= Exiting " + GetCurrentMethod.Name + "====================")
        End Try
    End Sub
    ''' <summary>
    ''' To check revenue codes found in EDS1 matching with 
    ''' Non-service revenue codes, if matches then add to 
    ''' the declared datatable
    ''' </summary>
    ''' <param name="objEDS1Value"></param>
    ''' <returns></returns>
    Private Function PopulateMatchedNonServiceCodes(ByVal objEDS1Value() As EDS1Value) As Boolean
        InitializeMatchedNonServiceCodesTable()
        Dim flag As Boolean = False
        Try
            For i As Integer = 0 To UBound(objEDS1Value)
                For j As Integer = 0 To UBound(objEDS1Value(i).ServiceLine.RevCode)

                    Dim revCode As String = objEDS1Value(i).ServiceLine.RevCode(j)
                    Dim remarkCode As String = GetRemarkCode(revCode)
                    Dim charge As String = objEDS1Value(i).ServiceWiseTotalCharge

                    If remarkCode <> String.Empty Then
                        ' Revenue code found in lookup table
                        dtMatchedNonServiceCodes.Rows.Add(
                                        revCode,
                                        charge, remarkCode)

                        flag = True
                        WriteLog("Non-Covered Services | Rev Code : " & revCode &
                          " | Denial Remark Code : " & remarkCode & " | Bill charge : " & charge)
                    End If

                Next
            Next
        Catch ex As Exception
            WriteLog($"EXCEPTION from {GetCurrentMethod.Name} : | {ex.Message} |")

        End Try

        Return flag
    End Function

    Private Function GetUnitsFromEDS1(ByVal objEDS1Value() As EDS1Value) As String
        Dim unitsBilled As String = String.Empty
        Dim charge As String
        Dim revCode As String
        Try
            For i As Integer = 8 To 17 Step 3
                If GetText(i, 38, 5).Replace("-", "").Trim() = "" Then

                    For jj As Integer = 0 To UBound(objEDS1Value)
                        For kk As Integer = 0 To UBound(objEDS1Value(i).ServiceLine.RevCode)

                        Next
                    Next
                End If
                If i = 17 Then
                    Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                        SendF3()
                        Exit For
                    ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                        SendF3()
                        Exit For
                    Else
                        i = 5
                    End If
                End If
            Next
        Catch ex As Exception

        End Try
        Return unitsBilled
    End Function
    Public Function UBClaimResearchStart(ByRef UsrCntrlAdmin As frmMain, ByVal ICN As String, ByVal Engine As String,
                                         ByVal Payloc As String, ByVal UserID As String, ByVal PassWord As String,
                                         ByVal session As String, ByRef TPvalue As String, ByRef IsPGPolicy As Integer,
                                         ByVal dt_PG_Policyno As DataTable, ByVal ndbsession As String, id As Integer, ByRef MaxSuffix As Integer) As String
        WriteLog("Enter UBClaimResearchStart.......")
        Dim FinalComment As String
        isModifierPresent = False
        isC68REV = False
        NCopayNBR = String.Empty
        FinalHistoryICNList = ""
        isCurrentClaimIH = False
        isCurrentClaimOH = False
        billType = String.Empty
        FinalComment = String.Empty
        Dim FlagError As Boolean = False
        Dim flagNonCoveredServices As Boolean = False
        nonServiceCodePresent = False
        Dim EditWarning As String = String.Empty
        Dim count As Integer
        Dim MarketIDLine As String = String.Empty
        Dim MarketId As String = String.Empty
        Dim IRIAllowable As String = String.Empty
        Dim OCIAllowable As String = String.Empty
        DenialRC = New String() {}
        flagDenial = False
        OriginalICN = String.Empty
        OriginalDraft = String.Empty
        isNCopayNBR = False
        PaymentComment = String.Empty
        Dim CurrentOII As OIIDetails
        isHistoryICNVoided = False
        flagHistoryVoid = True
        Dim LineColorCheck As Boolean = False
        Dim MaxSuffixCountErrorFlag As Boolean = False
        IsHDClaim = False
        flag_QC = False
        MPCScreenCount = 0
        SuffixCountOfCurrentICN = 0
        strResearchComments = String.Empty
        'InitializeEmulator(session)
        SendClear()
        SendClear()
        WriteLog("Payloc Log in Starts...")
        Dim LogInFlag As Boolean = False
        Dim LogInCount As Integer = 0
        While Not LogInFlag
            Try
                Dim oUnet As UnetHelper.UnetHelper = New UnetHelper.UnetHelper(session)
                oUnet.PayLocSignOnProc(Payloc, Engine, session, UserID, PassWord)
                ''PaylocLogIn(Engine, Payloc, session, UserID, PassWord)

                '------------------------------------------------
                'If Trim(GetTextRect(1, 1, 24, 80)).Contains("INVALID COMMAND SYNTAX") Or Trim(GetTextRect(1, 1, 24, 80)).Contains("COMMAND UNRECOGNIZED") Then
                '    strResearchComments = "Payloc_timeout_retry"
                '    Return strResearchComments
                'End If
                Dim screenText As String = Trim(GetTextRect(1, 1, 24, 80)).ToUpper()
                If screenText.Contains("INVALID COMMAND SYNTAX") _
                    Or screenText.Contains("COMMAND UNRECOGNIZED") Then

                    WriteLog("PAYLOC DOWN DETECTED")

                    SendClear()

                    Threading.Thread.Sleep(5000)

                    strResearchComments = "Payloc_timeout_retry"

                    Return strResearchComments

                End If
                '-------------------------------------------------------


                If strResearchComments = "exception" Then
                    Return strResearchComments
                End If
                If InStr(GetText(24, 1, 80), "OFFICE INVALID") > 0 Then
                    Return "E030OFFICE INVALID"
                End If
                If InStr(GetText(24, 1, 80), "W018ADJ SIGN ON COMPLETE") > 0 Then
                    LogInFlag = True
                    WriteLog("Payloc Log in successful...")
                    Dim isSameUser As Boolean = IsDBUpdatedByCurrentUser(ICN, id)

                    If isSameUser = False Then
                        strResearchComments = "ICN picked by different user"
                        Return strResearchComments
                    End If
                Else
                    If LogInCount > 3 Then
                        Exit While
                    End If
                    LogInCount += 1
                    autECLPSObj.StopCommunication()
                    Do

                    Loop While autECLPSObj.CommStarted
                    autECLPSObj.StartCommunication()
                    Do
                        apiChk()
                    Loop Until autECLPSObj.CommStarted
                    Threading.Thread.Sleep(60000)
                End If
            Catch ex As OverflowException
                Threading.Thread.Sleep(60000)
                WriteLog(GetTextRect)
                WriteLog(GetCurrentMethod().Name + " Inside While Not LogInFlag")
                WriteLog(ex.ToString)
                Return ex.Message
                'LogInFlag = False
                FlagError = True
            Catch ex As System.Threading.ThreadAbortException
                WriteLog(GetCurrentMethod().Name + " ThreadAbortException")
                WriteLog(ex.ToString)
                WriteLog(GetTextRect)
                Return ex.Message
                'LogInFlag = False
                FlagError = True
            Catch ex As Exception
                Threading.Thread.Sleep(60000)
                WriteLog(GetCurrentMethod().Name + " Inside While Not LogInFlag")
                WriteLog(GetTextRect)
                WriteLog(ex.ToString)
                'LogInFlag = False
                Return ex.Message
                FlagError = True
            End Try
        End While
        WriteLog("Payloc Log in Ends...")
        'textprevA = GetTextRect(1, 1, 24, 80)
        'textprevB = GetTextRect(1, 1, 24, 80)
        Try
            If FlagError = True Then
                Return strResearchComments
            End If
            If PullEDS2(ICN, "2") = True Then
                apiChk()
                WriteLog(GetTextRect)
                TPvalue = Trim(GetText(16, 14, 1))
                billType = TPvalue
                If TPvalue = "8" Then
                    strResearchComments = "Bill Type 8 skipped"
                    Return "Bill Type 8 skipped"
                End If
                GetClaimDetails(CurrentClaimDetails, ICN, Payloc)
                Dim Res = From row In dt_PG_Policyno.AsEnumerable()
                          Where row.Field(Of String)("Policyno") = CurrentClaimDetails.Policy
                          Select row
                Dim inttt As Int32 = Res.Count()
                If inttt = 1 Then
                    IsPGPolicy = 1
                End If
                If strResearchComments = "Claim is not an electronic/paper submission" Then
                    Return strResearchComments
                End If
                GetEDS1(CurrentICNEDS1, ICN)
                If strResearchComments = "EDS1 NYS Line present" Then
                    Return strResearchComments
                End If
                flagNonCoveredServices = PopulateMatchedNonServiceCodes(CurrentICNEDS1)
                nonServiceCodePresent = flagNonCoveredServices
                'For i As Integer = 0 To CurrentICNEDS1.GetUpperBound(0)            'TO SKIP CLAIMS HAVING REV=762 or HCPC code=G0378/G0379
                '    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                '    For j As Integer = 0 To CurrentICNEDS1(i).ServiceLine.CPTCode.GetUpperBound(0)
                '        If CurrentICNEDS1(i).ServiceLine.CPTCode(j) = "G0378" Or CurrentICNEDS1(i).ServiceLine.CPTCode(j) = "G0379" Then
                '            'strResearchComments = "HCPC code=G0378/G0379"
                '            'Return strResearchComments
                '        End If
                '    Next
                '    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                '    For k As Integer = 0 To CurrentICNEDS1(i).ServiceLine.RevCode.GetUpperBound(0)
                '        If CurrentICNEDS1(i).ServiceLine.RevCode(k) = "0762" Then
                '            strResearchComments = "0762 REVCode present"
                '            Return strResearchComments
                '        End If
                '    Next
                'Next
                FindNotCov()
                GetDXCode(CurrentDXCode)
                MaxSuffix = MaxSuffixCount(ICN, MaxSuffixCountErrorFlag)
                SuffixCountOfCurrentICN = MaxSuffix
                If MaxSuffixCountErrorFlag = True Then
                    Return strResearchComments
                End If
                CurrentOII.GetOIIDetails()
                getOiOIM(isOI, isOIM, CurrentOII)
                Dim isMedicare As Boolean = MedicareCheck(CurrentOII)
                If isMedicare Then
                    strResearchComments = "OI/MEDICARE APPLICABLE FOR UB CLAIMS"
                    Return strResearchComments
                End If
                GETEDS9(CurrentICNEDS9, ICN)
                strICN = CurrentClaimDetails.ICN

                For k As Integer = LBound(CurrentICNEDS9) To UBound(CurrentICNEDS9)
                    For h As Integer = LBound(CurrentICNEDS9(k).LineCount) To UBound(CurrentICNEDS9(k).LineCount)
                        Dim PC As String = ""

                        PC = CurrentICNEDS9(k).PC(h)
                        If InStr(PC, "TH") > 0 Then
                            'rowNum = (10 + (CInt(CurrentICNEDS9(k).LineCount(h)) - 1) * 2).ToString()
                            'If InStr(GetText(rowNum + 1, 6, 80), "RESPONSE RECD FROM SUBR VENDOR - CODE = A") > 0 Then
                            '    Exit For
                            'End If
                            If InStr(CurrentICNEDS9(k).Comment(h), "CODE = A") > 0 Then
                                Exit For
                            End If
                        End If
                        If InStr(PC, "CV") > 0 OrElse InStr(PC, "TH") > 0 OrElse InStr(PC, "BH") > 0 Then
                            'strResearchComments = "COVID Claim"
                            'Return strResearchComments
                        End If
                    Next
                Next
                Dim IsAvPend As Boolean = False
                Dim PendAvProcessing As Boolean = False
                For k As Integer = LBound(CurrentICNEDS9) To UBound(CurrentICNEDS9)
                    For h As Integer = LBound(CurrentICNEDS9(k).LineCount) To UBound(CurrentICNEDS9(k).LineCount)
                        Dim PC As String = ""
                        Dim Comment As String = ""
                        PC = CurrentICNEDS9(k).PC(h)
                        Comment = CurrentICNEDS9(k).Comment(h)
                        If InStr(PC, "AV") > 0 And InStr(Comment, "OVERPAYMENT RECOVERY") > 0 And CurrentICNEDS9(k).LineColor(h) = 0 Then   'Check for AV in pendCode and Comment
                            Return "AV pend present in Multiple Suffix"
                            IsAvPend = True
                            Dim MatchICNList As String = FindMatchICNList(UsrCntrlAdmin)

                            Dim avOverPayment As AVOverpayment = New AVOverpayment(Payloc, MatchICNList, ICN)
                            PendAvProcessing = avOverPayment.AVOverpaymentProcessFlow(IsAvPend)
                            If PendAvProcessing Then
                                Return "Claim process by tool"
                            Else
                                'If AV pend is unsuccessful
                                Return strResearchComments
                            End If
                        End If

                    Next
                Next
                For i As Integer = 0 To CurrentICNEDS1.GetUpperBound(0)
                    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                    For j As Integer = 0 To CurrentICNEDS1(i).ServiceLine.M1.GetUpperBound(0)
                        If CurrentICNEDS1(i).ServiceLine.M1(j) = "50" Then
                            isModifierPresent = True
                        End If
                    Next
                    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                    For k As Integer = 0 To CurrentICNEDS1(i).ServiceLine.M2.GetUpperBound(0)
                        If CurrentICNEDS1(i).ServiceLine.M2(k) = "50" Then
                            isModifierPresent = True
                        End If
                    Next
                Next
                For i As Integer = 0 To CurrentICNEDS1.GetUpperBound(0)
                    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                    For j As Integer = 0 To CurrentICNEDS1(i).ServiceLine.M1.GetUpperBound(0)
                        If CurrentICNEDS1(i).ServiceLine.M1(j) = "57" Then
                            Return "Modifier 57 present"
                        End If
                    Next
                    If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                    For k As Integer = 0 To CurrentICNEDS1(i).ServiceLine.M2.GetUpperBound(0)
                        If CurrentICNEDS1(i).ServiceLine.M2(k) = "57" Then
                            Return "Modifier 57 present"
                        End If
                    Next
                Next
                If isModifierPresent = True And MaxSuffix > 1 Then
                    strResearchComments = "50 modifier present in multiple suffix"
                    Return strResearchComments
                End If
                If CheckClaimAlreadyProcess(MaxSuffix) = True Then

                    strResearchComments = "Already Processed Claim"
                    WriteLog("================ALREADY PROCESSED CLAIM========================")
                    Return strResearchComments
                Else
                    If MaxSuffix > 3 Then
                        strResearchComments = "Multiple Suffix greater than three"
                        Return strResearchComments
                    End If
                    For ii As Integer = 1 To MaxSuffix
                        If GetText(4, 15, 5) <> "EDS 9" Then
                            pullEDSScreen(ICN, 9)

                        End If
                        For ss As Integer = LBound(CurrentICNEDS9) To UBound(CurrentICNEDS9)
                            For dd As Integer = LBound(CurrentICNEDS9(ss).LineCount) To UBound(CurrentICNEDS9(ss).LineCount)
                                Dim status As String = CurrentICNEDS9(ss).Status(dd)
                                Dim LineColor As Integer = CurrentICNEDS9(ss).LineColor(dd)
                                Dim NBR As String = CurrentICNEDS9(ss).NBR(dd)
                                If NBR = String.Empty Then
                                    NBR = "0"
                                End If
                                Dim PC As String = CurrentICNEDS9(ss).PC(dd)
                                'If (LineColor = 0 AndAlso CInt(NBR) = ii AndAlso PC = "CR" AndAlso InStr(AllBusinessPartner1, PC) = 0 AndAlso InStr(AllBusinessPartner2, PC) = 0 AndAlso InStr(AllBusinessPartner3, PC) = 0 AndAlso InStr(AllBusinessPartner4, PC) = 0) OrElse (PC = "Z1") OrElse (PC = "Z3") Then
                                '**** Removed CInt(NBR) = ii condition as per Rajib on 03rd june,24 *********
                                If (LineColor = 0 AndAlso CInt(NBR) = ii AndAlso PC = "CR" AndAlso InStr(AllBusinessPartner1, PC) = 0 AndAlso InStr(AllBusinessPartner2, PC) = 0 AndAlso
                                    InStr(AllBusinessPartner3, PC) = 0 AndAlso InStr(AllBusinessPartner4, PC) = 0) OrElse (PC = "Z1") OrElse (PC = "Z3") Then
                                    LineColorCheck = True
                                    'If ii = 1 Then
                                    SendKeyes($"A{ii}", 1, 76)
                                    SendEnter()
                                    WriteLog(GetTextRect())
                                    If InStr(GetText(24, 1, 80), "E009INVALID CONTROL LINE") > 0 Then Return "E009INVALID CONTROL LINE"
                                    If GetText(2, 2, 2) = "MP" Then
                                        IncreementMPCCount()
                                        'RemovePZ1()
                                        WriteLog(GetTextRect)
                                    End If

                                    If GetText(2, 2, 2) = "AR" Then
                                        strResearchComments = "Auto routed to ARI screen. Valid Notification selection is required"
                                        Return strResearchComments
                                    End If
                                    If InStr(Trim(GetText(24, 2, 79)), "E028NOT AUTHORIZED,") Then
                                        strResearchComments = "E028NOT AUTHORIZED,"
                                        Return strResearchComments
                                    End If
                                    If InStr(Trim(GetText(24, 2, 79)), "CLAIM AWAITING AUTO ADJUDICATION") Then
                                        strResearchComments = "CLAIM AWAITING AUTO ADJUDICATION"
                                        Return strResearchComments
                                    End If
                                    If GetText(2, 2, 3) <> "MPC" Then
                                        'FlagError = True
                                        WriteLog(GetTextRect())
                                        strResearchStatus = "SKIPPED"
                                        strResearchComments = "unable to pull payment screen: " + Trim(GetText(24, 1, 80, True))
                                        Return strResearchComments
                                        Exit Function
                                    End If
                                    If GetText(10, 4, 2, True) = "OI" OrElse GetText(10, 4, 3) = "OIM" OrElse (InStr(PISLLINE, GetText(10, 4, 6).Trim) > 0 Or InStr("OI,OIM", GetText(10, 4, 6, True)) > 0) Then
                                        strResearchComments = "OI/OIM PRESENT"
                                        Return strResearchComments
                                    End If
                                    Dim check39Flag As Boolean = CheckRC39()
                                    If check39Flag = True Then
                                        strResearchComments = "39 or SF present"
                                        Return strResearchComments
                                    End If

                                ElseIf LineColor = 1 Then
                                    'SendKeyes("ret,i" + ICN + ",m", 2, 2)
                                    'SendEnter()
                                    'apiChk()
                                    'Return strResearchComments = "No details found in EDS 9"
                                End If
                            Next
                        Next
                    Next
                    If LineColorCheck = False Then
                        strResearchComments = "No details in EDS 9"
                        Return strResearchComments
                    End If
                    FinalComment = String.Empty
                    FinalComment = WorkONPaymentScreen(MaxSuffix, ICN, Payloc)
                    If FinalComment = String.Empty Then
                        'SendOverrideInScreen("01")
                        ''pullEDSScreen(ICN, "9")
                        'For SFX As Integer = 1 To MaxSuffix
                        '    'SendClear()
                        '    'SendClear()
                        '    pullEDSScreen(CurrentClaimDetails.ICN, 1)
                        '    SendKeyes("A" & SFX, 1, 76)
                        '    SendEnter()
                        '    If GetText(2, 2, 2) = "MP" Then
                        '        SendKeyes("EDS", 2, 2)
                        '        SendEnter()
                        '    End If

                        '    RemoveNotCoveredFromEDS6Point5()
                        '    SendKeyes("A", 1, 75)
                        '    SendEnter()
                        '    SendOverrideInScreen("01")
                        'Next
                        WriteLog(GetTextRect(), True)
                        CheckBypassHistoryField()
                        ReadClaimPaymentScreen(CurrentClaimPaymentScreen, CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount, False, ICN, MaxSuffix, True)
                        ReadPaymentScreen(CurreentICNPaymentScreen, MaxSuffix)
                        CurrentPaymentScreen(PaymentScreenData, MaxSuffix)

#Region "ICES DENIAL and OTHER DENIALS"
                        If flagDenial = True Then
                            implicitDenial(MaxSuffix)
                        End If
                        If flagNonCoveredServices Then
                            implicitDenialInEDS(MaxSuffix)
                        End If
                        '****************iCES Denial********************************
                        ICESDenialFlag = False
                        icesSuccess = False
                        '*****************************Added 11th June,24 by Sourav*********************************

                        ICESCheck(ICN, MaxSuffix)
                        'If InStr(strResearchComments, "ICES denial in multiple suffix") > 0 Then Return strResearchComments
                        If InStr(strResearchComments, "No matching row found in JOC") > 0 Then
                            Return strResearchComments
                        End If
                        '******************************************************************************************
                        If InStr(strResearchComments, "iCES logic applies based on EDS1 FLAG") > 0 Then
                            Return strResearchComments
                        End If
                        If InStr(strResearchComments, "Skip ICES:No Denial RC pulled by system") > 0 Then
                            Return strResearchComments
                        End If
                        If strResearchComments = "Skip ICES finalRC" Then
                            Return strResearchComments
                        End If
                        If ICESDenialFlag = True And icesSuccess = False Then
                            strResearchComments = strResearchComments + " ICES denial failed"
                            Return strResearchComments
                        End If

#Region ">>>>     To check whether all suffixes of claim is fully denied, process the claim accordingly     <<<<<<"

                        Dim dt_SuffixCompletelyDenied As New DataTable("SuffixCompletelyDenied")
                        dt_SuffixCompletelyDenied.Columns.Add("Suffix", GetType(Integer))
                        dt_SuffixCompletelyDenied.Columns.Add("FullDenyApplicable", GetType(Boolean))
                        dt_SuffixCompletelyDenied.Columns.Add("DenySuccessful", GetType(Boolean))
                        If (ICESDenialFlag = True And icesSuccess = True) Or flagDenial = True Or flagNonCoveredServices = True Then                                   'Added 10th Aug,2022
                            Dim isSuffixFullyDenied As Boolean
                            For I As Integer = 1 To MaxSuffix

                                If GetText(6, 2, 5) = "EDS 1" Or GetText(3, 75, 5) = "EDS 2" Then
                                    SendKeyes("A" & I, 1, 76)
                                    SendEnter()
                                Else
                                    pullEDSScreen(ICN, 1)
                                    SendKeyes("A" & I, 1, 76)
                                    SendEnter()
                                    apiChk()
                                End If

                                isSuffixFullyDenied = isDeniedEntireSuffixFromMPCScreen()             'Added 17th June,24    *********
                                dt_SuffixCompletelyDenied.Rows.Add(I, isSuffixFullyDenied, False)
                                If isSuffixFullyDenied = False Then Continue For
                            Next

                            For Each row As DataRow In dt_SuffixCompletelyDenied.Rows
                                Dim suffix As Integer = CInt(row("Suffix"))

                                Dim fullDeny As Boolean = CBool(row("FullDenyApplicable"))
                                If Not fullDeny Then Continue For

                                If fullDeny = True Then
                                    'Navigate to suffix if required
                                    pullEDSScreen(ICN, 1)
                                    SendKeyes("A" & suffix, 1, 76)
                                    SendEnter()
                                    apiChk()

                                    Dim COUNT2 As Integer = 0
                                    Do
                                        SendEnter()
                                        If COUNT2 > 50 Then Exit Do
                                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                        EditWarning = EditWarning & GetText(24, 1, 80, True)
                                        COUNT2 = COUNT2 + 1
                                    Loop
                                    If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                        SendF12()
                                        If CheckClaimProcessedBeforeMPP() = True And suffix = 1 Then
                                            strResearchComments = "Current claim is already processed"
                                            WriteLog(GetTextRect)
                                            Return strResearchComments
                                        End If
                                        SendF12()
                                        SendMPP()
                                        isHistoryICNVoided = False
                                        If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                            'dt.Rows.Add(I, True, False)
                                            WriteLog("MPP failed")
                                            WriteLog(GetTextRect)

                                            strResearchComments = $"MPP failed for suffix {suffix}:  {GetText(24, 1, 80, True)}"
                                            strResearchStatus = "skipped"
                                            FlagError = True
                                            Return strResearchComments
                                        ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then

                                            'Mark this suffix as successfully denied
                                            row("DenySuccessful") = True
                                            WriteLog("Suffix " & suffix & " successfully denied")
                                            WriteLog(GetTextRect)
                                            'strResearchComments = "Claim process by tool"
                                            'PaymentComment = "Claim fully denied"

                                        End If
                                    Else
                                        strResearchComments = EditWarning
                                        strResearchStatus = "skipped"
                                        FlagError = True
                                        Return strResearchComments
                                    End If
                                End If

                            Next

                            Dim allSuffixesDeniedSuccessfully As Boolean = False

                            For Each row As DataRow In dt_SuffixCompletelyDenied.Rows

                                If CBool(row("FullDenyApplicable")) Then

                                    If CBool(row("DenySuccessful")) Then
                                        allSuffixesDeniedSuccessfully = True
                                        'Exit For
                                    End If

                                End If

                            Next

                            If allSuffixesDeniedSuccessfully Then

                                PaymentComment = "All suffixes fully denied"
                                strResearchComments = "Claim process by tool; All suffix processed"

                                WriteLog(strResearchComments)

                                Return strResearchComments

                            End If
#End Region
                            WriteLog("=============================After denial is performed===========================")
                            ReadPaymentScreen(CurrentIcnMPCScreenIfICES, MaxSuffix)

                            'WriteLog("=============================After denial is performed===========================")
                            'ReadPaymentScreen(CurrentIcnMPCScreenIfICES, MaxSuffix)
                        End If
                        '*******************************************************************************************
#End Region
                        '=============Split Mod50===============================
                        If isModifierPresent = True Then
                            If GetText(2, 2, 2) <> "MP" Then
                                SendKeyes("A", 1, 76)
                                SendEnter()
                            End If
                            RemovePZ1()
                            If GetText(2, 2, 2) = "MP" Then
                                SendKeyes("EDS", 2, 2)
                                SendEnter()
                            End If
                            RemoveNotCoveredFromEDS6Point5()
                            goToEDS6Page(1)
                            splitLineOperationFormodifier50InEDS6()
                            If InStr(strResearchComments, "more than 1 lines with 50 modifier") > 0 Then
                                Return strResearchComments
                            End If

                            pullEDSScreen(ICN, 1)
                            GetEDS1(CurrentICNEDS1, ICN)
                            CheckBypassHistoryField()
                            ReadClaimPaymentScreen(CurrentClaimPaymentScreen, CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount, False, ICN, 1, True)
                        End If
                        '====================================================
                        If GetText(10, 29, 2) <> "01" Then
                            SendOverrideInScreen("01")
                            If GetText(2, 2, 2) = "MP" Then
                                SendKeyes("EDS", 2, 2)
                                SendEnter()
                            End If
                            RemoveNotCoveredFromEDS6Point5()
                            If GetText(3, 2, 5) = "EDS 6" Then
                                SendKeyesFor01("A", 1, 75)
                                SendEnter()
                            End If
                        End If
                        '-------- Commented for REDUNDANCY    on 11/3/2026
                        'CheckBypassHistoryField()
                        'ReadClaimPaymentScreen(CurrentClaimPaymentScreen, CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount, False, ICN, MaxSuffix, True)
                        'ReadPaymentScreen(CurreentICNPaymentScreen, MaxSuffix)
                        'CurrentPaymentScreen(PaymentScreenData, MaxSuffix)
                        '----------------------------------------------------------------

                        If GetText(2, 2, 2) <> "MP" Then
                            If GetText(6, 2, 5) <> "EDS 1" Then
                                pullEDSScreen(ICN, 1)
                            End If
                            SendKeyes("A", 1, 76)
                            'apiChk()
                            'autECLPSObj.SendKeys("[ENTER]")
                            'apiChk()
                            SendEnter()
                        End If
                        Dim firstFlag As Boolean = False
                        If GetText(2, 2, 2) = "MP" Then
                            IncreementMPCCount()
                            'RemovePZ1()
                            EditWarning = String.Empty
                            count = 0
                            Do
                                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                If count > 40 Then Exit Do
                                If firstFlag = True Then
                                    EditWarning += GetText(24, 1, 80, True)
                                End If
                                If InStr(GetText(24, 1, 80), "W1616PRV") > 0 And firstFlag = False Then
                                    EditWarning += GetText(24, 1, 80, True)
                                    firstFlag = True
                                End If
                                SendEnter()
                                apiChk()
                                count += 1
                            Loop
                        Else
                            strResearchComments = "Unable to pull MPC screen from " + GetCurrentMethod.Name
                            Return strResearchComments
                        End If

                        Dim TotalCurrentAllowable As Decimal
                        If ICESDenialFlag = True Then
                            TotalCurrentAllowable =
                                CurrentIcnMPCScreenIfICES.
                                    Where(Function(c) c.Allowable IsNot Nothing).
                                    SelectMany(Function(c) c.Allowable).
                                    Sum(Function(a) Convert.ToDecimal(a))
                        Else

                            TotalCurrentAllowable =
                                CurrentClaimPaymentScreen.
                                    Where(Function(c) c.Allowable IsNot Nothing).
                                    SelectMany(Function(c) c.Allowable).
                                    Sum(Function(a) Convert.ToDecimal(a))
                        End If

                        If TotalCurrentAllowable = 0.0 Then
                            Return strResearchComments = "Current claim Allowable not present in MPC"
                        End If

                        Try
                            WriteLog("Edit warning: " + EditWarning + vbNewLine)
                            Dim value1 As String
                            Dim value2 As String
                            Dim value3 As String
                            Dim value4 As String
                            Dim matchSuccess As Boolean = False

                            If InStr(EditWarning, "W1616PRV") > 0 Then
                                '---------------------------------------------------------
                                ''Dim myString1, myString2, myString3, myString4 As String
                                'Dim indx1 As Integer = InStr(EditWarning, "PRV") + 3
                                'Dim indx2 As Integer = EditWarning.Length - indx1
                                'Dim txt As String = EditWarning.Substring(indx1) ', indx2)
                                ''Dim indx3 As Integer = InStr(txt, ",W1") Or InStr(txt, ",W2") Or InStr(txt, ",W3") Or InStr(txt, ",W4") Or InStr(txt, ",W5") Or InStr(txt, ",W6") Or InStr(txt, ",W7") Or InStr(txt, ",W8") Or InStr(txt, ",W9") Or InStr(txt, ",W0") Or InStr(txt, ",E1") Or InStr(txt, ",E2") Or InStr(txt, ",E3") Or InStr(txt, ",E4") Or InStr(txt, ",E5") Or InStr(txt, ",E6") Or InStr(txt, ",E7") Or InStr(txt, ",E8") Or InStr(txt, ",E9") Or InStr(txt, ",E0") Or InStr(txt, ",WE1") Or InStr(txt, ",WE2") Or InStr(txt, ",WE3") Or InStr(txt, ",WE4") Or InStr(txt, ",E5") Or InStr(txt, ",WE6") Or InStr(txt, ",WE7") Or InStr(txt, ",WE8") Or InStr(txt, ",WE9") Or InStr(txt, ",WE0") Or InStr(txt, ",EW1") Or InStr(txt, ",EW2") Or InStr(txt, ",EW3") Or InStr(txt, ",EW4") Or InStr(txt, ",EW5") Or InStr(txt, ",EW6") Or InStr(txt, ",EW7") Or InStr(txt, ",EW8") Or InStr(txt, ",EW9") Or InStr(txt, ",EW0")
                                ''MarketIDLine = txt.Substring(0, indx3 - 1)
                                'If InStr(txt, "W1616PRV") > 0 And (Left(txt, 3) <> "EPO" And Left(txt, 3) <> "POS" And Left(txt, 3) <> "HMO") Then
                                '    Dim ind1 As Integer = InStr(txt, "PRV") + 3
                                '    Dim ind2 As Integer = txt.Length - ind1
                                '    Dim txt2 As String = txt.Substring(ind1)
                                '    MarketIDLine = txt2
                                'Else
                                '    MarketIDLine = txt
                                'End If
                                '---------------------------------------------------------
                                Dim searchString As String = "W1616PRV"
                                Dim validFirstValues As String() = {"IND", "MIN", "PP1", "PPO", "EPO", "POS", "EP1", "PS1", "HM1", "HM2", "HM3", "HM4", "HM5", "HM6", "HM7", "HM8", "HM9", "HMA", "HMS", "HML", "GIL", "HMS", "GIS", "HMP", "HMO"}

                                Dim startIndex As Integer = EditWarning.IndexOf(searchString)
                                While startIndex <> -1
                                    Dim substring As String = EditWarning.Substring(startIndex + searchString.Length).Trim()
                                    Dim values As String() = substring.Split(",c")

                                    If values.Length >= 4 AndAlso validFirstValues.Contains(values(0).Trim()) And ((values(3).Trim()).Length = 5 Or (values(3).Trim()).Contains(" ")) Then
                                        Dim result As String() = values.Take(4).Select(Function(v) v.Trim()).ToArray()

                                        value1 = result(0)
                                        If result(1).Contains(" ") Then
                                            value2 = result(1).Split(" ")(1)
                                        Else
                                            value2 = result(1)
                                        End If
                                        If result(2).Contains(" ") Then
                                            value3 = result(2).Split(" ")(1)
                                        Else
                                            value3 = result(2)
                                        End If
                                        If result(3).Contains(" ") Then
                                            value4 = result(3).Split(" ")(1)
                                        Else
                                            value4 = result(3)
                                        End If
                                        matchSuccess = True
                                    End If

                                    startIndex = EditWarning.IndexOf(searchString, startIndex + searchString.Length)
                                End While

                                If matchSuccess = False Then
                                    WriteLog("Value extraction failed from Editwarning")
                                    strResearchComments = "Value extraction failed from Editwarning"
                                    Return "Value extraction failed from Editwarning"
                                End If
                            Else
                                WriteLog("W1616PRV not present in editline")
                                strResearchComments = "W1616PRV not present in editline"
                                Return "W1616PRV not present in editline"
                            End If
                            Dim checkClaim As String
                            checkClaim = IHAndOHSegrigation()
                            If checkClaim = "IH" Then
                                isCurrentClaimIH = True
                            ElseIf checkClaim = "OH" Then
                                isCurrentClaimOH = True
                            End If
                            Dim Product As String = value1
                            WriteLog("Product= " + Product)

                            'MarketId = Trim(arrMarketIDLine(2)).Substring(0, 5)

                            'MarketId Logic is changed on 10th Aug,2021
                            Dim marketIDArray() As String
                            'MarketId = Trim(arrMarketIDLine(2))
                            MarketId = value3
                            WriteLog("MarketId = " + MarketId)

                            Dim Type As String = value2
                            WriteLog("Type = " + Type)
                            If Integer.Parse(Type) = 57 Or Integer.Parse(Type) = 58 Or Integer.Parse(Type) = 61 Then    'Added 9th Nov,26 to skip harvard claims
                                strResearchComments = "Skipping Harvard pilgrim claims"
                                Return strResearchComments
                            End If
                            If Not (Integer.Parse(Type) >= 10 And Integer.Parse(Type) <= 99) And Type <> "00" Then
                                strResearchComments = "MPC screen: Error in extraction from Edit"
                                Return strResearchComments
                            End If
                            Dim Ipa As String = value4
                            WriteLog("Ipa = " + Ipa)

                            Dim DCC As String = CurrentClaimDetails.DCC
                            Dim FLN As String = CurrentClaimDetails.FLN
                            Dim Fdate As String = CurrentClaimDetails.Fdate
                            Dim ProvTin As String = CurrentClaimDetails.Prifix_Tin_Suffix
                            Dim DOBEDS1 As String = CurrentClaimDetails.PatDOB
#Region "Timely filing"
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>Code for Timely filing<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                            WriteLog(">>>>>>Timely filing check starts<<<<<<<")
                            Dim locality, AdjId As String
                            If boolOffshore Then
                                locality = "Offshore"
                                AdjId = "398676"
                            Else
                                locality = "Onshore"
                                AdjId = "430664"
                            End If
                            If TPvalue = "7" Then
                                Dim MatchICNList As String = FindMatchICNList(UsrCntrlAdmin)
                                If InStr(strResearchComments, "Remark Code S6 present in HxICN") > 0 Then Return "Remark Code S6 present in HxICN"
                                If InStr(strResearchComments, "Remark Code UP/LP/EW present in HxICN") > 0 Then Return "Remark Code UP/LP/EW present in HxICN"
                                If InStr(strResearchComments, "Less than 50%") > 0 Then Return strResearchComments
                                If InStr(strResearchComments, "Exception") > 0 Then Return strResearchComments
                                Dim hxFilingDate = DetermineHxFilingDate(MatchICNList)
                                If hxFilingDate = "" Then
                                    WriteLog("Hx filing date not found/matched")
                                End If
                                Dim daysDiff As String = ""
                                Dim limit As String = ""
                                If productCodeDict.Count = 0 Then
                                    initializeNDBDict()
                                End If
                                Dim currentFilingDate As Date = ConvertJuliantoDate((CurrentClaimDetails.FLN).Substring(0, 5))
                                WriteLog("currentFilingDate= " + currentFilingDate.ToString)
                                Dim tf As TimelyFiling = New TimelyFiling(ICN, currentFilingDate, hxFilingDate)
                                daysDiff = tf.GetDaysDifference()
                                WriteLog("CurrentFilingDate - HxDate - 1 =" + daysDiff.ToString)
                                Dim isNDBLogIn = tf.InitializeNDBLogOn(ndbsession)
                                If isNDBLogIn = False Then
                                    strResearchComments = "NDB login failed"
                                    Return strResearchComments
                                End If

                                Dim service As New ApiService()
                                limit = service.CallPpiPriApi(ICN, Payloc, Engine, UserID, AdjId, locality, ProvTin, MarketId, Type, Product, Ipa, Fdate)
                                If limit = "" Then
                                    limit = tf.CheckNDBFilingLimit(CurrentClaimDetails.Tin, CurrentClaimDetails.Prefix_Tin.Substring(0, 1),
                                                                                   CurrentClaimDetails.Prifix_Tin_Suffix.Substring(10, 5), Product, Type, MarketId, Ipa, ndbsession)
                                End If
                                WriteLog($"Timely filing limit:{limit}")
                                If limit = "" Then
                                    strResearchComments = "No matched rows from NDB"
                                    Return strResearchComments
                                End If

                                WriteLog("If " + daysDiff.ToString + " > " + limit.ToString())
                                If CInt(daysDiff) > CInt(limit) Then
                                    For SFX As Integer = 1 To MaxSuffix
                                        'If GetText(2, 2, 2) <> "MP" Then
                                        'End If
                                        SendClear()
                                        SendClear()
                                        pullEDSScreen(CurrentClaimDetails.ICN, 1)
                                        SendKeyes("A" & SFX, 1, 76)
                                        SendEnter()
                                        If GetText(2, 2, 2) = "MP" Then
                                            SendKeyes("EDS", 2, 2)
                                            SendEnter()
                                        End If
                                        'code to deny claim with RC =  7G
                                        denyWith7G()
                                        Dim COUNT1 As Integer = 0
                                        Do
                                            SendEnter()
                                            If COUNT1 > 50 Then Exit Do
                                            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                            EditWarning = EditWarning & GetText(24, 1, 80, True)
                                            COUNT1 = COUNT1 + 1
                                        Loop
                                        If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                            'SendF12()
                                            'If CheckClaimProcessedBeforeMPP() = True Then
                                            '    strResearchComments = "Current claim is already processed"
                                            '    Return strResearchComments
                                            'End If
                                            'SendF12()
                                            SendMPP()
                                            WriteLog(GetTextRect)
                                            isHistoryICNVoided = False
                                            If SFX = MaxSuffix Then
                                                If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                                    strResearchComments = GetText(24, 1, 80, True)
                                                    strResearchStatus = "skipped"
                                                    WriteLog($"Suffix-{SFX} T8 deny failed")
                                                    FlagError = True
                                                    Return strResearchComments
                                                ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                                    strResearchComments = "Claim process by tool"
                                                    PaymentComment = "Claim denied with T8"
                                                    Return strResearchComments
                                                End If
                                            End If
                                        Else
                                            strResearchComments = $"T8 DENY FAILED FOR SUFFIX{SFX}:  {EditWarning}"
                                            strResearchStatus = "skipped"
                                            FlagError = True
                                            Return strResearchComments
                                        End If
                                    Next
                                End If
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
#End Region

#Region "0762 REV code"
                            flag0762 = False
                            For i As Integer = 0 To CurrentICNEDS1.GetUpperBound(0)            'TO SKIP CLAIMS HAVING REV=762 or HCPC code=G0379
                                If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                                For j As Integer = 0 To CurrentICNEDS1(i).ServiceLine.CPTCode.GetUpperBound(0)
                                    If CurrentICNEDS1(i).ServiceLine.CPTCode(j) = "G0379" Then
                                        strResearchComments = "No HCPC code from HardCopy// G0379"
                                        Return strResearchComments
                                    End If
                                Next
                                If CurrentICNEDS1(i).FDate = Nothing Then Exit For
                                For k As Integer = 0 To CurrentICNEDS1(i).ServiceLine.RevCode.GetUpperBound(0)
                                    If CurrentICNEDS1(i).ServiceLine.RevCode(k) = "0762" Then
                                        flag0762 = True
                                    End If
                                Next
                            Next
                            '===============================
                            If flag0762 Then
                                If MaxSuffix > 1 Then Return strResearchComments = "0762 Rev code in multiple suffix"
                                Dim revClass1 As rev0762 = New rev0762(ICN)
                                revClass1.g0378Flow(dt_762Denial, CurrentICNEDS1)
                                If strResearchComments = "No HCPC code from HardCopy" Then Return strResearchComments
                                Dim isClaimFullyDenied As Boolean = False
                                For I As Integer = 1 To MaxSuffix

                                    If GetText(6, 2, 5) = "EDS 1" Or GetText(3, 75, 5) = "EDS 2" Then
                                        SendKeyes("A" & I, 1, 76)
                                        SendEnter()
                                    Else
                                        pullEDSScreen(ICN, 1)
                                        SendKeyes("A" & I, 1, 76)
                                        SendEnter()
                                        apiChk()
                                    End If
                                    isClaimFullyDenied = isDeniedEntireSuffixFromMPCScreen()             'Added 1st Dec,24    *********
                                    If isClaimFullyDenied = True Then
                                        Dim COUNT2 As Integer = 0
                                        Do
                                            SendEnter()
                                            If COUNT2 > 50 Then Exit Do
                                            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                            EditWarning = EditWarning & GetText(24, 1, 80, True)
                                            COUNT2 = COUNT2 + 1
                                        Loop
                                        If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                            SendF12()
                                            If CheckClaimProcessedBeforeMPP() = True Then
                                                strResearchComments = "Current claim is already processed"
                                                Return strResearchComments
                                            End If
                                            SendF12()
                                            SendMPP()
                                            isHistoryICNVoided = False
                                            If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                                strResearchComments = GetText(24, 1, 80, True)
                                                strResearchStatus = "skipped"
                                                FlagError = True
                                                Return strResearchComments
                                            ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                                strResearchComments = "Claim process by tool"
                                                PaymentComment = "Claim fully denied"
                                                Return strResearchComments
                                            End If
                                        Else
                                            strResearchComments = EditWarning
                                            strResearchStatus = "skipped"
                                            FlagError = True
                                            Return strResearchComments
                                        End If
                                    End If
                                Next
                            End If
                            '===============================
#End Region
                            InitializeEmulator(session)
                            If checkClaim = "IH" Then
                                isCurrentClaimIH = True
                                CalculateIRI(UsrCntrlAdmin, ICN, Engine, ProvTin, Fdate, Product, Type, MarketId, Ipa, locality, AdjId, UserID)
                                If InStr(Trim(GetText(24, 3, 78)), "E") > 0 Then
                                    IRIAllowable = "0"
                                Else
                                    IRIAllowable = Trim(GetText(6, 67, 11))
                                    If IRIAllowable = "--------.--" Then
                                        IRIAllowable = "0"
                                    End If
                                End If
                                allowableFromOCI = IRIAllowable
                            Else
                                isCurrentClaimOH = True
                                Dim objOci As New clsOciAllowable()
                                Dim WhethergetOCIAllowable As Boolean
                                SendClear()
                                SendClear()
                                'UsrCntrlAdmin.OCI(ICN, DCC, FLN, Fdate, ProvTin, Product, Type, MarketId, Ipa, "", DOBEDS1)
                                'objOci.getOCIAllowablen(UsrCntrlAdmin, ICN, DCC, FLN, Fdate, ProvTin, Product, Type, MarketId, Ipa, "", DOBEDS1)
                                WhethergetOCIAllowable = objOci.getOCIAllowable(UsrCntrlAdmin, ICN, Engine, DCC, FLN, Fdate, ProvTin, Product, Type, MarketId, Ipa, "", DOBEDS1, locality, AdjId, UserID)
                                WriteLog(GetTextRect())

                                If InStr(strResearchComments, "Object reference not set") > 0 Then
                                    Return "Object reference not for OCI API"
                                End If
                                If WhethergetOCIAllowable = False Then
                                    strResearchComments = strResearchComments + "/ Allowable N.A."
                                    Return strResearchComments
                                End If
                                If InStr(Trim(GetText(24, 3, 78)), "E") > 0 Then
                                    OCIAllowable = "0"
                                Else
                                    OCIAllowable = Trim(GetText(22, 62, 9))
                                    If OCIAllowable = "------.--" Then
                                        OCIAllowable = "0"
                                    End If
                                End If
                                allowableFromOCI = OCIAllowable

                            End If

                            Dim PSAllowable As Decimal = 0.0
                            Dim PSPaymentAmount As Decimal = 0.0
                            If ICESDenialFlag = True Then
                                For JJ As Integer = LBound(CurrentIcnMPCScreenIfICES) To UBound(CurrentIcnMPCScreenIfICES)

                                    For hh As Integer = LBound(CurrentIcnMPCScreenIfICES(JJ).Allowable) To UBound(CurrentIcnMPCScreenIfICES(JJ).Allowable)

                                        PSAllowable += CDec(CurrentIcnMPCScreenIfICES(JJ).Allowable(hh))
                                        PSPaymentAmount += CDec(CurrentIcnMPCScreenIfICES(JJ).Payment(hh))

                                    Next
                                Next
                            Else

                                For JJ As Integer = LBound(CurreentICNPaymentScreen) To UBound(CurreentICNPaymentScreen)

                                    For hh As Integer = LBound(CurreentICNPaymentScreen(JJ).Allowable) To UBound(CurreentICNPaymentScreen(JJ).Allowable)

                                        PSAllowable += CDec(CurreentICNPaymentScreen(JJ).Allowable(hh))
                                        PSPaymentAmount += CDec(CurreentICNPaymentScreen(JJ).Payment(hh))

                                    Next
                                Next
                            End If

                            Dim IsRCCA As Integer = 0
                            Dim FlagNotMatch As Boolean = False
                            Dim AfterANPIMatchList As String = String.Empty
                            If checkClaim = "IH" Then

                                If CDec(IRIAllowable) = PSAllowable Or flag_QC = True Then

                                    Dim MatchICNList As String = FindMatchICNList(UsrCntrlAdmin)
                                    If InStr(strResearchComments, "FindMatchICNList") Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "Remark Code UP/LP/EW present in HxICN") > 0 Then Return "Remark Code UP/LP/EW present in HxICN"
                                    If InStr(MatchICNList, "W062NO CLAIMS ON FILE") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "HH or IJ present in HxICN") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "OI present in HxICN") > 0 Then Return "OI present in HxICN"
                                    If InStr(strResearchComments, "Less than 50%") > 0 Then
                                        Return strResearchComments
                                    End If
                                    dt_ListOfHxMatchingANPI = New DataTable()
                                    GenerateDt()
                                    MatchAttendingPhysician(MatchICNList, FlagNotMatch)
                                    If TPvalue = "1" Then

                                        For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                            If Row("ANPIMatchFlag") = False Then
                                                Return "ANPI not match for HxICN:" + Row("HxICN")
                                            End If
                                        Next
                                        AfterANPIMatchList = MatchICNList
                                    ElseIf TPvalue = "7" Then
                                        For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                            If Row("ANPIMatchFlag") = True Then
                                                AfterANPIMatchList = AfterANPIMatchList + "," + Row("HxICN")
                                            End If
                                            If Row("ANPIMatchFlag") = False Then
                                                strResearchComments = "ANPI not match for HxICN:" + Row("HxICN")
                                                Return "ANPI not match for HxICN:" + Row("HxICN")
                                            End If
                                        Next
                                    Else
                                        Return "Other BILLTYPE"
                                    End If
                                    SendClear()
                                    SendClear()
                                    Dim isPresentDenialinHXClaim As Boolean = checkForDenialRCinHxClaim(MatchICNList)
                                    If isPresentDenialinHXClaim = True Then
                                        Return "HX ICN denied with iCES/CES Remark Code"
                                    End If
                                    Dim therapyNBRError As String = String.Empty
                                    If MatchICNList <> String.Empty And TPvalue = "7" Then
                                        If isNCopayNBR = True Then
                                            'NCopayNBR = GetCopayNBRfromHX(AfterANPIMatchList, ICN)
                                            NCopayNBR = GetCopayNBRforTherapyClaims(therapyNBRError)
                                        End If
                                        If therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS" Then
                                            Return therapyNBRError
                                        End If
                                        AllHistoryPerform(AfterANPIMatchList)
                                        FinalHistoryICNList = AfterANPIMatchList
                                        WriteLog("strResearchComments:" + strResearchComments)
                                        If InStr(strResearchComments, "Data mismatch in allowable and total paid in Current Claim") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "Overpayment in multiple suffix") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "W342 NEW COINS REVERSAL") > 0 Then Return strResearchComments
                                        If isHistoryICNVoided = True And InStr(strResearchComments, "Claim process by tool") = 0 Then
                                            WriteLog($"History ICN voided={isHistoryICNVoided}")
                                            Return strResearchComments
                                        End If
                                        If InStr(strResearchComments, "Claim process by tool") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "OI/OIM") > 0 Then
                                            Return strResearchComments
                                        End If
                                        If InStr(strResearchComments, "Index out of range") > 0 Or strResearchComments = "Hist ICN: Unable to pull MHI" Then
                                            Return strResearchComments
                                        End If
                                        If InStr(strResearchComments, "ROV done") > 0 Then Return strResearchComments
                                    ElseIf MatchICNList = String.Empty And TPvalue = "1" Then
                                        If isNCopayNBR = True Then
                                            NCopayNBR = GetCopayNBRforTherapyClaims(therapyNBRError)
                                        End If
                                        If therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS" Then
                                            Return therapyNBRError
                                        End If
                                        Dim bt1ErrorFlag As Boolean = False
                                        If CDec(allowableFromOCI) > 5000 Then
                                            IsHDClaim = True
                                        End If
                                        If ICESDenialFlag Then
                                            placingiCESDenialInfo(CurrICN, SuffixCountOfCurrentICN)
                                        End If
                                        ProcessClaimAsFreshSubmission(CurrICN, bt1ErrorFlag)
                                        If InStr(strResearchComments, "Claim process by tool") > 0 Then Return strResearchComments
                                        If bt1ErrorFlag = True Then Return strResearchComments
                                    ElseIf MatchICNList <> String.Empty And TPvalue = "1" Then
                                        Return "Manual intervention required: HARD EDIT PRESENT BT"
                                    Else
                                        strResearchComments = "History ICN not found"
                                        Return strResearchComments
                                    End If

                                    'HistoryClaimDetails(HistoryICNDetails, MatchICNList)

                                    'BasicHistoryClaimDetails(BasicHistoryICNDetails, MatchICNList)

                                    'Dim checkHXClaim As String = CalimDetermination()
                                Else
                                    strResearchComments = "IRI Allowable is not match"
                                    Return strResearchComments
                                End If

                            Else

                                If CDec(OCIAllowable) = PSAllowable Then
                                    Dim MatchICNList As String = FindMatchICNList(UsrCntrlAdmin)
                                    If InStr(MatchICNList, "W062NO CLAIMS ON FILE") > 0 And TPvalue = "1" Then
                                        MatchICNList = ""
                                    End If
                                    If InStr(strResearchComments, "HH or IJ present in HxICN") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "Remark Code S6 present in HxICN") > 0 Then Return "Remark Code S6 present in HxICN"
                                    If InStr(strResearchComments, "Remark Code UP/LP/EW present in HxICN") > 0 Then Return "Remark Code UP/LP/EW present in HxICN"
                                    If InStr(strResearchComments, "OI present in HxICN") > 0 Then Return "OI present in HxICN"
                                    If InStr(strResearchComments, "Less than 50%") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "W062NO CLAIMS ON FILE") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If strResearchComments = "Hard copy ICN is not found" Then
                                        Return strResearchComments
                                    End If
                                    If InStr(strResearchComments, "exception") > 0 Then
                                        Return strResearchComments
                                    End If
                                    If TPvalue <> "1" Then
                                        dt_ListOfHxMatchingANPI = New DataTable()
                                        GenerateDt()
                                        MatchAttendingPhysician(MatchICNList, FlagNotMatch)
                                    End If

                                    If TPvalue = "1" Then

                                        For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                            If Row("ANPIMatchFlag") = False Then
                                                Return "ANPI not match for HxICN:" + Row("HxICN")
                                            End If
                                        Next
                                        AfterANPIMatchList = MatchICNList
                                    ElseIf TPvalue = "7" Then

                                        For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                            If Row("ANPIMatchFlag") = True Then
                                                AfterANPIMatchList = AfterANPIMatchList + "," + Row("HxICN")
                                            End If
                                            If Row("ANPIMatchFlag") = False Then
                                                strResearchComments = "ANPI not match for HxICN:" + Row("HxICN")
                                                Return "ANPI not match for HxICN:" + Row("HxICN")
                                            End If
                                        Next
                                    Else
                                        Return "Other BILLTYPE"
                                    End If
                                    SendClear()
                                    SendClear()
                                    Dim isPresentDenialinHXClaim As Boolean = checkForDenialRCinHxClaim(MatchICNList)
                                    If isPresentDenialinHXClaim = True Then
                                        Return "HX ICN denied with iCES/CES Remark Code"
                                    End If
                                    Dim therapyNBRError As String = String.Empty
                                    If MatchICNList <> String.Empty And TPvalue = "7" Then  ''''BILLTYPE 7
                                        If isNCopayNBR = True Then
                                            NCopayNBR = GetCopayNBRforTherapyClaims(therapyNBRError)
                                        End If
                                        If therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS" Then
                                            Return therapyNBRError
                                        End If
                                        AllHistoryPerform(AfterANPIMatchList)
                                        FinalHistoryICNList = AfterANPIMatchList
                                        WriteLog("strResearchComments:" + strResearchComments)
                                        If InStr(strResearchComments, "Data mismatch in allowable and total paid in Current Claim") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "Overpayment in multiple suffix") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "W342 NEW COINS REVERSAL") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "History voided by different Adjuster") > 0 Then Return strResearchComments
                                        If isHistoryICNVoided = True And InStr(strResearchComments, "Claim process by tool") = 0 Then
                                            'Dim isSuccess = ProtectPriorityTwoClaims(strDraFT, AfterANPIMatchList)
                                            'pullEDSScreen(CurrICN, 1)
                                            'SendKeyes("A", 1, 76)
                                            'SendEnter()
                                            'REMOVE01IFSKIPPED(CurrICN, MaxSuffix)
                                            'Dim EditWarningP2 As String = ""
                                            'Dim COUNT2 As Integer = 0
                                            'Dim flagOktopay As Boolean = False
                                            'Do
                                            '    SendEnter()
                                            '    If COUNT2 > 50 Then Exit Do
                                            '    If InStr(EditWarningP2, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                            '    EditWarningP2 = EditWarningP2 & GetText(24, 1, 80, True)
                                            '    COUNT2 = COUNT2 + 1
                                            'Loop
                                            'If (InStr(EditWarningP2, "E2492") > 0 Or InStr(EditWarningP2, "W2172CORRECTED BILLING") > 0) And isSuccess = True Then
                                            '    isHistoryICNVoided = False  'Turning this as false to send the ICN with flag 4
                                            '    Return "Priority 2 claim protected"
                                            'End If
                                            'If isSuccess = False Then Return strResearchComments + "//Priority 2 claim protection failed"
                                            WriteLog($"History ICN voided={isHistoryICNVoided}")
                                            Return strResearchComments
                                        End If
                                        'ALLHstoryPerform(MatchICNList)
                                        If InStr(strResearchComments, "OI/OIM") > 0 Then
                                            Return strResearchComments
                                        End If
                                        If InStr(strResearchComments, "Claim process by tool") > 0 Then Return strResearchComments
                                        If InStr(strResearchComments, "ROV done") > 0 Then Return strResearchComments
                                    ElseIf MatchICNList = String.Empty And TPvalue = "1" Then     '''' BILLTYPE 1
                                        WriteLog("Billtype 1 and no hx ICNs")
                                        If isNCopayNBR = True Then
                                            NCopayNBR = GetCopayNBRforTherapyClaims(therapyNBRError)
                                        End If
                                        If therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS" Then
                                            Return therapyNBRError
                                        End If
                                        Dim bt1ErrorFlag As Boolean = False
                                        If CDec(allowableFromOCI) > 5000 Then
                                            IsHDClaim = True
                                        End If
                                        If ICESDenialFlag Then
                                            placingiCESDenialInfo(CurrICN, SuffixCountOfCurrentICN)
                                        End If
                                        ProcessClaimAsFreshSubmission(CurrICN, bt1ErrorFlag)
                                        If InStr(strResearchComments, "Claim process by tool") > 0 Then Return strResearchComments
                                        If bt1ErrorFlag = True Then Return strResearchComments
                                    ElseIf MatchICNList <> String.Empty And TPvalue = "1" Then
                                        Return "Manual intervention required:Suffix greater than one BT1"
                                    Else
                                        strResearchComments = "History ICN not found"
                                        Return strResearchComments
                                    End If
                                Else
                                    strResearchComments = "OCI Allowable is not match"
                                    WriteLog(GetCurrentMethod.Name() + " OCI allowable is not match")
                                    WriteLog(GetTextRect)
                                    Return strResearchComments
                                End If

                            End If
                        Catch e As IndexOutOfRangeException
                            WriteLog("Index out of range exception")
                            strResearchComments = GetCurrentMethod.Name() + " method: Index out of range"
                            Return strResearchComments
                            Exit Function
                        Catch e As System.InvalidOperationException
                            strResearchComments = "Invalid Operation"
                            Return strResearchComments
                        Catch ex As Exception
                            strResearchComments = "Exception from " + GetCurrentMethod.Name + ": " + ex.Message.ToString
                            Return strResearchComments
                        End Try


                    Else  '(part of If FinalComment = String.Empty)

                        Return strResearchComments

                    End If

                    If LineColorCheck = False Then
                        strResearchComments = "No details in EDS 9"
                    End If

                End If

            Else
                If strResearchComments = "SELECTED CLAIM NOT FOUND" Then
                    Return strResearchComments
                Else
                    strResearchComments = "Claim is not Corrected Claim"
                End If

            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name() + ": " + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name() + " method: exception"
            Return strResearchComments
        End Try


        Return strResearchComments
    End Function
    Public Sub GetSixPointOriginalValue()

    End Sub
    Public Sub GetDXCode(ByRef D() As STDXCODE)
        Try
            SendKeyes("dx", 1, 76)
            SendEnter()
            apiChk()
            WriteLog("=================OPEN DX SCREEN=========================")
            WriteLog(GetTextRect)
            ReDim D(0)
            Dim LineCount As Integer = 0
            Dim x As Integer = 0
            For dd As Integer = 7 To 11 Step 1
                For hh As Integer = 6 To 64 Step 13
                    Dim checkLine As String = String.Empty
                    checkLine = Trim(GetText(dd, hh, 7))
                    If checkLine = "-------" Then Exit For
                    If x = 0 Then
                        ReDim D(LineCount).DXCode(x)
                    Else
                        ReDim Preserve D(LineCount).DXCode(x)
                    End If
                    D(LineCount).DXCode(x) = Trim(GetText(dd, hh, 7))
                    x += 1
                Next
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try
    End Sub
    Public Function MaxSuffixCount(ByVal icn As String, ByRef errorFlag As Boolean) As Integer
        Dim i As Integer
        Dim CheckLine As String
        Dim Suffix As Integer
        'Suffix = String.Empty
        CheckLine = String.Empty
        Try
            pullEDSScreen(icn, "1")
            For i = 8 To 21 Step 1
                CheckLine = Trim(GetText(i, 2, 4))
                If IsNumeric(CheckLine) = False And CheckLine <> "" Then
                    errorFlag = True
                    strResearchComments = "Insufficient data in EDS1 screen"
                    WriteLog("Error from " + GetCurrentMethod.Name + " in EDS1")
                    Return Suffix
                End If
                If CheckLine <> Nothing Then
                    Suffix = Trim(GetText(i, 60, 2))
                Else
                    Return Suffix
                End If
                If i = 21 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO ADDITIONAL LINES TO DISPLAY") > 0 Then Exit For
                    SendF8()
                    apiChk()
                    i = 7
                End If
            Next

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try
        Return Suffix
    End Function
    Public Sub GetClaimDetails(ByRef MyclaimDetails As ClaimDetails, ByVal icn As String, Payloc As String)
        Try
            ReDimension(CurrentClaimDraftDetails, 0)
            ReDim CurrentClaim(0)

            SendClear()
            SendClear()
            pullEDSScreen(icn, "1")
            apiChk()
            apiChk()
            If Trim(GetText(6, 2, 5)) = "EDS 1" Then
                WriteLog("====================OPEN EDS1 TO STORE DATA==================")
                WriteLog(GetTextRect)
                If Trim(GetText(7, 2, 14)) = "BILLED CHARGES" Then
                    SendF9()
                    MyclaimDetails.ICN = Trim(GetText(1, 6, 10))
                    MyclaimDetails.SSN = Trim(GetText(1, 22, 10))
                    MyclaimDetails.Policy = Trim(GetText(2, 6, 6))
                    MyclaimDetails.PatientName = Trim(GetText(1, 37, 11))
                    MyclaimDetails.PatientRelationship = Trim(GetText(1, 52, 2))
                    MyclaimDetails.Fdate = Trim(GetText(3, 21, 6))
                    MyclaimDetails.Ldate = Trim(GetText(3, 33, 6))
                    MyclaimDetails.PatDOB = Trim(GetText(1, 59, 8))
                    MyclaimDetails.TotalCharge = Trim(GetText(2, 53, 10))
                    CurrentClaim(0).Policy = GetText(2, 6, 6)
                    CurrentClaim(0).SSN = GetText(1, 22, 10)
                    CurrentClaim(0).EmployeeName = GetText(1, 38, 10)
                    CurrentClaim(0).Relationship = GetText(1, 52, 2)
                ElseIf Trim(GetText(6, 2, 5)) = "EDS 1" And Trim(GetText(7, 2, 4)) = "REVE" Then
                    MyclaimDetails.ICN = Trim(GetText(1, 6, 10))
                    MyclaimDetails.SSN = Trim(GetText(1, 22, 10))
                    MyclaimDetails.Policy = Trim(GetText(2, 6, 6))
                    MyclaimDetails.PatientName = Trim(GetText(1, 37, 11))
                    MyclaimDetails.PatientRelationship = Trim(GetText(1, 52, 2))
                    MyclaimDetails.Fdate = Trim(GetText(3, 21, 6))
                    MyclaimDetails.Ldate = Trim(GetText(3, 33, 6))
                    MyclaimDetails.PatDOB = Trim(GetText(1, 59, 8))
                    MyclaimDetails.TotalCharge = Trim(GetText(2, 53, 10))
                    CurrentClaim(0).Policy = GetText(2, 6, 6)
                    CurrentClaim(0).SSN = GetText(1, 22, 10)
                    CurrentClaim(0).EmployeeName = GetText(1, 37, 10)
                    CurrentClaim(0).Relationship = GetText(1, 52, 2)
                End If
            End If
            SendClear()
            SendClear()
            pullEDSScreen(icn, "3")
            WriteLog(GetTextRect())
            MyclaimDetails.FLN = Trim(GetText(2, 17, 10))
            MyclaimDetails.DCC = Trim(GetText(2, 28, 3))
            'Below condition of DCC added on 8/10/2021
            If MyclaimDetails.DCC = "951" Or MyclaimDetails.DCC = "807" Or MyclaimDetails.DCC = "910" Or MyclaimDetails.DCC = Payloc Then
                strResearchComments = "Claim is not an electronic/paper submission"
                Exit Sub
            End If
            MyclaimDetails.Tin = Trim(GetText(18, 19, 9))
            MyclaimDetails.Prifix_Tin_Suffix = Trim(GetText(23, 19, 1)) + Trim(GetText(23, 20, 9)) + Trim(GetText(23, 30, 5))
            MyclaimDetails.Prefix_Tin = Trim(GetText(23, 19, 10))
            CurrentClaim(0).ICN = icn
            CurrentClaim(0).Tin = GetText(23, 19, 1) + "-" + GetText(23, 20, 9) + "-" + GetText(23, 30, 5)
            CurrentClaimCount = -1
            SendClear()
            SendClear()
            pullEDSScreen(icn, "1")
            apiChk()
            WriteLog("Exit GetClaimDetails()..........")
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try
    End Sub
    Public Sub GetEDS1(ByRef OREDS1() As EDS1Value, ByVal icn As String)
        Try
            Dim X As Integer
            Dim Y As Integer

            Dim pagecount As Integer = 0
            Dim CheckLine As String
            CheckLine = String.Empty
            Dim count As Integer
            Dim RevCodeCount As Integer
            count = 0
            RevCodeCount = 0
            X = 0
            Dim LineCount As Integer = -1
            count = 0
            'ReDim OREDS1(0)
            For Y = 8 To 21 Step 1

                CheckLine = Trim(GetText(Y, 2, 4))
                If InStr(Trim(GetText(Y, 66, 6)), "NYS") Then
                    strResearchComments = "EDS1 NYS Line present"
                    Exit Sub
                End If
                If CheckLine = String.Empty Then Exit For
                If Trim(GetText(Y, 66, 6)) <> "" Then 'New translation

                    X = 0
                    LineCount += 1
                    count += 1
                    If LineCount = 0 Then
                        ReDim OREDS1(LineCount)
                    Else
                        ReDim Preserve OREDS1(LineCount)
                    End If
                    OREDS1(LineCount).FDate = Trim(GetText(3, 21, 6))
                    OREDS1(LineCount).LDate = Trim(GetText(3, 33, 6))
                    OREDS1(LineCount).ServiceLine.LineTranslation = Trim(GetText(Y, 66, 6))
                    OREDS1(LineCount).ServiceWiseTotalCharge = Trim(GetText(Y, 72, 10))
                    OREDS1(LineCount).ServiceLineCount = count

                End If

                If X = 0 Then
                    RevCodeCount += 1
                    ReDim OREDS1(LineCount).ServiceLine.RevCode(X)
                    ReDim OREDS1(LineCount).ServiceLine.CPTCode(X)
                    ReDim OREDS1(LineCount).ServiceLine.M1(X)
                    ReDim OREDS1(LineCount).ServiceLine.M2(X)
                    ReDim OREDS1(LineCount).ServiceLine.DateofService(X)
                    ReDim OREDS1(LineCount).ServiceLine.BillCharge(X)
                    ReDim OREDS1(LineCount).ServiceLine.RevCodeCount(X)
                    ReDim OREDS1(LineCount).ServiceLine.Unit(X)
                    ReDim OREDS1(LineCount).ServiceLine.PS(X)
                    ReDim OREDS1(LineCount).ServiceLine.ServiceCode(X)
                Else
                    RevCodeCount += 1
                    ReDim Preserve OREDS1(LineCount).ServiceLine.RevCode(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.CPTCode(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.M1(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.M2(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.DateofService(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.BillCharge(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.RevCodeCount(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.Unit(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.PS(X)
                    ReDim Preserve OREDS1(LineCount).ServiceLine.ServiceCode(X)
                End If
                OREDS1(LineCount).ServiceLine.RevCode(X) = Trim(GetText(Y, 2, 4))
                OREDS1(LineCount).ServiceLine.CPTCode(X) = Trim(GetText(Y, 36, 5))
                OREDS1(LineCount).ServiceLine.M1(X) = Trim(GetText(Y, 41, 2))
                OREDS1(LineCount).ServiceLine.M2(X) = Trim(GetText(Y, 43, 2))
                OREDS1(LineCount).ServiceLine.DateofService(X) = Trim(GetText(Y, 46, 4))
                OREDS1(LineCount).ServiceLine.BillCharge(X) = Trim(GetText(Y, 51, 8))
                OREDS1(LineCount).ServiceLine.Unit(X) = Trim(GetText(Y, 31, 4))
                OREDS1(LineCount).ServiceLine.RevCodeCount(X) = X
                OREDS1(LineCount).ServiceLine.PS(X) = Trim(GetText(Y, 63, 2))
                OREDS1(LineCount).ServiceLine.ServiceCode(X) = Trim(GetText(Y, 66, 5))
                X += 1
                If Y = 21 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "PF4 TO REVIEW EDC/ICES DECISION") > 0 Then
                        SendF8()
                    End If
                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO ADDITIONAL LINES TO DISPLAY") > 0 Then Exit For

                    If InStr(autECLPSObj.GetText(24, 1, 80), "PF8 FOR ADDITIONAL LINES") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "I049 - MORE DATA ON EDSS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "PF4 TO REVIEW EDC/ICES DECISION") > 0 Then
                        SendF8()
                        'WriteLog(vbCrLf)
                        'WriteLog(GetTextRect())
                    Else
                        Exit For
                    End If
                    Y = 7

                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try

    End Sub
    Public Function DetarmineClaimisAllReadyProcess(ByVal icn As String) As String
        SendClear()
        SendClear()
        Dim i As Integer
        Dim NBR As String
        Dim FinalNBR As String
        NBR = String.Empty
        FinalNBR = String.Empty
        pullEDSScreen(icn, "9")
        For i = 10 To 22 Step 2
            If GetColorScheme(i, 2) = 0 Then
                NBR = Trim(GetText(i, 33, 3))
                If NBR <> String.Empty Then
                    FinalNBR += NBR + ","
                End If
            ElseIf GetColorScheme(i, 2) = 0 Then
                If InStr(Trim(GetText(i, 27, 4)), "APAY") > 0 Or InStr(Trim(GetText(i, 27, 4)), "CPAY") > 0 Or InStr(Trim(GetText(i, 27, 4)), "MPAY") > 0 Then
                    NBR = Trim(GetText(i, 33, 3))
                    If NBR <> String.Empty Then
                        FinalNBR += NBR + ","
                    End If
                End If
            End If
            If i = 22 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY ON LAST HISTORY PAGE") > 0 Then Exit For
                SendF8()
                i = 8
            End If
        Next
        If FinalNBR <> String.Empty Then Return FinalNBR
        Return ""
    End Function
    Public Function GetEditFromPaymentScreen() As String
        'WriteLog("------------ inside " + GetCurrentMethod().Name + " ------------")
        Dim EditWarning As String = ""
        Dim Count As Integer = 0
        Try
            Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                If Count > 50 Then Exit Do
                EditWarning += GetText(24, 1, 80, True)
                SendEnter()
            Loop
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return EditWarning
    End Function
    Public Function CheckWorkQueueForCorrectedClaim(ByVal ICN As String) As Boolean
        Dim WQ As String = ""
        SendClear()
        SendClear()
        pullEDSScreen(ICN, "9")
        For I As Integer = 10 To 22 Step 2
            If GetText(I, 1, 80, True) <> "" Then
                WQ += GetText(I, 19, 2) + ","
            End If
            If I = 22 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY ON LAST HISTORY PAGE") = 0 Then
                    SendF8()
                    I = 8
                Else
                    Exit For
                End If
            End If
        Next
        If InStr(WQ, "CR") = 0 Then Return False
        Return True
    End Function

    Public Sub GETEDS9(ByRef OREDS9() As EDS9Value, icn As String)
        Try
            Dim I As Integer
            Dim J As Integer
            Dim NBR As String
            NBR = String.Empty
            Dim checkline As String
            SendClear()
            SendClear()
            pullEDSScreen(icn, "9")
            ReDim OREDS9(0)
            I = 0
            Dim count As Integer = 1
            Dim Linecount As Integer = 0
            OREDS9(0).JulianDate = Trim(GetText(2, 23, 5))
            For J = 10 To 22 Step 2

                checkline = Trim(GetText(J, 5, 6))

                If checkline = "" Then Exit For
                If I = 0 Then
                    ReDim OREDS9(Linecount).WQ(I)
                    ReDim OREDS9(Linecount).LineColor(I)
                    ReDim OREDS9(Linecount).Status(I)
                    ReDim OREDS9(Linecount).NBR(I)
                    ReDim OREDS9(Linecount).PC(I)
                    ReDim OREDS9(Linecount).TopsID(I)
                    ReDim OREDS9(Linecount).LineCount(I)
                    ReDim OREDS9(Linecount).Comment(I)

                Else
                    ReDim Preserve OREDS9(Linecount).WQ(I)
                    ReDim Preserve OREDS9(Linecount).LineColor(I)
                    ReDim Preserve OREDS9(Linecount).Status(I)
                    ReDim Preserve OREDS9(Linecount).NBR(I)
                    ReDim Preserve OREDS9(Linecount).PC(I)
                    ReDim Preserve OREDS9(Linecount).TopsID(I)
                    ReDim Preserve OREDS9(Linecount).LineCount(I)
                    ReDim Preserve OREDS9(Linecount).Comment(I)
                End If
                OREDS9(Linecount).WQ(I) = Trim(GetText(J, 19, 2))
                OREDS9(Linecount).LineColor(I) = GetColorScheme(J, 2)
                OREDS9(Linecount).Status(I) = Trim(GetText(J, 27, 4))
                OREDS9(Linecount).NBR(I) = Trim(GetText(J, 33, 3))
                OREDS9(Linecount).PC(I) = Trim(GetText(J, 23, 2))
                OREDS9(Linecount).TopsID(I) = Trim(GetText(J, 49, 8))
                OREDS9(Linecount).LineCount(I) = count
                OREDS9(Linecount).Comment(I) = Trim(GetText(J + 1, 6, 70))
                I = I + 1
                count += 1
                'Linecount += 1
                If J = 22 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY ON LAST HISTORY PAGE") > 0 Then Exit For
                    SendF8()
                    J = 8
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    'Public Function FindHistoryICNNew(ByRef MyList() As PossibleICN, ByRef FlagError As Boolean, ByRef TotalPaidInHx As String, ByVal isHicf As Boolean, ByVal isUB As Boolean, ByVal HxICN As String) As String
    '    Dim value As String = "Not Used"
    '    Dim isHicf1 As Boolean = False
    '    Dim isUB1 As Boolean = False
    '    Dim isIH1 As Boolean = False
    '    Dim hxFLn As String    'FlagError As Boolean = False,
    '    Dim PossibleICN As String, CurrentClaimProcCode As String, X As Integer, Y As Integer
    '    Dim Count3 As Integer = 0
    '    Dim J As Integer, ProcCodeMatchCount As Integer, MyClaim(100) As Claim, IndexCount As Integer = 0
    '    Dim ListOfPossibleICN() As PossibleICN
    '    TotalPaidInHx = "0"
    '    CurrentClaimProcCode = ""
    '    For X = 0 To CurrentClaimCount - 1
    '        For Y = LBound(CurrentClaim(X).ServiceLine) To UBound(CurrentClaim(X).ServiceLine) - 1
    '            If InStr(CurrentClaimProcCode, Mid(CurrentClaim(X).ServiceLine(Y), 4, 5)) = 0 And InStr(InvalidProcCode, Trim(Mid(CurrentClaim(X).ServiceLine(Y), 4, 5))) = 0 Then
    '                CurrentClaimProcCode = CurrentClaimProcCode & Mid(CurrentClaim(X).ServiceLine(Y), 4, 5) & ","
    '            End If
    '        Next
    '    Next
    '    While Microsoft.VisualBasic.Right(CurrentClaimProcCode, 1) = ","
    '        CurrentClaimProcCode = Microsoft.VisualBasic.Left(CurrentClaimProcCode, Len(CurrentClaimProcCode) - 1)
    '    End While
    '    ReDim CurrentAHI(100)
    '    PossibleICN = ""
    '    For X = 0 To CountAHI
    '        'If X = 33 Then MsgBox("Stop")
    '        If CurrentAHI(X).ICN <> CurrentClaim(0).ICN Then
    '            If (CDbl(Trim(CurrentAHI(X).Paid)) <> 0 Or CDbl(Trim(CurrentAHI(X).Ded)) <> 0) Or CDbl(CurrentAHI(X).NC) = 0 Then
    '                If InStr(PossibleICN, CurrentAHI(X).ICN) = 0 And CurrentAHI(X).ICN <> CurrentClaim(0).ICN Then
    '                    PossibleICN = PossibleICN & CurrentAHI(X).ICN & ","
    '                End If
    '            ElseIf InStr(CurrentAHI(X).OV, "C") > 0 Then
    '                If InStr(PossibleICN, CurrentAHI(X).ICN) = 0 And CurrentAHI(X).ICN <> CurrentClaim(0).ICN Then
    '                    PossibleICN = PossibleICN & CurrentAHI(X).ICN & ","
    '                End If
    '            Else

    '                For Each RC In Split(CurrentAHI(X).RemarkCode, ",")
    '                    If Trim(RC) <> "" Then
    '                        'If InStr(ICESList, RC) > 0 Or InStr("05,TQ,G1,NI,EW,UP,LY,11,GR,CA", RC) > 0 Then
    '                        If InStr(PossibleICN, CurrentAHI(X).ICN) = 0 And CurrentAHI(X).ICN <> CurrentClaim(0).ICN Then
    '                            PossibleICN = PossibleICN & CurrentAHI(X).ICN & ","
    '                        End If
    '                        'End If
    '                    End If
    '                Next
    '            End If
    '        End If
    '    Next
    '    While Microsoft.VisualBasic.Right(PossibleICN, 1) = ","
    '        PossibleICN = Microsoft.VisualBasic.Left(PossibleICN, Len(PossibleICN) - 1)
    '    End While
    '    If PossibleICN = "" Then Return ""
    '    If PossibleICN.Length - PossibleICN.Replace(",", "").Length > 2 Then
    '        strResearchComments = "More than 2 Possible ICN Found"
    '        ExitAllAndReturnWithException()
    '    End If

    '    IndexCount = 0
    '    hxFLn = ""
    '    If HxICN <> "" Then
    '        PullMHI(HxICN)

    '        If FlagError = True Then
    '            Exit Function
    '        End If
    '        J = ReadMHI(MyClaim)
    '        hxFLn = MyClaim(0).FLN
    '    End If
    '    ReDim ListOfPossibleICN(1000)
    '    For Each AHIICN As String In PossibleICN.Split(",")
    '        If AHIICN.Trim <> "" Then
    '            Dim AhiIndx As Integer = GetIndexInAhiByICN(CurrentAHI, AHIICN)
    '            SendClear()
    '            SendClear()
    '            ListOfPossibleICN(IndexCount).IsDXMatch = False
    '            autECLPSObj.SendKeys("RET,I" & AHIICN & ",M", 2, 2)
    '            SendEnter()

    '            J = ReadMHI(MyClaim)
    '            isUB1 = False
    '            isHicf1 = False
    '            isIH1 = False
    '            CheckForClaimType(MyClaim, isUB1, isIH1, isHicf1, J)
    '            ProcCodeMatchCount = 0
    '        End If
    '    Next




    '    Return value
    'End Function
    Public Sub CheckForClaimType(ByRef myClaim() As Claim, ByRef isUB As Boolean, ByRef isIH As Boolean, ByRef isHICF As Boolean, ByVal Count As Integer)
        Dim J As Integer, MHIControlLine As String, K As Integer, FlagTemp As Boolean
        Dim NotificationClaimNumber As Integer
        NotificationClaimNumber = Count
        FlagTemp = False
        isUB = False
        isIH = False
        isHICF = False
        For J = 0 To NotificationClaimNumber - 1
            If CDbl(myClaim(J).Draft) <> 0 Then
                FlagTemp = True
                MHIControlLine = "RHI," & myClaim(0).Policy & "," & myClaim(0).SSN & "," & myClaim(0).EmployeeName & "," & myClaim(0).Relationship & "," & myClaim(J).Draft & ",I" & myClaim(0).ICN
                SendClear()
                SendClear()
                SendKeyes(MHIControlLine, 2, 2)
                SendEnter()
                'WriteLog("RHI for ICN " & myClaim(0).ICN & " is ")
                ' WriteLog(GetTextRect)
                If autECLPSObj.GetText(8, 3, 3) = "REV" Then
                    isUB = True
                    For K = 8 To 23
                        If IsNumeric(autECLPSObj.GetText(K, 3, 4)) Then
                            If CInt(autECLPSObj.GetText(K, 3, 4)) < 220 Then
                                isIH = True
                                Exit For
                            End If
                        End If
                        If K = 23 Then
                            If InStr(autECLPSObj.GetText(24, 1, 80), "W759 MORE RECORDS") Then
                                autECLPSObj.SendKeys("[PF8]")
                                apiChk()
                                K = 7
                            Else
                                Exit For
                            End If
                        End If
                    Next
                    GoTo CheckForIH
                ElseIf GetText(8, 2, 4) = "PROC" Then
                    isHICF = True : GoTo CheckForIH
                End If
            End If
        Next
        If FlagTemp = False Then
            SendClear()
            SendClear()

            pullEDSScreen(myClaim(0).ICN, "1")
            If autECLPSObj.GetText(7, 2, 4) = "REVE" Then
                isUB = True
            Else
                isHICF = True
            End If
        End If
        If isUB = False AndAlso isHICF = False Then
            For I As Integer = 0 To NotificationClaimNumber - 1
                For KK As Integer = LBound(myClaim(I).ServiceLine) To UBound(myClaim(I).ServiceLine) - 1
                    If Left(myClaim(I).ProcCode(KK), 1) = "C" AndAlso Right(myClaim(I).ProcCode(KK).Trim, 3) = "REV" Then
                        isUB = True
                        Exit For
                    ElseIf InStr("OPS,SP,PR,IC,WD,MISC", myClaim(I).ProcCode(KK).Trim) > 0 Then
                        isUB = True
                        Exit For
                    End If
                Next
            Next
            If isUB = False Then isHICF = True
        End If

CheckForIH:
        If isHICF = True Then
            If Mid(myClaim(0).ServiceLine(0), 1, 2) = "IH" Or Mid(myClaim(0).ServiceLine(0), 1, 2) = "CF" Or Mid(myClaim(0).ServiceLine(0), 1, 2) = "CL" Then
                isIH = True
            Else
                isIH = False
            End If
        End If
    End Sub
    Public Function GetIndexInAhiByICN(ByRef MyAHI() As AHI, ByVal ICN As String) As Integer
        For I As Integer = LBound(MyAHI) To UBound(MyAHI)
            If MyAHI(I).ICN = ICN Then Return I
        Next
        Return -999
    End Function
    Function ReadAHI(ByRef c() As AHI) As Integer
        Dim X As Integer, Y As Integer, IcnList As String = ""
        X = 0
        For Y = 4 To 20 Step 2
            If Replace(Replace(autECLPSObj.GetText(Y, 1, 80), "-", ""), " ", "") = "" Then Exit For
            If InStr(IcnList, autECLPSObj.GetText(Y + 1, 54, 10)) = 0 AndAlso autECLPSObj.GetText(Y + 1, 54, 10) <> CurrICN AndAlso GetText(Y + 1, 52, 1) = "I" Then
                IcnList += autECLPSObj.GetText(Y + 1, 54, 10) + ","
                c(X).FDOS = autECLPSObj.GetText(Y, 4, 6)
                c(X).LDOS = autECLPSObj.GetText(Y, 11, 6)
                c(X).Paid = autECLPSObj.GetText(Y, 33, 5)
                c(X).Ded = autECLPSObj.GetText(Y, 43, 5)
                c(X).NC = Trim(autECLPSObj.gettext(Y, 54, 9))
                c(X).RemarkCode = Replace(Replace(Replace(autECLPSObj.GetText(Y, 64, 8), " ", ","), "--", ""), ",,", ",")
                c(X).OV = Replace(Replace(Replace(autECLPSObj.gettext(Y, 73, 8), " ", ","), "--", ""), ",,", ",")
                c(X).ICN = autECLPSObj.GetText(Y + 1, 54, 10)
                c(X).ProviderName = GetText(Y + 1, 2, 28, True)
                c(X).ProcDT = GetText(Y + 1, 73, 6)
                X = X + 1
            End If

            If Y = 20 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 Then Exit For
                SendF8()
                Y = 2
            End If
        Next
        SendKeyes("AHI", 2, 2)
        SendEnter()
        ReDimension(c, X - 1)
        X = 0
        IcnList = ""
        For Y = 4 To 20 Step 2
            If Replace(Replace(autECLPSObj.GetText(Y, 1, 80), "-", ""), " ", "") = "" Then Exit For
            If InStr(IcnList, autECLPSObj.GetText(Y + 1, 54, 10)) = 0 AndAlso autECLPSObj.GetText(Y + 1, 54, 10) <> CurrICN AndAlso GetText(Y + 1, 52, 1) = "I" Then
                IcnList += autECLPSObj.GetText(Y + 1, 54, 10) + ","
                c(X).FDOS = autECLPSObj.GetText(Y, 4, 6)
                c(X).LDOS = autECLPSObj.GetText(Y, 11, 6)
                c(X).Paid = autECLPSObj.GetText(Y, 33, 5)
                c(X).Ded = autECLPSObj.GetText(Y, 43, 5)
                c(X).NC = Trim(autECLPSObj.gettext(Y, 54, 9))
                c(X).RemarkCode = Replace(Replace(Replace(autECLPSObj.GetText(Y, 64, 8), " ", ","), "--", ""), ",,", ",")
                c(X).OV = Replace(Replace(Replace(autECLPSObj.gettext(Y, 73, 8), " ", ","), "--", ""), ",,", ",")
                c(X).ICN = autECLPSObj.GetText(Y + 1, 54, 10)
                c(X).ProviderName = GetText(Y + 1, 2, 28, True)
                X = X + 1
            End If

            If Y = 20 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 Then Exit For
                SendF8()
                Y = 2
            End If
        Next
        ReadAHI = X - 1
    End Function

    Public Sub RemovePZ1(ByVal CurrentSuffix As Integer, Optional ByVal Flag As Boolean = False)
        Try
            Dim TempStr As String = ""
            If Flag = False Then CheckAndRemove30Line(CurrentSuffix)
            RemoveBlankLine()

            If InStr(GetText(2, 1, 80), "PZ1") = 0 Then
                Do
                    TempStr = TempStr & GetText(24, 1, 80)
                    SendEnter()
                    apiChk()
                    apiChk()
                    'MsgBox TempStr
                Loop Until (InStr(TempStr, GetText(24, 1, 80)) > 0 And TempStr <> "") Or InStr(GetText(2, 1, 80), "PZ1") > 0

            End If
            If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                SendOverrideInScreen("01")
            End If
            RemoveBlankAndPISLLine()
            If strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE" Then
                Exit Sub
            End If

            If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                SendOverrideInScreen("01")
                SendEnter()
            End If
            RemoveDenialWithoutRemarkCode()
            CaptureScreenLog()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub

    Public Sub RemovePZ1()
        Try
            Dim TempStr As String = ""
            RemoveBlankLine()
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                Exit Sub
            End If

            If InStr(GetText(2, 1, 80), "PZ1") = 0 Then
                Do
                    TempStr = TempStr & GetText(24, 1, 80)
                    SendEnter()
                    apiChk()
                    apiChk()
                    'MsgBox TempStr
                Loop Until (InStr(TempStr, GetText(24, 1, 80)) > 0 And TempStr <> "") Or InStr(GetText(2, 1, 80), "PZ1") > 0

            End If
            If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                SendOverrideInScreen("01")
            End If
            RemoveBlankAndPISLLine()
            If strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE" Then
                Exit Sub
            End If

            If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                SendOverrideInScreen("01")
                SendEnter()
            End If
            RemoveDenialWithoutRemarkCode()
            CaptureScreenLog()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try
    End Sub

    Public Sub CheckAndRemove30Line(ByVal CurrentSuffix As Integer)
        Try
            Dim PageCount As Integer = 1
            Dim Linecount As Integer = 0
            Dim linecount2 As Integer = 0
            Dim LineCountPaymentScreen As Integer = 0

            For i = 10 To 22 Step 2
                If InStr(InvalidProcCode, GetText(i, 4, 5)) = 0 Then linecount2 += 1
                If GetText(i, 1, 80, True).Replace("-", "").Trim.Replace("N", "") <> "" AndAlso (GetText(i, 29, 2) = "30") Then 'OrElse isLinePresentInEds1(CurrentICNEDS1, linecount2) = False Then  'OrElse isLinePresentInEds1(CurrentClaimModifier, CurrentSuffix, linecount2) = False) Then
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    Linecount = 0
                    For j = 6 To 18 Step 3 'Infinite Loop exception is handled
                        If Linecount >= linecount2 Then Exit For
                        If GetText(j, 1, 80, True) = "" OrElse PageCount > 2 Then Exit For
                        SendKeyes("--", j, 69)
                        '//Sanjay : Trim spaces for PISL line. Please create a custom gettext function and trim all the space once.
                        If InStr(InvalidProcCode, GetText(j, 11, 5, True)) = 0 Then
                            Linecount += 1
                        End If
                        '// Problem is here

                        'isLinePresentInEds1(CurrentClaimModifier, CurrentSuffix, Linecount) = False
                        If GetText(j, 38, 2) = "30" Then 'OrElse isLinePresentInEds1(CurrentICNEDS1, linecount2) = False Then
                            Dim a = IIf(InStr(CStr(CDbl(GetText(3, 25, 7) + "." + GetText(3, 33, 2)) - CDbl(GetText(j, 49, 6).Replace("-", "0") + "." + GetText(j, 56, 2).Replace("-", "0"))), ".") = 0, CStr(CDbl(GetText(3, 25, 7) + "." + GetText(3, 33, 2)) - CDbl(GetText(j, 49, 6).Replace("-", "0") + "." + GetText(j, 56, 2).Replace("-", "0"))) + ".00", CStr(CDbl(GetText(3, 25, 7) + "." + GetText(3, 33, 2)) - CDbl(GetText(j, 49, 6).Replace("-", "0") + "." + GetText(j, 56, 2).Replace("-", "0"))))
                            SendKeyes("------", j, 11)
                            SendKeyes("BLANK", j, 11)
                            SendEnter()
                            If GetText(j, 38, 2) <> "30" Then
                                SendKeyes(ConvertToFormat(Split(a, ".")(0), 7, True, "0"), 3, 25)
                                SendKeyes(ConvertToFormat(Split(a, ".")(1), 2, , "0"), 3, 33)
                            End If

                            'Exit For
                        Else
                            SendKeyes("01", j, 38)
                            SendEnter()
                        End If

                        If j = 18 Then
                            If InStr(GetText(24, 1, 80, True), "NO MORE") > 0 OrElse PageCount > 2 OrElse InStr(GetText(24, 1, 80, True), "ALREADY") Then Exit For
                            SendEnter()
                            SendF8()
                            PageCount += 1
                            j = 3
                        End If
                    Next

                    SendEnter()

                    SendKeyes("A", 1, 75)
                    SendEnter()

                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If

                    If GetText(2, 2, 2, True) = "AR" Then
                        SendKeyes("S", 3, 5)
                        SendEnter()
                    End If
                    If GetText(2, 2, 2) <> "MP" Then
                        strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                        Exit Sub
                        ExitAllAndReturnWithException()
                    End If
                    If GetText(2, 2, 2, True) = "AR" Then
                        SendKeyes("S", 3, 5)
                        SendEnter()
                    End If
                    If GetText(2, 2, 2) <> "MP" Then
                        strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                        ExitAllAndReturnWithException()
                    End If
                    GetCorrectBillChargeInPaymentScreen()
                    PutCopayNBR()
                    Exit For
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub RemoveBlankLine()
        Try
            Dim FlagLineFound As Boolean = False
            Dim FlagBlankLineFound As Boolean = False
            Dim RepeatCount As Integer = 1

            Do
                FlagLineFound = False
                FlagBlankLineFound = False
                For I As Integer = 22 To 10 Step -2
                    If GetText(I, 1, 80).Replace("-", "").Trim.Replace("N", "") <> "" And FlagLineFound = False Then
                        FlagLineFound = True
                    Else
                        If FlagLineFound Then
                            If GetText(I, 1, 80).Replace("-", "").Trim.Replace("N", "") = "" Then
                                FlagBlankLineFound = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                If FlagBlankLineFound Then
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    Dim PageCount As Integer = 1
                    For I As Integer = 6 To 18 Step 3 'Infinite Loop exception is handled
                        If GetText(I, 1, 80, True) = "" Then Exit For
                        SendKeyes("--", I, 69)
                        SendKeyes("--", I, 72)
                        SendKeyes("--", I, 75)
                        SendKeyes("--", I, 78)
                        If GetText(I, 11, 6) = "------" Then
                            SendKeyes("BLANK-", I, 11)
                            SendEnter()
                            If InStr("BLANK-,------", GetText(I, 11, 6)) = 0 Then
                                I -= 3
                            End If
                        ElseIf GetText(I, 11, 6) = "BLANK-" Then
                            SendKeyes("------", I, 11)
                            SendEnter()
                            If GetText(I, 11, 6) = "BLANK-" Then
                                SendKeyes("------", I, 11)
                                SendEnter()
                            End If
                            If InStr("BLANK-,------", GetText(I, 11, 6)) = 0 Then
                                I -= 3
                            End If
                        Else
                            SendKeyes("01", I, 38)
                            SendKeyes("--", I, 45)
                            If GetColorScheme(I + 1, 6) <> 0 Then
                                SendKeyes("opn", I, 2)
                                SendEnter()
                                'SendKeyes("01", 7, 45)
                                SendKeyes("------", 7, 58)
                                SendKeyes("--", 7, 65)
                                SendKeyes("--", 7, 68)
                                SendEnter()
                                SendF3()
                                If CDbl(GetText(I, 59, 6).Replace("-", "0") + "." + GetText(I, 66, 2).Replace("-", "0")) > 0 Then
                                    If GetColorScheme(I, 59) <> 0 Then
                                        'flagerror = True
                                        Exit Sub
                                    Else
                                        SendKeyes("------", I, 59)
                                        SendKeyes("--", I, 66)
                                        SendEnter()
                                    End If
                                End If
                            Else
                                'End If
                                SendKeyes("000000", I + 1, 6)
                                SendKeyes("00", I + 1, 13)
                                SendKeyes("000", I + 1, 20)
                                SendKeyes("---", I + 1, 50)
                                SendKeyes("---", I + 1, 54)
                                SendKeyes("01", I, 38)
                                If IsNumeric(GetText(I, 49, 9).Replace(" ", ".")) AndAlso CDbl(GetText(I, 49, 9).Replace(" ", ".")) = 0.01 Then
                                    'SendKeyes("000000", KK + 1, 6)
                                    'SendKeyes("01", KK + 1, 13)
                                    'SendKeyes("LN-", KK + 1, 20)
                                    'PlaceLNorLYinHalfScreen(I)
                                End If
                                SendEnter()
                                If GetColorScheme(I + 1, 6) = 3 Then
                                    SendKeyes("opn", I, 2)
                                    SendEnter()
                                    'SendKeyes("01", 7, 45)
                                    SendKeyes("------", 7, 58)
                                    SendKeyes("--", 7, 65)
                                    SendKeyes("--", 7, 68)
                                    SendEnter()
                                    SendF3()
                                    If CDbl(GetText(I, 59, 6).Replace("-", "0") + "." + GetText(I, 66, 2).Replace("-", "0")) > 0 Then
                                        If GetColorScheme(I, 59) <> 0 Then
                                            'flagerror = True
                                            Exit Sub
                                        Else
                                            SendKeyes("------", I, 59)
                                            SendKeyes("--", I, 66)
                                            SendEnter()
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If I = 18 Then
                            SendEnter()
                            SendF8()
                            If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 OrElse PageCount > 2 Then
                                Exit For
                            Else
                                I = 3
                                PageCount += 1
                            End If
                        End If
                    Next

                    'SendEnter()     'CHECK WHY THE ENTER
                    If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                        SendKeyes("N", 4, 79)
                    End If
                    SendEnter()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    'apiChk()
                    'apiChk()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If

                    If GetText(2, 2, 2, True) = "AR" Then
                        SendKeyes("S", 3, 5)
                        SendEnter()
                    End If
                    If GetText(2, 2, 2) <> "MP" Then
                        strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                        Exit Sub
                        ExitAllAndReturnWithException()
                    End If
                    PutCopayNBR()
                End If

                RepeatCount += 1
            Loop Until Not FlagBlankLineFound OrElse RepeatCount > 2
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub PutCopayNBR()
        Try
            For ii As Integer = 1 To SuffixCountOfCurrentICN
                SendClear()
                SendClear()
                pullEDSScreen(CurrICN, 2)
                SendKeyes("A" & ii, 1, 76)
                SendEnter()
                apiChk()

                If GetText(2, 2, 3) <> "MPC" Then Exit Sub
                For I As Integer = 10 To 22 Step 2
                    If GetText(I, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                    If GetText(I, 4, 6, True) = "NCOPAY" And CInt(GetText(I, 25, 3, True).Replace("-", "0")) = 0 Then
                        SendKeyes(ConvertToFormat(GetCopayNBR.ToString, 3, True, "0"), I, 25)
                    End If
                Next
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try
    End Sub
    Public Function GetCopayNBR() As String
        Dim Count As Integer = 0
        Dim Draft As String = ""
        If ListOfPossibleICNActual(0).ICN IsNot Nothing Then 'AndAlso IsNumeric(ListOfPossibleICNActual(0).ICN) Then
            For J As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                If ListOfPossibleICNActual(J).ICN IsNot Nothing Then 'AndAlso IsNumeric(ListOfPossibleICNActual(J).ICN) Then
                    Draft = ""
                    For K As Integer = LBound(ListOfPossibleICNActual(J).TempClaim) To GetClaimCountFromStruct(ListOfPossibleICNActual(J).TempClaim) - 1
                        If CDbl(ListOfPossibleICNActual(J).TempClaim(K).Draft) > 0 AndAlso InStr(Draft, ListOfPossibleICNActual(J).TempClaim(K).Draft) = 0 Then
                            Draft += ListOfPossibleICNActual(J).TempClaim(K).Draft + ","
                            If ListOfPossibleICNActual(J).TempClaim(K).Voided = False Then
                                For L As Integer = LBound(ListOfPossibleICNActual(J).TempClaim(K).ServiceLine) To UBound(ListOfPossibleICNActual(J).TempClaim(K).ServiceLine) - 1
                                    If ListOfPossibleICNActual(J).TempClaim(K).ProcCode(L) = "NCOPAY" AndAlso CDbl(ListOfPossibleICNActual(J).TempClaim(K).NotCovered(L)) > 0 Then
                                        Count += CInt(ListOfPossibleICNActual(J).TempClaim(K).NBR(L).Trim)
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            Next
        Else
            Return 1
        End If
        Return IIf(Count = 0, 1, Count)
    End Function
    Public Function GetClaimCountFromStruct(ByRef MyClaim() As Claim) As Integer
        Dim I As Integer
        For I = LBound(MyClaim) To UBound(MyClaim)
            If MyClaim(I).ICN Is Nothing OrElse MyClaim(I).ICN.Trim = "" Then Return I 'Or IsNumeric(MyClaim(I).ICN) = False Then Return I
        Next
        Return I
    End Function
    Public Sub GetCorrectBillChargeInPaymentScreen()
        Try
            Dim TotalCharge As Decimal
            If GetColorScheme(6, 4) = 3 Then
                For I As Integer = 10 To 22 Step 2
                    'If GetText(I, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                    If InStr(InvalidProcCode, GetText(I, 4, 5)) = 0 Then
                        TotalCharge += CDec(GetText(I, 41, 5).Replace("-", "0") + "." + GetText(I, 47, 2).Replace("-", "0"))
                    End If
                Next
                Dim a = IIf(InStr(CStr(TotalCharge), ".") = 0, CStr(TotalCharge) + ".00", CStr(TotalCharge))
                SendKeyes(ConvertToFormat(Split(a, ".")(0), 6, True, "0"), 6, 4)
                SendKeyes(ConvertToFormat(Split(a, ".")(1), 2, , "0"), 6, 11)
                SendEnter()
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub RemoveDenialWithoutRemarkCode()
        'Dim Flag As Boolean = False
        'For I As Integer = 10 To 22 Step 2
        '    Dim BC = GetText(I, 41, 5).Replace(" ", "0").Replace("-", "0") + "." + GetText(I, 47, 2).Replace(" ", "0").Replace("-", "0")
        '    Dim NC = GetText(I, 50, 5).Replace(" ", "0").Replace("-", "0") + "." + GetText(I, 56, 2).Replace(" ", "0").Replace("-", "0")
        '    If IsNumeric(BC) AndAlso IsNumeric(NC) AndAlso CDbl(NC) > 0.01 AndAlso GetText(I, 36, 2).Replace("-", "") = "" Then
        '        Flag = True
        '        Exit For
        '    End If
        'Next
        'If Flag = False Then Exit Sub
        Try
            CorrectLNBillCharge()

            For i = 10 To 22 Step 2
                If GetText(i, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                If GetText(i, 36, 2) = "--" Then
                    If CDbl(GetText(i, 41, 8, True).Replace(" ", ".").Replace("-", "0")) = CDbl(GetText(i, 50, 8, True).Replace(" ", ".").Replace("-", "0")) Then
                        SendKeyes("00000", i, 50)
                        SendKeyes("00", i, 56)
                    End If
                End If
            Next
            SendEnter()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub CorrectLNBillCharge()
        Try
            Dim Flag As Boolean = False
            For I As Integer = 10 To 22 Step 2
                If InStr(PISLLINE, GetText(I, 4, 6, True).Replace("-", "").Trim) = 0 And GetText(I, 4, 6, True).Replace("-", "").Trim <> "" AndAlso InStr(InvalidProcCode, GetText(I, 4, 6, True).Replace("-", "").Trim) = 0 Then
                    Dim BC = GetText(I, 41, 5).Replace(" ", "0").Replace("-", "0") + "." + GetText(I, 47, 2).Replace(" ", "0").Replace("-", "0")
                    If IsNumeric(BC) = False Then
                        Flag = True
                        Exit For
                    ElseIf CDbl(BC) = 0 Then
                        Flag = True
                        Exit For
                    ElseIf CDbl(BC) = 0.01 Then
                        If GetText(I, 36, 2) <> "LY" AndAlso InStr(GetText(I, 4, 6, True), "F") > 0 Then
                            SendKeyes("LN-", I, 36)
                        Else
                            SendKeyes("LY-", I, 36)
                        End If

                    End If
                End If
            Next
            If Flag = False Then Exit Sub
            SendKeyes("EDS", 2, 2)
            SendEnter()
            Dim PageCount As Integer = 1
            For KK As Integer = 6 To 18 Step 3 'Infinite Loop exception is handled
                If GetText(KK, 11, 6, True) = "" Then
                    Exit For
                End If
                If InStr(PISLLINE, GetText(KK, 11, 6, True)) = 0 And GetText(KK, 11, 6, True) <> "" AndAlso InStr(InvalidProcCode, GetText(KK, 11, 6, True)) = 0 Then
                    If IsNumeric(GetText(KK, 49, 9).Replace(" ", ".")) = False Then
                        'SendKeyes("000000", KK, 49)
                        'SendKeyes("01", KK, 56)
                        'SendKeyes("000000", KK + 1, 6)
                        'SendKeyes("01", KK + 1, 13)
                        'SendKeyes("LN", KK, 45)
                        'SendKeyes("LN", KK + 1, 20)
                        'PlaceLNorLYinHalfScreen(KK)
                        'SendEnter()
                    End If
                End If
                If KK = 18 Then
                    SendEnter()
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 OrElse PageCount > 2 Then
                        Exit For
                    Else
                        KK = 3
                        PageCount += 1
                    End If
                End If
            Next
            SendEnter() 'check
            SendKeyes("N", 4, 79)
            SendEnter()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If

            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                Exit Sub
                ExitAllAndReturnWithException()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub RemoveBlankAndPISLLine()
        Try
            Dim FlagLineFound = False
            Dim FlagBlankLineFound = False
            For I As Integer = 22 To 10 Step -2
                If GetText(I, 4, 6, True).Replace("-", "") <> "" AndAlso InStr(PISLLINE, GetText(I, 4, 6, True).Replace("-", "")) > 0 Then
                    FlagBlankLineFound = True
                    Exit For
                End If
                Dim check As String = GetText(I, 1, 80).Replace("-", "").Trim.Replace("N", "")
                If GetText(I, 1, 80).Replace("-", "").Trim.Replace("N", "") <> "" And FlagLineFound = False Then
                    FlagLineFound = True
                Else
                    If FlagLineFound Then
                        If GetText(I, 1, 80).Replace("-", "").Trim.Replace("N", "") = "" Then
                            FlagBlankLineFound = True
                            Exit For
                        End If
                    End If
                End If
            Next
            If FlagBlankLineFound = False Then Exit Sub
            SendEnter()
            SendKeyes("EDS", 2, 2)
            SendEnter()
            For KK As Integer = 6 To 18 Step 3

                If GetText(KK, 11, 6, True) = "------" OrElse Trim(GetText(KK, 11, 6, True)) = "" Then
                    Exit For
                End If

                If (GetText(KK, 11, 6, True) = "OI" AndAlso isOI = False) OrElse
                   (GetText(KK, 11, 6, True) = "OIM" AndAlso isOIM = False) OrElse
                   (InStr(PISLLINE, GetText(KK, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(KK, 11, 6, True)) = 0) Then

                    'If (GetText(KK, 11, 6, True) = "OI" AndAlso isOI = False) OrElse
                    '  (GetText(KK, 11, 6, True) = "OIM" AndAlso isOIM = False) OrElse
                    '  (InStr("ACPD", GetText(KK, 11, 6, True)) > 0 AndAlso InStr("ACPD", GetText(KK, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(KK, 11, 6, True)) = 0) Then

                    'SendKeyes("BLANK-", KK, 11)
                    SendKeyes("------", KK, 11)
                    SendKeyes("------", KK, 18)
                    SendKeyes("------", KK, 25)
                    SendKeyes("-----", KK, 32)
                    SendKeyes("--", KK, 38)
                    SendKeyes("-", KK, 41)
                    SendKeyes("-", KK, 43)
                    SendKeyes("---", KK, 45)
                    SendKeyes("------", KK, 49)
                    SendKeyes("--", KK, 56)
                    SendEnter()
                    SendKeyes("N", 4, 79)
                    SendEnter()
                    If GetColorScheme(KK, 11) = 3 AndAlso (InStr(GetText(24, 1, 80), "ENTIRE LINE MUST BE DASHED OFF") > 0 OrElse InStr(GetText(24, 1, 80), "CANNOT DASH OUT UNET LINE") > 0) Then
                        SendKeyes("blank-", KK, 11)
                        SendEnter()

                    Else
                        Dim Count As Integer = 0
                        While GetText(KK, 11, 6, True) = "BLANK-"
                            SendKeyes("------", KK, 11)
                            SendKeyes("------", KK, 18)
                            SendKeyes("------", KK, 25)
                            SendKeyes("-----", KK, 32)
                            SendKeyes("--", KK, 38)
                            SendKeyes("-", KK, 41)
                            SendKeyes("-", KK, 43)
                            SendKeyes("---", KK, 45)
                            'SendKeyes("------", KK, 49)
                            'SendKeyes("--", KK, 56)
                            SendEnter()
                            Count += 1
                            If Count > 5 Then
                                'WriteLog(GetTextRect())
                                strResearchStatus = "Skipped"
                                strResearchComments = "unable to remove pisl line"
                                Exit Sub

                                ExitAllAndReturnWithException()
                            End If
                        End While
                        KK -= 3
                    End If

                Else
                    SendKeyes("--", KK, 69)
                    SendKeyes("--", KK, 72)
                    SendKeyes("--", KK, 75)
                    SendKeyes("--", KK, 78)
                    If IsNumeric(GetText(KK, 49, 9).Replace(" ", ".")) AndAlso CDbl(GetText(KK, 49, 9).Replace(" ", ".")) <> 0.01 Then
                        If InStr("LN,LY", GetText(KK, 45, 2, True)) = 0 Then
                            SendKeyes("--", KK, 45)
                            SendKeyes("---", KK + 1, 20)
                            SendKeyes("---", KK + 1, 50)
                            SendKeyes("---", KK + 1, 54)
                            SendKeyes("------", KK + 1, 6)
                            SendKeyes("--", KK + 1, 13)
                            If GetText(KK, 38, 2) = "30" Then
                                SendKeyes("BLANK-", KK, 11)
                            Else
                                SendKeyes("09", KK, 38)
                            End If
                            SendEnter()
                        End If
                    End If
                    SendKeyes("opn", KK, 2)
                    SendEnter()
                    apiChk()



                    For nn As Integer = 8 To 17 Step 3
                        If Trim(GetText(nn, 7, 4)) = String.Empty Then Exit For
                        SendKeyes("------", nn, 57)
                        SendKeyes("--", nn, 64)
                        SendKeyes("--", nn, 67)
                        SendKeyes("------", nn, 71)
                        SendKeyes("--", nn, 78)
                        If nn = 17 Then
                            EDS6Point5ValueMatch(PreviosValue)
                            SendF8()
                            EDS6Point5ValueMatch(CurrentValue)

                            If EDS6Point5IndexMatch() = False Then
                                nn = 5
                            Else
                                Exit For
                            End If

                        End If



                    Next

                    SendKeyes("Y", 2, 49)
                    SendEnter()
                    If InStr(GetText(24, 1, 80), "ENTER UNITS ON REPEATED LINE") Then
                        strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE"
                        Exit Sub
                    End If
                    SendF3()
                End If
                If KK = 18 Then
                    SendEnter()
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                        Exit For
                    Else
                        KK = 3
                    End If
                End If
            Next
            SendEnter()
            removeUPbyplacingDateinHalfScreen()
            'Check if still PISL line appears in payment screen, it needs to be reject, hence skip the icn
            For II As Integer = 10 To 22 Step 2
                If (InStr("ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,PLPRE,POD,PRE,REF,SSO,TDPD", GetText(II, 4, 5, True)) > 0 AndAlso InStr("PR, IC, NB, SP, WD", GetText(II, 4, 5, True)) = 0) Then
                    'WriteLog(GetTextRect())
                    strResearchStatus = "Skipped"
                    strResearchComments = "unable to remove pisl line"
                    Exit Sub
                    'ExitAllAndReturnWithException()
                End If
            Next
            SendOverrideInScreen("01")
            'SendKeyes("A", 1, 75)
            'SendEnter()
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub removeUPbyplacingDateinHalfScreen()
        Try
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If
            SendF7()
            SendKeyes("R", 22, 61)
            Dim julianDate As String = CurrentICNEDS9(0).JulianDate

            Dim Ldos As String = GetLdos()
            SendKeyes(Left(Ldos, 2), 22, 69)
            SendKeyes(Mid(Ldos, 3, 2), 22, 72)
            SendKeyes("20" & Right(Ldos, 2), 22, 75)
            SendEnter()
            SendKeyes("N", 4, 79)
            SendEnter()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If

            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                Exit Sub
                ExitAllAndReturnWithException()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Function GetLdos() As String
        Dim Temp As String = ""
        For I As Integer = 6 To 18 Step 3
            If GetText(I, 11, 6, True) = "" Then
                Exit For
            End If
            If Temp = "" Then
                Temp = IIf(IsNumeric(GetText(I, 25, 6)), GetText(I, 25, 6), "")
            Else
                'If Trim(ML(I).LDOS) <> "" Then
                If IsNumeric(GetText(I, 25, 6)) Then
                    If CDate(Microsoft.VisualBasic.Left(Temp, 2) & "/" & Mid(Temp, 3, 2) & "/" & Microsoft.VisualBasic.Right(Temp, 2)) < CDate(GetText(I, 25, 2) & "/" & GetText(I, 27, 2) & "/" & GetText(I, 29, 2)) Then
                        Temp = GetText(I, 25, 6)
                    End If
                End If
                'End If
            End If
            If I = 18 Then
                SendEnter()
                SendF8()
                If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                    Exit For
                Else
                    I = 3
                End If
            End If
        Next
        Return Temp
    End Function
    Public Function isLinePresentInEds1(ByRef currEDS1() As EDS1Value, ByVal LineCount As Integer) As Boolean

        'For I As Integer = LBound(ML) To UBound(ML)
        '    If ML(I).Suffix Is Nothing OrElse (CInt(ML(I).Suffix) = Suffix AndAlso ML(I).LineNumber = LineCount) Then Return True
        'Nextc
        If CInt(currEDS1(1).ServiceLineCount) = LineCount Then Return True
        Return False
    End Function
    Public Sub RemoveDenial()
        WriteLog(">>============Entering " + GetCurrentMethod.Name + "")
        Try
            Dim SafetyPageCount As Integer = 0
            Dim Flag As Boolean = False
            For I As Integer = 10 To 22 Step 2
                Dim BC = GetText(I, 41, 5).Replace(" ", "0").Replace("-", "0") + "." + GetText(I, 47, 2).Replace(" ", "0").Replace("-", "0")
                Dim NC = GetText(I, 50, 5).Replace(" ", "0").Replace("-", "0") + "." + GetText(I, 56, 2).Replace(" ", "0").Replace("-", "0")
                If IsNumeric(BC) AndAlso IsNumeric(NC) AndAlso CDbl(NC) > 0.01 Then
                    Flag = True
                    Exit For
                End If
            Next
            If Flag = False Then Exit Sub
            SendKeyes("EDS", 2, 2)
            SendEnter()
            For KK As Integer = 6 To 18 Step 3

                If GetText(KK, 11, 6, True) = "" Then
                    Exit For
                End If

                If InStr(PISLLINE, GetText(KK, 11, 6, True)) = 0 Then
                    If GetColorScheme(KK + 1, 6) <> 0 Then
                        SendKeyes("opn", KK, 2)
                        SendEnter()
                        'SendKeyes("01", 7, 45)
                        SendKeyes("------", 7, 58)
                        SendKeyes("--", 7, 65)
                        SendKeyes("--", 7, 68)
                        SendEnter()
                        SendF3()
                        If CDbl(GetText(KK, 59, 6).Replace("-", "0") + "." + GetText(KK, 66, 2).Replace("-", "0")) > 0 Then
                            If GetColorScheme(KK, 59) <> 0 Then
                                'flagerror = True
                                Exit Sub
                            Else
                                SendKeyes("------", KK, 59)
                                SendKeyes("--", KK, 66)
                                SendEnter()
                            End If
                        End If
                    Else
                        'End If
                        SendKeyes("000000", KK + 1, 6)
                        SendKeyes("---", KK + 1, 13)
                        SendKeyes("000", KK + 1, 20)
                        SendKeyes("---", KK + 1, 50)
                        SendKeyes("---", KK + 1, 54)
                        SendKeyes("01", KK, 38)
                        If IsNumeric(GetText(KK, 49, 9).Replace(" ", ".")) AndAlso CDbl(GetText(KK, 49, 9).Replace(" ", ".")) = 0.01 Then
                            'SendKeyes("000000", KK + 1, 6)
                            'SendKeyes("01", KK + 1, 13)
                            'SendKeyes("LN-", KK + 1, 20)
                            'PlaceLNorLYinHalfScreen(KK)
                        End If
                        SendEnter()
                        If GetColorScheme(KK + 1, 6) = 3 Then
                            SendKeyes("opn", KK, 2)
                            SendEnter()
                            'SendKeyes("01", 7, 45)
                            SendKeyes("------", 7, 58)
                            SendKeyes("--", 7, 65)
                            SendKeyes("--", 7, 68)
                            SendEnter()
                            SendF3()
                            If CDbl(GetText(KK, 59, 6).Replace("-", "0") + "." + GetText(KK, 66, 2).Replace("-", "0")) > 0 Then
                                If GetColorScheme(KK, 59) <> 0 Then
                                    'flagerror = True
                                    Exit Sub
                                Else
                                    SendKeyes("------", KK, 59)
                                    SendKeyes("--", KK, 66)
                                    SendEnter()
                                End If
                            End If
                        End If
                    End If
                    ''My Code'''''''''''''''''''''''''''''''''''''''''''''''''''
                    SendKeyes("opn", KK, 2)
                    SendEnter()
                    apiChk()



                    For nn As Integer = 8 To 17 Step 3
                        If Trim(GetText(nn, 7, 4)) = String.Empty Then Exit For
                        SendKeyes("------", nn, 57)
                        SendKeyes("--", nn, 64)
                        SendKeyes("--", nn, 67)
                        SendKeyes("------", nn, 71)
                        SendKeyes("--", nn, 78)
                        If nn = 17 Then
                            EDS6Point5ValueMatch(PreviosValue)
                            SendF8()
                            EDS6Point5ValueMatch(CurrentValue)

                            If EDS6Point5IndexMatch() = False Then
                                nn = 5
                            Else
                                Exit For
                            End If

                        End If



                    Next

                    SendKeyes("Y", 2, 49)
                    SendEnter()
                    apiChk()
                    apiChk()
                    SendF3()
                    If InStr(GetText(24, 2, 80).Trim(), "INVALID ENTRY") AndAlso GetColorScheme(8, 44) = 3 AndAlso Not IsNumeric(Trim(GetText(8, 44, 2))) Then
                        SendKeyes("--", 8, 44)
                        autECLPSObj.SendKeys("[enter]")
                        apiChk()
                        SendF3()
                    End If

                End If

                If KK = 18 Then
                    SendEnter()
                    If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                        SendKeyes("N", 4, 79)
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                    Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                        Exit For
                    ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                        Exit For
                    Else
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            strResearchComments = $"{GetCurrentMethod.Name} + Error: Force exit due to infinite loop detection (Max Pages Exceeded)"
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                        KK = 3
                    End If
                End If
            Next
            'SendEnter()
            'removeUPbyplacingDateinHalfScreen()
            SendKeyes("N", 4, 79)
            SendEnter()
            apiChk()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        Finally
            WriteLog(">>============Exit " + GetCurrentMethod.Name + "")
        End Try
    End Sub
    Public Function isLineMatchInEDS1() As Boolean
        'If GetText(2, 2, 2) = "MP" Then
        '    SendKeyes("EDS", 2, 2)
        '    SendEnter()
        '    apiChk()
        'End If

        'For ee As Integer = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)

        'Next
        Dim FlagStatus As Boolean = False
        Dim oo As Integer = 0
        Dim Revcode As String = String.Empty
        Dim ProcCode As String = String.Empty
        Dim DateOfService As String = String.Empty
        Dim Charge As String = String.Empty
        Dim ERevcode As String = String.Empty
        Dim EProcCode As String = String.Empty
        Dim EDateOfService As String = String.Empty
        Dim ECharge As String = String.Empty
        Dim FCharge As Double = 0.0
        Dim EFCharge As Double = 0.0
        Dim BillCharge As String = String.Empty
        Dim EBillCharge As String = String.Empty
        Dim M1 As String = String.Empty
        Dim EM1 As String = String.Empty
        Dim M2 As String = String.Empty
        Dim EM2 As String = String.Empty
        Dim gg As Integer = 0
        Dim ll As Integer = 0
        Try
            If CurrentICNEDS1.Length = CurrentEDS6Point5Screen.Length Then
                oo = CurrentICNEDS1.Length - 1
            Else
                FlagStatus = False
                Return FlagStatus
            End If
            For kk As Integer = 0 To oo
                For gg = LBound(CurrentICNEDS1(kk).ServiceLine.RevCodeCount) To UBound(CurrentICNEDS1(kk).ServiceLine.RevCodeCount)

                    Revcode = CurrentICNEDS1(kk).ServiceLine.RevCode(gg) + "," + Revcode
                    ProcCode = CurrentICNEDS1(kk).ServiceLine.CPTCode(gg) + "," + ProcCode
                    DateOfService = CurrentICNEDS1(kk).ServiceLine.DateofService(gg) + "," + DateOfService
                    Charge = CurrentICNEDS1(kk).ServiceLine.BillCharge(gg)
                    FCharge = CDbl(Charge)
                    BillCharge = FCharge.ToString() + "," + BillCharge
                    M1 = CurrentICNEDS1(kk).ServiceLine.M1(gg) + "," + M1
                    M2 = CurrentICNEDS1(kk).ServiceLine.M2(gg) + "," + M2

                Next
                For ll = LBound(CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.LineCount) To UBound(CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.LineCount)

                    ERevcode = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.RevCode(ll) + "," + ERevcode
                    EProcCode = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.ProcCode(ll) + "," + EProcCode
                    EDateOfService = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.DateOfService(ll) + "," + EDateOfService
                    ECharge = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.Charge(ll)
                    EFCharge = CDbl(ECharge)
                    EBillCharge = EFCharge.ToString() + "," + EBillCharge
                    EM1 = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.M1(ll) + "," + EM1
                    EM2 = CurrentEDS6Point5Screen(kk).EDS6ServiceLineCount.M2(ll) + "," + EM2

                Next
                If InStr(Revcode, ERevcode) > 0 AndAlso InStr(ProcCode, EProcCode) > 0 AndAlso InStr(DateOfService, EDateOfService) > 0 AndAlso InStr(BillCharge, EBillCharge) > 0 AndAlso InStr(M1, EM1) > 0 AndAlso InStr(M2, EM2) > 0 Then
                    FlagStatus = True
                Else
                    FlagStatus = False
                    Return FlagStatus
                    Exit For
                End If

                gg = 0
                ll = 0
                Revcode = String.Empty
                ProcCode = String.Empty
                DateOfService = String.Empty
                BillCharge = String.Empty
                ERevcode = String.Empty
                EProcCode = String.Empty
                EDateOfService = String.Empty
                EBillCharge = String.Empty
                M1 = String.Empty
                M2 = String.Empty
                EM1 = String.Empty
                EM2 = String.Empty
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " =>" + ex.Message.ToString)
            strResearchComments = "EXCEPTION=>" + ex.Message.ToString
        End Try
        Return FlagStatus
    End Function
    Public Sub PutOriginalValue()
        WriteLog("==============IN " + GetCurrentMethod.Name + "============")
        Try
            Dim count As Integer
            Dim check As String = ""
            count = 6
            For dd As Integer = LBound(Original) To UBound(Original)
                check = ""
                If check = "PISLINE Present" Then
                    dd = 0
                End If

                If GetText(2, 2, 2) = "MP" Then
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    apiChk()
                End If
                If (GetText(count, 11, 6, True) = "OI") OrElse
               (GetText(count, 11, 6, True) = "OIM") OrElse (GetText(count, 11, 6, True) = "BLANK-") OrElse
               (InStr(PISLLINE, GetText(count, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(count, 11, 6, True)) = 0) Then
                    check = "PISLINE Present"
                Else
                    SendKeyes(Original(dd).OV, count, 38)
                    SendKeyes(Original(dd).RemarkCode, count, 45)
                    SendKeyes("opn", count, 2)
                    SendEnter()
                    apiChk()
                    PutValue(dd)
                    If count = 18 Then
                        SendF8()
                        count = 3
                    End If
                End If
                count += 3
            Next
            SendKeyes("N", 4, 79)
            SendEnter()
            apiChk()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            WriteLog(GetTextRect(), True)
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
            WriteLog(GetTextRect())
            strResearchComments = GetCurrentMethod.Name() + " exception: " + ex.Message.ToString
            Exit Sub
        End Try
        WriteLog("==============EXIT " + GetCurrentMethod.Name + "============")
    End Sub
    Public Sub PutValue(ByVal dd As Integer)
        Try
            Dim count As Integer
            count = 8
            For ee As Integer = LBound(Original(dd).EDSLineCount.RevCode) To UBound(Original(dd).EDSLineCount.RevCode)
                Dim value1() As String = Original(dd).EDSLineCount.NoteCover(ee).Split(".")
                SendKeyes(value1(0), count, 57)
                SendKeyes(value1(1), count, 64)
                SendKeyes(Original(dd).EDSLineCount.RemarkCode(ee), count, 67)
                If count = 17 Then
                    SendF8()
                    count = 5
                End If
                count += 3

            Next
            SendF3()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Sub GetEDSScreenOriginalData(ByRef D() As OriginalEDSValue, maxSuffix As Integer)
        Try
            Dim SafetyPageCount As Integer = 0
            Dim SafetyPageCountForEDS6point5 As Integer = 0
            Do
                WriteLog("===============ENTERING " + GetCurrentMethod.Name + "==============")
                If GetText(2, 2, 2) = "MP" Then
                    autECLPSObj.SendKeys("EDS", 2, 2)
                    SendEnter()
                    apiChk()
                Else
                    SendClear()
                    SendClear()
                    pullEDSScreen(CurrICN, 2)
                    SendKeyes("A", 1, 76)
                    SendEnter()
                    apiChk()

                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                End If
                If strResearchComments = "eds not valid" Then
                    Exit Sub
                End If

                If GetText(3, 2, 5) = "EDS 6" Then       'Added on 18th Oct,22 to send Y in bypass history field
                    For dd As Integer = 6 To 18 Step 3
                        If (GetText(dd, 11, 6, True) = "OI") OrElse
                       (GetText(dd, 11, 6, True) = "OIM") OrElse (GetText(dd, 11, 6, True) = "BLANK-") OrElse
                       (InStr(PISLLINE, GetText(dd, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(dd, 11, 6, True)) = 0) Then
                            RemovePISLline()
                            apiChk()
                            If GetText(2, 2, 2) = "MP" Then
                                autECLPSObj.SendKeys("EDS", 2, 2)
                                SendEnter()
                                apiChk()
                            End If
                        End If
                        SendKeyes("opn", 6, 2)
                        SendEnter()
                        apiChk()
                        If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                            For cc As Integer = 7 To 19 Step 3
                                If GetText(cc - 1, 8, 9) <> "         " Then
                                    autECLPSObj.SendKeys("--------", cc, 6)
                                    apiChk()
                                    autECLPSObj.SendKeys("---", cc, 20)
                                    apiChk()
                                    autECLPSObj.SendKeys("--------", cc, 28)
                                    apiChk()
                                    autECLPSObj.SendKeys("---", cc, 42)
                                    apiChk()
                                    autECLPSObj.SendKeys("------", cc, 50)
                                    apiChk()
                                End If
                            Next
                            SendEnter()
                            apiChk()
                        End If
                        If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                            SendKeyesFor01(OriginalDraft, 4, 32)
                            SendKeyesFor01(OriginalICN, 4, 15)
                            SendEnter()
                            apiChk()
                            SendKeyes("A", 1, 75)
                            SendEnter()
                            apiChk()
                            If GetText(2, 2, 2) = "MP" Then
                                autECLPSObj.SendKeys("EDS", 2, 2)
                                SendEnter()
                                apiChk()
                            End If
                        End If
                        If GetText(3, 2, 5) = "EDS 6" Then
                            SendKeyes("opn", 6, 2)
                            SendEnter()
                            apiChk()
                        End If
                        If GetText(2, 73, 7) = "EDS 6.5" Then
                            SendKeyes("Y", 2, 49)
                            SendEnter()
                            apiChk()
                            SendF3()
                        End If
                        Exit For
                    Next
                End If

                ReDim D(10)
                Dim LineCount As Integer = -1
                Dim x As Integer = 0
                WriteLog(GetTextRect())
                For dd As Integer = 6 To 18 Step 3
                    If GetText(3, 2, 5) = "EDS 6" Then
                        If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                            RemoveNotCoveredFromEDS6Point5()
                            SendEnter()
                            If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                For cc As Integer = 7 To 19 Step 3
                                    If GetText(cc - 1, 8, 9) <> "         " Then
                                        autECLPSObj.SendKeys("--------", cc, 6)
                                        apiChk()
                                        autECLPSObj.SendKeys("---", cc, 20)
                                        apiChk()
                                        autECLPSObj.SendKeys("--------", cc, 28)
                                        apiChk()
                                        autECLPSObj.SendKeys("---", cc, 42)
                                        apiChk()
                                        autECLPSObj.SendKeys("------", cc, 50)
                                        apiChk()
                                    End If
                                Next
                                SendEnter()
                                apiChk()
                            End If
                            If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                                SendKeyesFor01(OriginalDraft, 4, 32)
                                SendKeyesFor01(OriginalICN, 4, 15)
                                SendEnter()
                                apiChk()
                                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                    For cc As Integer = 7 To 19 Step 3
                                        If GetText(cc - 1, 8, 9) <> "         " Then
                                            autECLPSObj.SendKeys("--------", cc, 6)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 20)
                                            apiChk()
                                            autECLPSObj.SendKeys("--------", cc, 28)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 42)
                                            apiChk()
                                            autECLPSObj.SendKeys("------", cc, 50)
                                            apiChk()
                                        End If
                                    Next
                                    SendEnter()
                                    apiChk()
                                End If
                                SendKeyes("A", 1, 75)
                                SendEnter()
                                apiChk()
                                If GetText(2, 2, 2) = "MP" Then
                                    autECLPSObj.SendKeys("EDS", 2, 2)
                                    SendEnter()
                                    apiChk()
                                End If
                            End If

                        End If
                        If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") <> True And GetText(24, 2, 80).Trim.Contains("NO MORE RECORDS") <> True And GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") <> True And GetText(24, 2, 80).Trim <> "" And InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") = 0 Then
                            strResearchComments = "Manual intervention required: " + GetText(24, 2, 80).Trim
                            SendClear()
                            If GetText(2, 2, 2) <> "MP" Then SendClear()
                            Exit Sub
                        End If
                    End If
                    If (GetText(dd, 11, 6, True) = "OI") OrElse
                       (GetText(dd, 11, 6, True) = "OIM") OrElse (GetText(dd, 11, 6, True) = "BLANK-") OrElse
                       (InStr(PISLLINE, GetText(dd, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(dd, 11, 6, True)) = 0) Then
                        Continue For
                    Else
                        Dim checkservline As String = Trim(GetText(dd, 2, 3))
                        If checkservline <> "---" Then Exit For
                        If GetText(dd, 2, 3) <> "" Then
                            LineCount += 1
                            If LineCount = 0 Then
                                ReDim D(LineCount)
                            Else
                                ReDim Preserve D(LineCount)
                            End If
                            D(LineCount).EDSLineCount.LineTranslation = Trim(GetText(dd, 11, 5))
                            D(LineCount).OV = Trim(GetText(dd, 38, 2))
                            If Not IsNumeric(D(LineCount).OV) Then
                                D(LineCount).OV = "--"
                            End If
                            D(LineCount).RemarkCode = Trim(GetText(dd, 45, 2))
                            D(LineCount).NoteCover = Trim(GetText(dd, 59, 6)) + "." + Trim(GetText(dd, 66, 2))
                            autECLPSObj.SendKeys("opn", dd, 2)
                            SendEnter()
                            apiChk()

                            If GetText(3, 2, 5) = "EDS 6" Then
                                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") <> True And GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") <> True And GetText(24, 2, 80).Trim <> "" Then
                                    strResearchComments = "Manual intervention required" + GetText(24, 2, 80).Trim
                                    SendClear()
                                    If GetText(2, 2, 2) <> "MP" Then SendClear()
                                    Exit Sub
                                End If

                                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                                    For cc As Integer = 7 To 19 Step 3
                                        If GetText(cc - 1, 8, 9) <> "         " Then
                                            autECLPSObj.SendKeys("--------", cc, 6)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 20)
                                            apiChk()
                                            autECLPSObj.SendKeys("--------", cc, 28)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 42)
                                            apiChk()
                                            autECLPSObj.SendKeys("------", cc, 50)
                                            apiChk()
                                        End If
                                    Next
                                    SendEnter()
                                    apiChk()
                                End If

                                If GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") Then
                                    autECLPSObj.SendKeys("N", 4, 79)
                                    SendEnter()
                                    apiChk()
                                End If
                            End If

                        End If
                        Dim hh As Integer = 0
                        x = 0
                        Dim count As Integer
                        If GetText(2, 73, 7) = "EDS 6.5" Then                   'Added on 3rd March,24 to send Y in bypass field everytime in eds 6.5
                            SendKeyes("Y", 2, 49)
                            SendEnter()
                        End If
                        For hh = 8 To 17 Step 3
                            Dim checkline As String = Trim(GetText(hh, 2, 3))
                            Dim DateOfservice As String = String.Empty
                            If checkline <> "---" Then
                                SendF3()
                                If GetText(24, 2, 80).Trim() = "ENTER UNITS ON REPEATED LINE" Then
                                    strResearchComments = "ENTER UNITS ON REPEATED LINE"
                                    SendClear()
                                    If GetText(2, 2, 2) <> "MP" Then SendClear()
                                    Exit Sub
                                End If
                                If InStr(GetText(24, 2, 80).Trim(), "INVALID ENTRY") AndAlso GetColorScheme(8, 44) = 3 AndAlso Not IsNumeric(Trim(GetText(8, 44, 2))) Then
                                    SendKeyes("--", 8, 44)
                                    autECLPSObj.SendKeys("[enter]")
                                    apiChk()
                                    SendF3()
                                End If

                                If GetText(24, 2, 80).Trim = "MUST NOT BE BLANK" Then
                                    For ff = 8 To (hh - 3) Step 3
                                        If GetColorScheme(ff, 67) = 3 Then
                                            SendKeyes("000", ff, 67)
                                            apiChk()
                                            SendEnter()
                                            apiChk()
                                            If GetText(24, 2, 80).Trim = "RC NOT VALID" Then
                                                SendEnter()
                                                apiChk()
                                            End If
                                        End If
                                    Next
                                    SendF3()
                                End If
                                Exit For
                            End If
                            If x = 0 Then
                                ReDim D(LineCount).EDSLineCount.RevCode(x)
                                ReDim D(LineCount).EDSLineCount.RemarkCode(x)
                                ReDim D(LineCount).EDSLineCount.NoteCover(x)
                                count = 1
                            Else
                                ReDim Preserve D(LineCount).EDSLineCount.RevCode(x)
                                ReDim Preserve D(LineCount).EDSLineCount.RemarkCode(x)
                                ReDim Preserve D(LineCount).EDSLineCount.NoteCover(x)
                                count += 1
                            End If
                            D(LineCount).EDSLineCount.RevCode(x) = Trim(GetText(hh, 7, 4))
                            D(LineCount).EDSLineCount.RemarkCode(x) = Trim(GetText(hh, 67, 2))
                            D(LineCount).EDSLineCount.NoteCover(x) = Trim(GetText(hh, 57, 6)) + "." + Trim(GetText(hh, 64, 2))
                            x += 1
                            If hh = 17 Then
                                WriteLog(GetTextRect)
                                EDS6Point5ValueMatch(PreviosValue)
                                SendF8()
                                If GetText(24, 2, 80).Trim = "MUST NOT BE BLANK" Then
                                    For ff = 8 To hh Step 3
                                        If GetColorScheme(ff, 67) = 3 Then
                                            SendKeyes("000", ff, 67)
                                            apiChk()
                                            SendEnter()
                                            apiChk()
                                            If GetText(24, 2, 80).Trim = "RC NOT VALID" Then
                                                SendEnter()
                                                apiChk()
                                            End If
                                        End If
                                    Next
                                End If
                                EDS6Point5ValueMatch(CurrentValue)

                                If EDS6Point5IndexMatch() = False Then
                                    hh = 5
                                Else
                                    SendF3()
                                    Exit For
                                End If
                            End If
                        Next
                        If dd = 18 Then
                            WriteLog(GetTextRect())
                            Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                            SendF8()
                            If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                Exit For
                            ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                Exit For
                            Else
                                SafetyPageCount = SafetyPageCount + 1
                                If SafetyPageCount > 50 Then
                                    WriteLog(GetCurrentMethod.Name + “ Error: Loop terminated to prevent infinite spike (Max pages exceeded).”)
                                    strResearchComments = "Error: Loop terminated to prevent infinite spike (Max pages exceeded)"
                                    Exit For
                                End If
                            End If

                            'If GetText(24, 1, 80).Trim = "" Then
                            '    Dim cnt As Integer = 1
                            '    Do
                            '        SendF8()
                            '        cnt = cnt + 1
                            '    Loop While (cnt <= 10)
                            '    strResearchComments = "Next page NA for EDS 6/6.5"
                            '    Exit For
                            'End If
                            'SendF8()

                            dd = 3
                        End If
                    End If

                Next

                If strResearchComments = "Next page NA for EDS 6/6.5" Then
                    Exit Sub
                End If
                'SendKeyes("N", 4, 79)
                autECLPSObj.SendKeys("N", 4, 79)
                apiChk()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()

                'SendKeyes("A", 1, 75)
                autECLPSObj.SendKeys("A", 1, 75)
                'SendEnter()
                'apiChk()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
                apiChk()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                    SendKeyesFor01(OriginalDraft, 4, 32)
                    SendKeyesFor01(OriginalICN, 4, 15)
                    SendEnter()
                    apiChk()
                    autECLPSObj.SendKeys("A", 1, 75)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                        SendClear()
                        'SendClear()

                    End If
                End If
                If GetText(2, 2, 3) = "MHI" OrElse GetText(2, 2, 3) = "MHN" Then
                    strResearchComments = "Already Processed Claim"
                    Exit Sub
                End If
                If GetText(3, 2, 5) = "EDS 6" Then
                    If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") <> True And GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") <> True And GetText(24, 2, 80).Trim <> "" Then
                        strResearchComments = "Manual intervention required" + GetText(24, 2, 80).Trim
                        SendClear()
                        If GetText(2, 2, 2) <> "MP" Then SendClear()
                        Exit Sub
                    End If

                    If GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") Then
                        autECLPSObj.SendKeys("N", 4, 79)
                        apiChk()
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                End If

            Loop While strResearchComments.Trim.Contains("HIT CLEAR TO RETURN TO TOPS")

            If strResearchComments.Trim.Contains("Manual intervention required") Then
                Exit Sub
            End If
            WriteLog(GetTextRect)
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " =>" + ex.Message.ToString)
        Finally
            WriteLog("===============EXITING " + GetCurrentMethod.Name + "==============")
        End Try

    End Sub
    Public Sub EDSScreenLineCount(ByRef EDS6Count() As EDS6Point5Screen, maxSuffix As Integer)
        'For nn As Integer = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
        '    Dim kkkk As Integer = CurrentICNEDS1(nn).ServiceLine.RevCode.Length
        'Next
        Try
            For i As Integer = 1 To maxSuffix

            Next
            Dim SafetyPageCount As Integer = 0
            WriteLog("Enter EDSScreenLineCount..............")
            SendClear()
            SendClear()
            pullEDSScreen(CurrICN, 2)
            SendKeyes("A" + maxSuffix.ToString, 1, 76)
            SendEnter()
            apiChk()

            autECLPSObj.SendKeys("EDS", 2, 2)
            SendEnter()
            apiChk()
            Dim kkkk As Integer
            '''''''''''''''''''''''''''''''''Added by Sourav''''''''''''''''17/05/2023'''''''''''''''''''''''''''''''''''''''''''''''''''''            
            For aa As Integer = 10 To 22 Step 2
                If GetText(aa, 1, 2) = "--" Then Exit For
                If InStr(PISLLINE, GetText(aa, 4, 6, True)) > 0 And Trim(GetText(aa, 1, 80).Replace("-", "")) <> "" Then
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    apiChk()
                    For KK As Integer = 6 To 18 Step 3
                        If InStr(PISLLINE, GetText(KK, 11, 6, True)) > 0 And Trim(GetText(KK, 1, 80)) <> "" Then
                            'SendKeyes("BLANK-", KK, 11)
                            SendKeyes("------", KK, 11)
                            SendKeyes("------", KK, 18)
                            SendKeyes("------", KK, 25)
                            SendKeyes("-----", KK, 32)
                            SendKeyes("--", KK, 38)
                            SendKeyes("-", KK, 41)
                            SendKeyes("-", KK, 43)
                            SendKeyes("--", KK, 45)
                            SendKeyes("------", KK, 49)
                            SendKeyes("--", KK, 56)
                        End If

                        If KK = 18 Then
                            'SendEnter()
                            SendF8()
                            If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                Exit For
                            Else
                                SafetyPageCount = SafetyPageCount + 1

                                ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                If SafetyPageCount > 50 Then
                                    WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                    Exit For
                                End If
                                KK = 3
                            End If
                        End If
                    Next
                    SendEnter()
                    SendKeyes("N", 4, 79)
                    SendEnter()
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    If GetText(2, 2, 2) = "MP" Then Exit For
                End If
            Next
            If GetText(2, 2, 2) <> "MP" Then
                SendClear()
                SendClear()
                pullEDSScreen(CurrentClaimDetails.ICN, 1)
                SendKeyes("A" + maxSuffix.ToString, 1, 76)
                SendEnter()
            End If
            '''''''''''''''''''''''''''''''''''''''Till this'''''''''''''''''''''''''''''''''''''''''''''''''''''
            'For ff As Integer = 1 To MaxSuffix
            '    If ff = 1 Then
            '        SendKeyes("A", 1, 76)
            '        SendEnter()
            '        apiChk()
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
                apiChk()
            End If
            If GetText(3, 2, 5) <> "EDS 6" Then
                strResearchComments = "Cannot move to EDS 6"
            End If
            If strResearchComments = "eds not valid" Then
                Exit Sub
            End If
            ReDim EDS6Count(maxSuffix)
            Dim LineCount As Integer = -1
            Dim x As Integer = 0
            SafetyPageCount = 0
            For dd As Integer = 6 To 18 Step 3
                Dim checkservline As String = Trim(GetText(dd, 2, 80))
                If checkservline = "" Then Exit For
                If GetText(dd, 2, 3) <> "" Then
                    LineCount += 1
                    If LineCount = 0 Then
                        ReDim EDS6Count(LineCount)
                    Else
                        ReDim Preserve EDS6Count(LineCount)
                    End If
                    EDS6Count(LineCount).EDS6ServiceLineCount.EDS6ServiceTranslation = Trim(GetText(dd, 11, 5))
                    SendKeyes("opn", dd, 2)
                    SendEnter()
                    apiChk()
                    kkkk = CurrentICNEDS1(LineCount).ServiceLine.RevCode.Length
                End If

                Dim hh As Integer = 0
                x = 0
                Dim count As Integer
                If GetText(2, 73, 7) = "EDS 6.5" Then
                    SendKeyes("Y", 2, 49)
                    SendEnter()
                    apiChk()
                End If
                For hh = 8 To 17 Step 3
                    Dim checkline As String = Trim(GetText(hh, 2, 3))
                    Dim DateOfservice As String = String.Empty
                    If checkline <> "---" Then
                        SendF3()
                        Exit For
                    End If
                    If x = 0 Then
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.LineCount(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.RevCode(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.Charge(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.DateOfService(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.ProcCode(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.M1(x)
                        ReDim EDS6Count(LineCount).EDS6ServiceLineCount.M2(x)
                        count = 1

                    Else
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.LineCount(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.RevCode(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.Charge(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.DateOfService(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.ProcCode(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.M1(x)
                        ReDim Preserve EDS6Count(LineCount).EDS6ServiceLineCount.M2(x)
                        count += 1
                    End If
                    EDS6Count(LineCount).EDS6ServiceLineCount.LineCount(x) = x
                    EDS6Count(LineCount).EDS6ServiceLineCount.RevCode(x) = Trim(GetText(hh, 7, 4))
                    EDS6Count(LineCount).EDS6ServiceLineCount.Charge(x) = Trim(GetText(hh, 47, 6)) + "." + Trim(GetText(hh, 54, 2))
                    DateOfservice = Trim(GetText(hh, 33, 4))
                    If DateOfservice = "----" Then
                        DateOfservice = "0000"
                    End If
                    EDS6Count(LineCount).EDS6ServiceLineCount.DateOfService(x) = DateOfservice
                    EDS6Count(LineCount).EDS6ServiceLineCount.ProcCode(x) = Trim(GetText(hh, 13, 5))
                    EDS6Count(LineCount).EDS6ServiceLineCount.M1(x) = Trim(GetText(hh, 20, 2))
                    EDS6Count(LineCount).EDS6ServiceLineCount.M2(x) = Trim(GetText(hh, 23, 2))
                    If Trim(GetText(hh, 57, 6)) <> "------" Then
                        SendKeyes("--------", hh, 57)
                    End If
                    x += 1
                    If hh = 17 Then
                        Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                        If count >= kkkk Then
                            SendF3()
                            Exit For
                        End If
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                            SendF3()
                            Exit For
                        ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                            SendF3()
                            Exit For
                        Else
                            hh = 5
                        End If
                    End If
                Next
                If dd = 18 Then
                    Dim Myscreen6 As String = GetTextRect(6, 1, 18, 79)
                    SendF8()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                        Exit For
                    ElseIf GetTextRect(6, 1, 18, 79) = Myscreen6 Then
                        Exit For
                    Else
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                        dd = 3
                    End If
                End If
            Next

            PutJulianDate()
            If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                SendKeyesFor01(OriginalDraft, 4, 32)
                SendKeyesFor01(OriginalICN, 4, 15)
                SendEnter()
                apiChk()
            End If
            SendKeyes("N", 4, 79)
            SendEnter()
            apiChk()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            WriteLog("Exit EDSScreenLineCount-------------------")
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Function CalculateDateMonth(ByVal value As String) As String
        Dim days As String
        Dim Month As Integer
        Dim FinalDay As Integer
        Dim FinalYear As String
        Dim FullDateTime As String
        Dim FinalMonth As String = String.Empty
        Dim FinalCalDay As String = String.Empty
        days = String.Empty
        FullDateTime = String.Empty
        Month = 0
        FinalDay = 0
        FinalYear = "20" + value.Substring(0, 2)
        days = value.Substring(2, 3)
        Dim Day As Integer
        Day = CInt(days)
        Month = Day / 30
        If Month.ToString().Length = 1 Then
            FinalMonth = "0" + Month.ToString()
        Else
            FinalMonth = Month.ToString()
        End If
        FinalDay = Day Mod 30
        If FinalDay.ToString().Length = 1 Then
            FinalCalDay = "0" + FinalDay.ToString()
        Else
            FinalCalDay = FinalDay.ToString()
        End If
        FullDateTime = FinalMonth + "/" + FinalCalDay + "/" + FinalYear.ToString()
        Return FullDateTime
    End Function
    Public Function CheckClaimAlreadyProcess(ByVal MaxSuffix As Integer) As Boolean
        Dim FlagCheck As Boolean = False
        'For dd As Integer = 1 To MaxSuffix
        Dim count As Integer
        Dim CRcheck As Boolean = False
        For kk As Integer = LBound(CurrentICNEDS9) To UBound(CurrentICNEDS9)
            For ll As Integer = LBound(CurrentICNEDS9(kk).LineCount) To UBound(CurrentICNEDS9(kk).LineCount)
                Dim status As String = CurrentICNEDS9(kk).Status(ll)
                Dim linecolour As Integer = CurrentICNEDS9(kk).LineColor(ll)
                Dim NBR As String = CurrentICNEDS9(kk).NBR(ll)
                If NBR = String.Empty Then
                    NBR = "0"
                End If
                Dim PC As String = CurrentICNEDS9(kk).PC(ll)

                'If CInt(NBR) = dd Then

                If linecolour = 1 Then
                    count = count + 1
                    If UBound(CurrentICNEDS9) = count Then
                        FlagCheck = True
                    End If
                End If
                'If linecolour = 1 OrElse status = "MPAY" OrElse status = "CPAY" OrElse status = "APAY" OrElse PC <> "CR" Then
                If status = "MPAY" OrElse status = "CPAY" OrElse status = "APAY" Then
                    FlagCheck = True
                    'ElseIf linecolour <> 1 OrElse status <> "MPAY" OrElse status <> "CPAY" OrElse status <> "APAY" OrElse PC = "CR" Then
                    'FlagCheck = False
                End If
                If PC = "CR" Then
                    CRcheck = True
                End If

                'End If
            Next
        Next
        If CRcheck = False Then
            FlagCheck = True
        End If
        'Next
        Return FlagCheck
    End Function
    Public Sub HardPendTA()
        Try
            If GetText(3, 2, 5) = "EDS 6" Then
                SendClear()
                If GetText(24, 1, 80).Contains("DO A OVER") = True And GetText(2, 2, 2) <> "MP" Then
                    SendClear()
                End If
            End If
            'If GetText(2, 2, 2) <> "MP" Then
            '    pullEDSScreen(CurrICN, 9)
            '    SendKeyes("A", 76)
            '    SendEnter()
            'End If
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("PTA", 2, 2)
                SendEnter()
                apiChk()
            End If
            If GetText(2, 2, 3) = "MPP" Then strResearchComments = ""
            If GetText(24, 1, 80).Contains("E125INVALID AMOUNT") = True Then
                strResearchComments = "E125INVALID AMOUNT"
                Exit Sub
            End If
            pullEDSScreen(CurrICN, 9)
            apiChk()
            SendKeyes("A", 1, 76)
            SendEnter()
            apiChk()
            If GetText(2, 2, 2) = "MP" Then
                autECLPSObj.SendKeys("EDS", 2, 2)
                SendEnter()
                apiChk()
            End If
        Catch ex As Exception
            WriteLog("EXCEPTION FROM " + GetCurrentMethod.Name())
        End Try
    End Sub
    Public Function WorkONPaymentScreen(ByVal MaxSuffix As Integer, ByVal icn As String, ByVal Payloc As String) As String
        WriteLog("---------------ENTERING " + GetCurrentMethod.Name + "------------------------")
        Dim FinalComment As String = String.Empty
        Dim EditWarning As String
        Dim Count As Integer
        Dim CheckisEDSLineMatch As Boolean = False
        Try
            'If MaxSuffix > 1 Then
            '    FinalComment = "Please Run on Single suffix ICN Right Now"
            '    Return FinalComment
            'End If
            If GetText(2, 2, 2) <> "MP" Then
                PullEDS2(icn, 1)
                SendKeyes("A", 1, 76)
                SendEnter()
            End If
            If GetText(2, 2, 2) = "MP" Then
                EditWarning = String.Empty
                StoreOriginalDraftandICN(icn, Payloc)
                'GetEDSScreenOriginalData(Original, MaxSuffix)
                If strResearchComments.Contains("EOC DATA DIFFERENT") = True Then
                    PullEDS2(icn, 2)
                    SendKeyes("A", 1, 76)
                    SendEnter()
                    HardPendTA()
                    If strResearchComments.Contains("E125INVALID AMOUNT") Then Return strResearchComments
                End If
                If strResearchComments = "Next page NA for EDS 6/6.5" Then
                    Return strResearchComments
                End If
                If strResearchComments.Trim.Contains("Manual intervention required") Then
                    Return strResearchComments
                End If
                If strResearchComments.Trim.Contains("Already Processed Claim") Then
                    Return strResearchComments
                End If
                If strResearchComments.Trim.Contains("ENTER UNITS ON REPEATED LINE") Then
                    Return strResearchComments
                End If
                GetDraftDetailsInMemory(CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount)
                ReadPaymentScreenForLog(CurrentPaymentScreenForLog, MaxSuffix)
                If InStr(strResearchComments, "W3057CRE") > 0 Then
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "W870") > 0 Or InStr(strResearchComments, "W1444") > 0 Or InStr(strResearchComments, "W1445") > 0 Then
                    Return strResearchComments
                End If
                For L As Integer = LBound(CurrentPaymentScreenForLog) To UBound(CurrentPaymentScreenForLog)
                    If CurrentPaymentScreenForLog(L).EditWarning Is Nothing Then Continue For
                    If InStr(CurrentPaymentScreenForLog(L).EditWarning, "W1107COPAY NBR AMT REQ") > 0 Or InStr(EditWarning, "WE2493COPAY NBR AMT REQ") > 0 Then
                        isNCopayNBR = True
                        WriteLog("EditWarning =" + CurrentPaymentScreenForLog(L).EditWarning)
                        WriteLog("isNCopayNBR = True")
                    End If
                Next
                If InStr(strResearchComments, "Unable to remove PZ1") > 0 Then
                    Return strResearchComments
                End If
                'Dim Text As String = GetPaymentScreen(CurrentPaymentScreenForLog)
                'WriteLog(Text)
                'WriteLog("Payment Screen Read Complete")
                CheckBypassHistoryField()
                ReadClaimPaymentScreen(CurrentClaimPaymentScreen, CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount, False, icn, MaxSuffix, True)
                If strResearchComments.Contains("ENTER UNITS ON REPEATED LINE") Or strResearchComments.Contains("Cannot change pricing until 3C Is split") Then
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "Exception") > 0 Then
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "W870") > 0 Or InStr(strResearchComments, "W1444") > 0 Or InStr(strResearchComments, "W1445") > 0 Then
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "Unable to remove PZ1") > 0 Then
                    Return strResearchComments
                End If
                If strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE" Then
                    Return strResearchComments
                End If
                If strResearchComments = "unable to remove pisl line" Then
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "WE3066MARS EDITING APPLIES") > 0 Then
                    'strResearchComments = ""
                    MARSEDITINGAPPLIES()
                    Return strResearchComments
                End If
                If InStr(strResearchComments, "RSC APLY JK") > 0 Then      'Added by Sourav 22/05/24 so that mod50 claims don't get skipped
                    strResearchComments = ""
                End If
                If InStr(strResearchComments, "WE1519ILLOGICAL RC") > 0 Then
                    strResearchComments = ""
                End If
                If InStr(strResearchComments, "WE124INVALID CODE") > 0 Then     'Added by Sourav @27/02/2025: To not skip claims with this Hard edit
                    If GetText(2, 2, 2) = "MP" Then
                        SendKeyes("EDS", 2, 2)
                        SendEnter()
                        apiChk()
                        apiChk()
                    End If
                    RemoveNotCoveredFromEDS6Point5()
                    If GetText(3, 2, 5) = "EDS 6" Then
                        SendKeyesFor01("A", 1, 75)
                        SendEnter()
                    End If
                    strResearchComments = ""
                End If
                If InStr(strResearchComments, "HARD EDIT - E# or WE Found") > 0 Then
                    Return strResearchComments
                End If
                If strResearchComments = "Check manual spi exists." Then
                    Return strResearchComments
                End If
                'EDSScreenLineCount(CurrentEDS6Point5Screen, MaxSuffix)
                If strResearchComments = "eds not valid" Then
                    Return strResearchComments
                End If
                'CheckisEDSLineMatch = isLineMatchInEDS1()
                CheckisEDSLineMatch = True
                SendEnter()

                If CheckisEDSLineMatch = True Then
                    Do
                        If EditWarning = String.Empty Then
                            EditWarning = String.Empty
                            Count = 0
                        End If
                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        If Count > 50 Then Exit Do
                        EditWarning += GetText(24, 1, 80, True)
                        SendEnter()
                    Loop
                    If InStr(EditWarning, "WE") > 0 OrElse InStr(EditWarning, ",E") > 0 Then

                        If InStr(EditWarning, ",WE") > 0 Then
                            'Dim str() As String = New String() {}
                            Dim startPosition As Integer = EditWarning.IndexOf(",WE") + 1
                            Dim temp As String = EditWarning.Substring(startPosition)
                            Dim endPosition As Integer = temp.IndexOf(",")

                            Dim edit As String = temp.Substring(0, endPosition)
                            If InStr(edit, "WE3066MARS EDITING APPLIES") > 0 Then
                                'strResearchComments = ""
                                MARSEDITINGAPPLIES()
                                Return strResearchComments
                            End If
                            If InStr(edit, "RSC APLY JK") > 0 Then      'Added by Sourav 22/05/24 so that mod50 claims don.t get skipped
                                strResearchComments = ""
                                Return ""
                            End If
                            If InStr(edit, "WE1519ILLOGICAL RC") > 0 Then
                                strResearchComments = ""
                                Return ""
                            End If
                            If InStr(edit, "WE124INVALID CODE") > 0 Then     'Added by Sourav @27/02/2025: To not skip claims with this Hard edit
                                If GetText(2, 2, 2) = "MP" Then
                                    SendKeyes("EDS", 2, 2)
                                    SendEnter()
                                    apiChk()
                                    apiChk()
                                End If
                                RemoveNotCoveredFromEDS6Point5()
                                If GetText(3, 2, 5) = "EDS 6" Then
                                    SendKeyesFor01("A", 1, 75)
                                    SendEnter()
                                End If
                                strResearchComments = ""
                            End If
                            strResearchComments = "HARD EDIT - E# or WE Found: " + edit

                            Return strResearchComments
                        End If
                        If InStr(EditWarning, ",E") > 0 Then
                            If InStr(EditWarning, "RSC APLY JK") > 0 Then      'Added by Sourav 22/05/24 so that mod50 claims don.t get skipped
                                strResearchComments = ""
                                Return ""
                            End If
                            Dim startPosition As Integer = EditWarning.IndexOf(",E") + 1
                            Dim temp As String = EditWarning.Substring(startPosition)
                            Dim endPosition As Integer = temp.IndexOf(",")

                            Dim edit As String = temp.Substring(0, endPosition)
                            If InStr(edit, "3066MARS EDITING APPLIES") > 0 Then
                                'strResearchComments = ""
                                MARSEDITINGAPPLIES()
                                Return strResearchComments
                            End If
                            strResearchComments = "HARD EDIT - E# or WE Found: " + edit
                            Return strResearchComments
                        End If
                    End If
                    EditWarningCheck(EditWarning)
                    If InStr(strResearchComments, "W870") > 0 Or InStr(strResearchComments, "W1444") > 0 Or InStr(strResearchComments, "W1445") > 0 Then
                        Return strResearchComments
                    End If
                    If InStr(EditWarning, "W1107COPAY NBR AMT REQ") > 0 Or InStr(EditWarning, "WE2493COPAY NBR AMT REQ") > 0 Then
                        isNCopayNBR = True
                    End If
                Else
                    strResearchComments = "EDS 1 and EDS 6.5 Line are not match"
                    Return strResearchComments
                End If
                'PasteHistoryICN(icn)
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " =>" + ex.Message.ToString)
            strResearchComments = "EXCEPTION=>" = ex.Message.ToString
        End Try
        WriteLog("RESEARCH COMMENTS= " + strResearchComments)
        WriteLog("=====================EXIT " + GetCurrentMethod.Name + "=======================")
        Return strResearchComments
    End Function
    Public Function AllowableCalulationFromMPC() As Double
        Dim Alowable As String
        Dim AlowableFinal As Double
        Dim sumAlowable As Double
        Alowable = String.Empty
        AlowableFinal = 0.0
        sumAlowable = 0.0
        If Trim(GetText(2, 2, 2)) = "MP" Then
            For jj As Integer = 11 To 23 Step 2
                If Trim(GetText(jj, 30, 5)) = "-----" Then Exit For
                Alowable = Trim(GetText(jj, 30, 5)) + "." + Trim(GetText(jj, 36, 2))
                If Alowable = ".00" Then
                    Alowable = "0.00"
                End If
                AlowableFinal = CDbl(Alowable)
                sumAlowable = sumAlowable + AlowableFinal

            Next
        End If

        Return sumAlowable
    End Function

    Public Function WorkOnMHI(ByVal icn As String) As Double
        Dim Alowable As String
        Dim FinalAlowale As Double = 0.00
        Try
            SendClear()
            SendClear()
            SendKeyes("ret,i" + icn + ",m", 2, 2)
            SendEnter()
            If Trim(GetText(2, 2, 3)) = "MHN" Then
                For dd As Integer = 10 To 18 Step 4
                    If Trim(GetText(dd + 1, 1, 1)) = "%" Then
                        If Trim(GetText(dd, 36, 2)) <> "74" AndAlso Trim(GetText(dd, 36, 2)) <> "69" Then

                            Alowable = Trim(GetText(dd + 1, 31, 4)) + "." + Trim(GetText(dd + 1, 36, 2))
                            If Alowable = "." OrElse Alowable = ".00" Then
                                Alowable = "0.00"
                            End If
                            FinalAlowale = CDbl(Alowable) + FinalAlowale
                        End If
                    End If

                    If dd = 18 Then
                        SendF8()
                        If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        SendF8()
                        dd = 6
                    End If

                Next
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return FinalAlowale
    End Function

    Public Function IHAndOHSegrigation() As String
        Dim CheckRevcode As String
        CheckRevcode = String.Empty
        Dim result As String
        result = String.Empty
        Dim IH As Boolean = False
        Dim OH As Boolean = False
        Try
            For dd As Integer = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                For ff As Integer = LBound(CurrentICNEDS1(dd).ServiceLine.RevCode) To UBound(CurrentICNEDS1(dd).ServiceLine.RevCode)
                    CheckRevcode = CurrentICNEDS1(dd).ServiceLine.RevCode(ff)
                    If CInt(CheckRevcode) >= 100 AndAlso CInt(CheckRevcode) <= 219 Then
                        IH = True
                    ElseIf CInt(CheckRevcode) >= 100 AndAlso CInt(CheckRevcode) <= 219 AndAlso CInt(CheckRevcode) >= 219 Then
                        IH = True
                    ElseIf CInt(CheckRevcode) > 219 Then
                        OH = True
                    End If
                Next
            Next
            If IH = True Then
                result = "IH"
            Else
                result = "OH"
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return result
    End Function
    Public Sub PullAHI()
        Try
            SendClear()
            SendClear()
            Dim Policy As String = String.Empty
            Dim SSN As String = String.Empty
            Dim Pat As String = String.Empty
            Dim PatRL As String = String.Empty
            Dim Suffix_Tin_Prefix = String.Empty
            Dim Fdate As String = String.Empty
            Dim Ldate As String = String.Empty
            Dim Prefix_Tin As String = String.Empty
            Policy = CurrentClaimDetails.Policy
            SSN = CurrentClaimDetails.SSN
            Pat = CurrentClaimDetails.PatientName
            PatRL = CurrentClaimDetails.PatientRelationship
            Suffix_Tin_Prefix = CurrentClaimDetails.Prifix_Tin_Suffix
            Fdate = CurrentClaimDetails.Fdate
            Ldate = CurrentClaimDetails.Ldate
            Prefix_Tin = CurrentClaimDetails.Prefix_Tin
            'SendKeyes("AHI" + "," + Policy + "," + SSN + "," + Pat + "," + PatRL + ",T" + Suffix_Tin_Prefix + ",," + Fdate + "," + Ldate, 2, 2)
            'SendEnter()
            SendKeyes("AHI" + "," + Policy + "," + SSN + "," + Pat + "," + PatRL + ",T" + Prefix_Tin + ",," + Fdate + "," + Ldate, 2, 2)
            SendEnter()
            apiChk()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
            WriteLog(GetTextRect)
        End Try

    End Sub
    Public Function FindPossibleICN() As String
        Dim OrICN As String = CurrentClaimDetails.ICN
        Dim icn As String
        Dim AHIICN As String
        icn = String.Empty
        AHIICN = String.Empty
        Dim HCFC As Boolean
        Dim UB As Boolean
        Dim isMHI As Boolean = True
        Dim isSkip As Boolean           'Added by Sourav to skip the entire claim if HH or IJ is present in HxICN:(19.05.2023)
        Try
            PullAHI()
            If InStr(Trim(GetText(24, 1, 80)), "NO CLAIMS ON FILE") > 0 Then Return "NO CLAIMS ON FILE"
            WriteLog(GetTextRect)
            For dd As Integer = 4 To 20 Step 2
                If Trim(GetText(dd, 4, 6)) = "------" Then Exit For

                If Trim(GetText(dd + 1, 54, 10)) = OrICN Then
                    Continue For
                Else
                    If GetText(dd + 1, 52, 1) = "F" Then
                        Continue For
                        '--------commented

                        'SendKeyes("S", dd, 2)
                        'SendEnter()
                        'apiChk()
                        'If InStr(Trim(GetText(2, 2, 3)), "GHI") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHN") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHB") > 0 Then
                        '    isMHI = False
                        '    SendKeyes("MHI", 2, 2)
                        '    autECLPSObj.SendKeys("[ENTER]")
                        '    apiChk()
                        'End If

                        'If InStr(Trim(GetText(10, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(10, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "MISC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "OPS") > 0 Then
                        '    UB = True
                        '    If isMHI = False Then
                        '        PullAHI()
                        '        isMHI = True
                        '    Else
                        '        SendKeyes("A", 3, 15)
                        '        SendEnter()
                        '        apiChk()
                        '    End If
                        'Else
                        '    HCFC = True
                        'End If
                        'AHIICN = Trim(GetText(dd + 1, 54, 10))

                        '----------commented

                        'If UB = True AndAlso Trim(GetText(dd + 1, 54, 10)) <> OrICN Then

                        'If InStr(icn, AHIICN) = 0 Then
                        '    icn = AHIICN + "," + icn
                        'End If
                        'End If

                    ElseIf GetText(dd + 1, 52, 1) = "I" Then
                        UB = False
                        SendKeyes("S", dd, 2)
                        SendEnter()
                        apiChk()
                        If InStr(Trim(GetText(2, 2, 3)), "GHI") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHN") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHB") > 0 Then
                            isMHI = False
                            SendKeyes("MHI", 2, 2)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        If Trim(GetText(2, 2, 2)) = "MH" Then
                            isMHI = True
                        End If
                        If InStr(Trim(GetText(10, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(10, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "MISC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "OPS") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "OA") > 0 Then
                            UB = True
                            If isMHI = False Then
                                PullAHI()
                                isMHI = True
                            Else
                                For ff As Integer = 12 To 20 Step 4
                                    If GetText(ff - 2, 36, 2) = "HH" Or GetText(ff - 2, 36, 2) = "IJ" Then
                                        strResearchComments = "HH or IJ present in HxICN"
                                        isSkip = True
                                        Return ""
                                    End If
                                    If GetText(ff - 2, 36, 2) = "S6" Then
                                        strResearchComments = "Remark Code S6 present in HxICN"
                                        isSkip = True
                                        Return ""
                                    End If
                                    If GetText(ff - 2, 36, 2) = "UP" Or GetText(ff - 2, 36, 2) = "LP" Or GetText(ff - 2, 36, 2) = "EW" Then
                                        strResearchComments = "Remark Code UP/LP/EW present in HxICN"
                                        isSkip = True
                                        Return ""
                                    End If
                                    If InStr(Trim(GetText(ff, 1, 3)), "ICN") > 0 And Trim(GetText(ff, 5, 10)) <> "0000000000" Then
                                        AHIICN = Trim(GetText(ff, 5, 10))
                                    End If

                                    If ff = 20 Then
                                        'SendF8()
                                        If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                                        SendF8()
                                        ff = 8
                                    End If
                                Next
                                SendKeyes("A", 3, 15)
                                SendEnter()
                                apiChk()
                                'IncreementMPCCount()

                                If Trim(GetText(2, 2, 2)) <> "AH" Then
                                    PullAHI()
                                End If
                            End If
                        ElseIf InStr(Trim(GetText(10, 4, 6)), "OI") > 0 Then
                            strResearchComments = "OI present in HxICN"
                            isSkip = True
                            Return ""
                        Else
                            'HCFC = True
                        End If

                    End If


                End If
                If UB = True AndAlso Trim(GetText(dd + 1, 54, 10)) <> OrICN Then
                    'AHIICN = Trim(GetText(dd + 1, 54, 10))
                    If InStr(icn, AHIICN) = 0 Then
                        icn = AHIICN + "," + icn
                    End If
                Else
                    SendKeyes("A", 3, 15)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    'IncreementMPCCount()
                    If Trim(GetText(2, 2, 2)) <> "AH" Then
                        PullAHI()
                    End If
                End If
                If dd = 20 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    SendF8()
                    dd = 2
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try

        Return icn
    End Function

    Public Function FindMatchICNList(ByRef UserCtrlFrm As frmMain) As String
        Dim FinalICNList1 As String = String.Empty
        Dim FinalICNList As String = String.Empty
        Dim isHardCopyICNUB As Boolean = False
        'Dim CurrICN = CurrentClaimDetails.ICN
        Try
            WriteLog("==================ENTERING " + GetCurrentMethod.Name + "========================")
            SendClear()
            SendClear()
            Dim checkCurrClaimOHandOPS As Boolean = CheckCurrentClaimforOHandOPS()
            Dim HxICNList As String = FindPossibleICN()
            WriteLog($"Possible ICNs found from AHI: {HxICNList}")
            If InStr(strResearchComments, "HH or IJ present in HxICN") > 0 Then Return "HH or IJ present in HxICN"
            If InStr(strResearchComments, "Remark Code S6 present in HxICN") > 0 Then Return "Remark Code S6 present in HxICN"
            If InStr(strResearchComments, "Remark Code UP/LP/EW present in HxICN") > 0 Then Return "Remark Code UP/LP/EW present in HxICN"
            If InStr(strResearchComments, "OI present in HxICN") > 0 Then Return "OI present in HxICN"
            If InStr(HxICNList, "NO CLAIMS ON FILE") > 0 Then Return "AHI: W062NO CLAIMS ON FILE"
            Dim tempList As String = HxICNList
            'Dim ICN() As String = ICNList.Split(",")
            Dim OrigCLMRefListofHxICN As String = openHardCopy(HxICNList, UserCtrlFrm)
            If InStr(OrigCLMRefListofHxICN, "Exception") > 0 Then Return strResearchComments = OrigCLMRefListofHxICN
            Dim OrigCLMRefListArr() As String = OrigCLMRefListofHxICN.Split(",")
            'Dim flagHardCopy = WorkOnHardCopyClaim(OrigCLMRefListArr)
            'If flagHardCopy = True Then
            '    For i As Integer = 0 To OrigCLMRefListArr.Length - 1
            '        If InStr(FinalICNList, OrigCLMRefListArr(i)) = 0 Then
            '            If FinalICNList = "" Then
            '                FinalICNList = OrigCLMRefListArr(i)
            '            Else
            '                FinalICNList = FinalICNList + "," + OrigCLMRefListArr(i)
            '            End If
            '        End If
            '    Next
            'Else
            '   For i As Integer = 0 To OrigCLMRefListArr.Length - 1
            '        If InStr(HxICNList, OrigCLMRefListArr(i)) = 0 Then
            '            tempList = tempList + "," + OrigCLMRefListArr(i)
            '        End If
            '    Next
            'End If
            For i As Integer = 0 To OrigCLMRefListArr.Length - 1
                If InStr(tempList, OrigCLMRefListArr(i)) = 0 AndAlso OrigCLMRefListArr(i) <> "" Then
                    tempList = tempList + "," + OrigCLMRefListArr(i)
                End If
            Next
            WriteLog($"Temp Hx ICN list for criteria matching:{tempList}")
            FinalICNList1 = GetPercentageMatchfromEDS1orRHI(tempList)
            If InStr(strResearchComments, "Less than 50%") > 0 Then Return ""
            If checkCurrClaimOHandOPS Then
                Dim OHandOPSIcnList As String = OHandOPSinHxICN(tempList)
                Dim OHOPSIcnListArr() As String = OHandOPSIcnList.Split(",")
                For i As Integer = 0 To OHOPSIcnListArr.Length - 1
                    If InStr(FinalICNList1, OHOPSIcnListArr(i)) = 0 Then
                        If FinalICNList1 = "" Then
                            FinalICNList1 = OHOPSIcnListArr(i)
                        Else
                            FinalICNList1 = FinalICNList1 + "," + OHOPSIcnListArr(i)
                        End If
                    End If
                Next

            End If
            Dim HardCopyICNofCurrClaim As String = Left(openHardCopy(CurrentClaimDetails.ICN + ",", UserCtrlFrm), 10)
            Dim IsvalidHardCopyICN As Boolean = CheckIfValidICN(Left(HardCopyICNofCurrClaim, 10))
            If IsvalidHardCopyICN Then
                isHardCopyICNUB = pullHxRHI(Left(HardCopyICNofCurrClaim, 10))
            End If
            If IsvalidHardCopyICN = False OrElse isHardCopyICNUB = False Then
                If FinalICNList.Trim.Contains(FinalICNList1) = False And InStr(FinalICNList, FinalICNList1) = 0 Then
                    FinalICNList = FinalICNList + "," + FinalICNList1
                End If
            Else
                If InStr(FinalICNList1, HardCopyICNofCurrClaim) = 0 Then
                    FinalICNList = FinalICNList1 + HardCopyICNofCurrClaim
                Else
                    FinalICNList = FinalICNList1
                End If
            End If
            WriteLog("FINAL LIST OF HISTORY ICNs: " + FinalICNList)
            WriteLog("==================EXITING" + GetCurrentMethod.Name + "========================")
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception" + ex.Message)

        End Try
        Return FinalICNList
    End Function

    Public Sub GetFLNDCCforHistICN(ByRef FLN As String, ByRef DCC As String, icn As String)
        Try
            If icn = CurrentClaimDetails.ICN Then
                FLN = CurrentClaimDetails.FLN
                DCC = CurrentClaimDetails.DCC
            Else
                PullMHI(icn)
                For dd As Integer = 12 To 20 Step 4

                    If InStr(Trim(GetText(dd, 1, 3)), "ICN") > 0 Then
                        FLN = GetText(dd, 34, 10)
                        DCC = GetText(dd, 49, 3)
                        Exit Sub
                    End If

                    If dd = 20 Then
                        'SendF8()
                        If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        SendF8()
                        dd = 8

                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub

    Public Function openHardCopy(HistICNList As String, ByRef UserCtrlFrm As frmMain) As String
        Dim tempList As String = ""
        'Dim HistICNList2 As String
        'If InStr(HistICNList, CurrentClaimDetails.ICN) = 0 Then
        '    HistICNList2 = CurrentClaimDetails.ICN + "," + HistICNList
        'End If
        Dim ICN() As String = HistICNList.Split(",")
        Dim HICN As String = String.Empty
        Dim PICN As String
        Dim FindIcn As String
        Dim LineCount As Integer = 0
        Dim strMainLine As String
        Dim LineCount1 As Integer = 0
        Dim arrMainLine() As String
        Dim CheckHardCopy As String = String.Empty
        Try
            For ff As Integer = LBound(ICN) To UBound(ICN) - 1
                PICN = ICN(ff)
                Dim FLN As String = String.Empty
                Dim DCC As String = String.Empty
                GetFLNDCCforHistICN(FLN, DCC, PICN)
                Dim getDoc360 As String = DOC360(FLN, DCC)
                Dim ICNDraft As String = String.Empty
                Dim ICNDraft1 As String = String.Empty
                Dim Draft As String = String.Empty
                If getDoc360 = "NA" Then
                    CheckHardCopy = "Hard copy is not open"
                Else
                    UserCtrlFrm.rtfEdss.Text = getDoc360
                    For Each strLine As String In UserCtrlFrm.rtfEdss.Lines
                        If InStr(strLine, "TYPE BILL:") > 0 Then
                            LineCount = LineCount + 1
                            strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                            While InStr(strMainLine, "|PAT STATUS CD:") = 0
                                strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                                If InStr(strMainLine, "ORIG CLM REF NUMBER:") > 0 Then
                                    Erase arrMainLine
                                    arrMainLine = strMainLine.Split(":")
                                    ICNDraft1 = arrMainLine(1).Replace("|", "")
                                    ICNDraft = Trim(ICNDraft1)
                                    HICN = Left(ICNDraft, 10)
                                    If IsAlphaNumeric(HICN) = False Then
                                        ICNDraft = PICN
                                        HICN = PICN
                                    End If
                                    If ICNDraft.Length > 11 Then
                                        Draft = Right(ICNDraft, 10)
                                        'Draft = ICNDraft.Substring(11, 10)
                                    End If

                                    If Left(HICN, 1) = "0" And ICNDraft.Length < 11 Then
                                        Draft = HICN
                                        HICN = ""
                                    Else
                                        HICN = HICN
                                    End If
                                    FindIcn = HICN
                                    If Trim(HICN) <> "" AndAlso InStr(tempList, HICN) = 0 AndAlso InStr(HistICNList, HICN) = 0 Then
                                        tempList = HICN + "," + tempList
                                        Exit For
                                    Else
                                        CheckHardCopy = "Hard copy ICN is not found"
                                    End If

                                End If
                                LineCount = LineCount + 1
                            End While

                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
            tempList = $"Exception in {GetCurrentMethod.Name}"
        End Try

        Return tempList
    End Function

    Public Function GetPercentageMatchfromEDS1orRHI(tempList As String) As String
        Dim FinalList As String = String.Empty
        Dim tempListArr() As String = tempList.Split(",")
        Dim FindIcn As String
        Dim PICN As String = String.Empty
        Dim eliminatedICNList As String = String.Empty
        Dim checkIfHxIcnIsIHClaim As Boolean = False
        Try
            For ff As Integer = LBound(tempListArr) To UBound(tempListArr) - 1
                If tempListArr(ff) <> "" Then
                    Dim EDS1RevCode As String = String.Empty
                    Dim EDS1HCPC As String = String.Empty
                    Dim EDS1M1 As String = String.Empty
                    Dim EDS1M2 As String = String.Empty
                    Dim EDS1Date As String = String.Empty
                    Dim EDS1Charge As String = String.Empty
                    Dim EDS1Unit As String = String.Empty
                    Dim HEDS1RevCode As String = String.Empty
                    Dim HEDS1HCPC As String = String.Empty
                    Dim HEDS1M1 As String = String.Empty
                    Dim HEDS1M2 As String = String.Empty
                    Dim HEDS1Date As String = String.Empty
                    Dim HEDS1Charge As String = String.Empty
                    Dim HEDS1Unit As String = String.Empty
                    FindIcn = String.Empty
                    PICN = tempListArr(ff)
                    SendClear()
                    SendClear()
                    SendKeyes("mei,i" + PICN + ",1", 2, 2)
                    SendEnter()
                    apiChk()
                    apiChk()
                    If InStr(Trim(GetText(24, 1, 80)), "E1569INVALID ICN") Then
                        Continue For
                    End If
                    If InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Or Trim(GetText(7, 2, 4)) <> "REVE" Then
                        Dim DraftComments As String = ""
                        GetHRHI(PICN, HistoryRHI, DraftComments)
                        If DraftComments = "Draft list is empty" Then
                            Continue For
                        End If
                        If InStr(strResearchComments, "W062NO CLAIMS ON FILE") > 0 Then
                            Continue For
                        End If
                        For rr As Integer = LBound(HistoryRHI) To UBound(HistoryRHI)

                            If InStr("IC,NB,SP,PR,WD", HistoryRHI(rr).RHIServiceLineCount.RHITranslation) > 0 Then
                                checkIfHxIcnIsIHClaim = True
                                WriteLog($"Hx ICN {PICN} is an IH claim. {HistoryRHI(rr).RHIServiceLineCount.RHITranslation} Proc code found")
                                Exit For
                            End If
                            If checkIfHxIcnIsIHClaim = True Then Exit For

                        Next

                        If (isCurrentClaimOH And checkIfHxIcnIsIHClaim) Or (isCurrentClaimIH And (Not checkIfHxIcnIsIHClaim)) Then
                            dt_ListOfHxMatchingANPI = New DataTable()
                            GenerateDt()
                            Dim FlagNotMatch As Boolean = False
                            MatchAttendingPhysician(PICN, FlagNotMatch)                     'Added 9/13/2025
                            For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                If Row("ANPIMatchFlag") = True Then
                                    If InStr(FinalList, PICN) = 0 Then
                                        FinalList = PICN + "," + FinalList
                                    End If
                                Else
                                    WriteLog($"Manual intv. reqd: Attending physician of History icn {PICN} not same for IH/OH case")
                                    strResearchComments = "Manual intv. reqd: Attending physician of History not same for IH/OH case"
                                    Return ""
                                End If
                            Next
                        Else

                            Dim kk As Integer = 0
                            Dim ll As Integer = 0
                            Dim count1 As Integer = 0
                            Dim TotalCount1 As Integer = 0
                            Dim Divresult1 As Decimal = 0.00
                            Dim PercentageResult1 As Integer = 0
                            For rr As Integer = LBound(HistoryRHI) To UBound(HistoryRHI)
                                For uu As Integer = LBound(HistoryRHI(rr).RHIServiceLineCount.RevCode) To UBound(HistoryRHI(rr).RHIServiceLineCount.RevCode)
                                    TotalCount1 += 1
                                    HEDS1RevCode = HistoryRHI(rr).RHIServiceLineCount.RevCode(uu)
                                    HEDS1Charge = HistoryRHI(rr).RHIServiceLineCount.Charge(uu)
                                    HEDS1Date = HistoryRHI(rr).RHIServiceLineCount.DateOfService(uu)
                                    HEDS1HCPC = HistoryRHI(rr).RHIServiceLineCount.ProcCode(uu)
                                    HEDS1Unit = HistoryRHI(rr).RHIServiceLineCount.Unit(uu)
                                    For kk = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                        For ll = LBound(CurrentICNEDS1(kk).ServiceLine.RevCode) To UBound(CurrentICNEDS1(kk).ServiceLine.RevCode)
                                            EDS1RevCode = CurrentICNEDS1(kk).ServiceLine.RevCode(ll)
                                            EDS1Charge = CurrentICNEDS1(kk).ServiceLine.BillCharge(ll)
                                            EDS1HCPC = CurrentICNEDS1(kk).ServiceLine.CPTCode(ll)
                                            EDS1Unit = "0" + CurrentICNEDS1(kk).ServiceLine.Unit(ll)
                                            If InStr(EDS1HCPC, "-") > 0 Then
                                                EDS1HCPC = Trim(EDS1HCPC.Replace("-", ""))
                                            End If
                                            EDS1Date = CurrentICNEDS1(kk).ServiceLine.DateofService(ll)
                                            If isCurrentClaimIH Then
                                                If HEDS1Unit = EDS1Unit And HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC Then   '//date matching criteria removed for IH claim on 26th July,2026
                                                    count1 += 1
                                                End If
                                            Else
                                                If HEDS1Unit = EDS1Unit And HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC AndAlso HEDS1Date = EDS1Date Then
                                                    count1 += 1
                                                End If

                                            End If
                                        Next
                                    Next
                                Next
                                kk = 0
                                ll = 0
                            Next
                            Divresult1 = count1 / TotalCount1
                            PercentageResult1 = Divresult1 * 100
                            WriteLog("%age match for HX ICN " + PICN + "=" + PercentageResult1.ToString)
                            If PercentageResult1 >= 40 Then     'changed to 40 on 02/12/2025;   MODIFIED to 
                                'If InStr(FinalList, PICN) = 0 Then
                                '    FinalList = PICN + "," + FinalList
                                'End If
                                WriteLog("Comparison %age =" + PercentageResult1.ToString)
                                dt_ListOfHxMatchingANPI = New DataTable()
                                GenerateDt()
                                Dim FlagNotMatch As Boolean = False
                                MatchAttendingPhysician(PICN, FlagNotMatch)                 'Added 9/13/2025
                                For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                    If Row("ANPIMatchFlag") = True Then
                                        If InStr(FinalList, PICN) = 0 Then
                                            FinalList = PICN + "," + FinalList
                                        End If
                                    Else
                                        strResearchComments = "Manual intv. reqd: Less than 50% & ANPI not match "
                                        Return ""
                                    End If
                                Next
                                '=========================================================================

                            End If
                        End If

                    Else
                        If Trim(GetText(6, 2, 5)) = "EDS 1" Then
                            WriteLog(GetTextRect())
                            GetEDS1(HistoryICNEDS1, PICN)
                            Dim hh As Integer = 0
                            Dim nn As Integer = 0
                            Dim count As Integer = 0
                            Dim TotalCount As Integer = 0
                            Dim PercentageResult As Integer = 0
                            Dim DivResult As Decimal = 0

                            For bb As Integer = LBound(HistoryICNEDS1) To UBound(HistoryICNEDS1)    'To find whether history icn is IH claim
                                For ll As Integer = LBound(HistoryICNEDS1(bb).ServiceLine.RevCode) To UBound(HistoryICNEDS1(bb).ServiceLine.RevCode)
                                    If InStr("IC,NB,SP,PR,WD", HistoryICNEDS1(bb).ServiceLine.ServiceCode(ll)) > 0 Then
                                        checkIfHxIcnIsIHClaim = True
                                        Exit For
                                    End If
                                    If checkIfHxIcnIsIHClaim = True Then Exit For
                                Next
                                If checkIfHxIcnIsIHClaim = True Then Exit For
                            Next

                            'if hx claim is IH and current is OH claim or vice-versa, then only Attending physician will be used to find match HX Icn   @13/12/2025
                            If (isCurrentClaimOH And checkIfHxIcnIsIHClaim) Or (isCurrentClaimIH And (Not checkIfHxIcnIsIHClaim)) Then
                                dt_ListOfHxMatchingANPI = New DataTable()
                                GenerateDt()
                                Dim FlagNotMatch As Boolean = False
                                MatchAttendingPhysician(PICN, FlagNotMatch)                     'Added 9/13/2025
                                For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                    If Row("ANPIMatchFlag") = True Then
                                        If InStr(FinalList, PICN) = 0 Then
                                            FinalList = PICN + "," + FinalList
                                        End If
                                    Else
                                        WriteLog($"Manual intv. reqd: Attending physician of History icn {PICN} not same for IH/OH case")
                                        strResearchComments = "Manual intv. reqd: Attending physician of History not same for IH/OH case"
                                        Return ""
                                    End If
                                Next
                            Else    'FOR every other cases the initial %age matching process will be followed   @13/12/2025

                                For bb As Integer = LBound(HistoryICNEDS1) To UBound(HistoryICNEDS1)
                                    For ll As Integer = LBound(HistoryICNEDS1(bb).ServiceLine.RevCode) To UBound(HistoryICNEDS1(bb).ServiceLine.RevCode)
                                        TotalCount += 1
                                        HEDS1RevCode = HistoryICNEDS1(bb).ServiceLine.RevCode(ll)
                                        HEDS1Charge = HistoryICNEDS1(bb).ServiceLine.BillCharge(ll)
                                        HEDS1Date = HistoryICNEDS1(bb).ServiceLine.DateofService(ll)
                                        HEDS1HCPC = HistoryICNEDS1(bb).ServiceLine.CPTCode(ll)
                                        HEDS1Unit = HistoryICNEDS1(bb).ServiceLine.Unit(ll)
                                        For hh = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                            For nn = LBound(CurrentICNEDS1(hh).ServiceLine.RevCode) To UBound(CurrentICNEDS1(hh).ServiceLine.RevCode)
                                                EDS1RevCode = CurrentICNEDS1(hh).ServiceLine.RevCode(nn)
                                                EDS1Charge = CurrentICNEDS1(hh).ServiceLine.BillCharge(nn)
                                                EDS1HCPC = CurrentICNEDS1(hh).ServiceLine.CPTCode(nn)
                                                EDS1Date = CurrentICNEDS1(hh).ServiceLine.DateofService(nn)
                                                EDS1Unit = CurrentICNEDS1(hh).ServiceLine.Unit(nn)
                                                If checkIfHxIcnIsIHClaim = True Then
                                                    If HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge Then
                                                        count += 1
                                                    End If
                                                Else
                                                    If HEDS1Unit = EDS1Unit And HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC AndAlso HEDS1Date = EDS1Date Then
                                                        count += 1
                                                    End If
                                                End If

                                            Next
                                        Next
                                    Next
                                    hh = 0
                                    nn = 0
                                Next
                                DivResult = count / TotalCount
                                PercentageResult = DivResult * 100

                                If PercentageResult >= 40 Then      'changed to 40 on 02/12/2025
                                    FindIcn = PICN
                                    'If InStr(FinalList, PICN) = 0 Then
                                    '    FinalList = PICN + "," + FinalList
                                    'End If
                                    WriteLog("Comparison %age =" + PercentageResult.ToString)
                                    'THIS part added to skip the entire claim if %age matching is less than 50% or anpi doesn't match ON 23/11/2025
                                    dt_ListOfHxMatchingANPI = New DataTable()
                                    GenerateDt()
                                    Dim FlagNotMatch As Boolean = False
                                    MatchAttendingPhysician(PICN, FlagNotMatch)                     'Added 9/13/2025
                                    For Each Row As DataRow In dt_ListOfHxMatchingANPI.Rows
                                        If Row("ANPIMatchFlag") = True Then
                                            If InStr(FinalList, PICN) = 0 Then
                                                FinalList = PICN + "," + FinalList
                                            End If
                                        Else
                                            strResearchComments = "Manual intv. reqd: Less than 50% & ANPI not match "
                                            Return ""
                                        End If
                                    Next

                                End If
                            End If


                        End If
                    End If
                End If
            Next
            Dim listB As String = AlternatePercentageMatching(eliminatedICNList)
            FinalList = FinalList + listB
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name() + " exception: " + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name() + " exception: " + ex.Message.ToString
        End Try

        Return FinalList
    End Function

    Public Function FindMatchICNList2(ByRef UserCtrlFrm As frmMain) As String
        SendClear()
        SendClear()
        Dim ICNList As String = FindPossibleICN()
        Dim ICN() As String = ICNList.Split(",")
        Dim PICN As String
        Dim FinalICN As String
        Dim FindIcn As String
        Dim CheckHardCopy As String = String.Empty
        PICN = String.Empty
        FinalICN = String.Empty
        Dim checkCount As Integer = 0
        Try
            For ff As Integer = LBound(ICN) To UBound(ICN) - 1
                Dim EDS1RevCode As String = String.Empty
                Dim EDS1HCPC As String = String.Empty
                Dim EDS1M1 As String = String.Empty
                Dim EDS1M2 As String = String.Empty
                Dim EDS1Date As String = String.Empty
                Dim EDS1Charge As String = String.Empty
                Dim HEDS1RevCode As String = String.Empty
                Dim HEDS1HCPC As String = String.Empty
                Dim HEDS1M1 As String = String.Empty
                Dim HEDS1M2 As String = String.Empty
                Dim HEDS1Date As String = String.Empty
                Dim HEDS1Charge As String = String.Empty
                FindIcn = String.Empty
                PICN = ICN(ff)
                SendClear()
                SendClear()
                SendKeyes("mei,i" + PICN + ",1", 2, 2)
                SendEnter()
                apiChk()
                If InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Then
                    If checkCount = 0 Then
                        Dim HICN As String = String.Empty
                        Dim LineCount As Integer = 0
                        Dim strMainLine As String
                        Dim LineCount1 As Integer = 0
                        Dim arrMainLine() As String
                        Dim FLN As String = CurrentClaimDetails.FLN
                        Dim DCC As String = CurrentClaimDetails.DCC
                        Dim getDoc360 As String = DOC360(FLN, DCC)
                        Dim ICNDraft As String = String.Empty
                        Dim ICNDraft1 As String = String.Empty
                        Dim Draft As String = String.Empty
                        If getDoc360 = "NA" Then
                            CheckHardCopy = "Hard copy is not open"
                        Else
                            UserCtrlFrm.rtfEdss.Text = getDoc360
                            For Each strLine As String In UserCtrlFrm.rtfEdss.Lines
                                If InStr(strLine, "TYPE BILL:") > 0 Then
                                    LineCount = LineCount + 1
                                    strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                                    While InStr(strMainLine, "|PAT STATUS CD:") = 0
                                        strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                                        If InStr(strMainLine, "ORIG CLM REF NUMBER:") > 0 Then
                                            Erase arrMainLine
                                            arrMainLine = strMainLine.Split(":")
                                            ICNDraft1 = arrMainLine(1).Replace("|", "")
                                            ICNDraft = Trim(ICNDraft1)
                                            HICN = Left(ICNDraft, 10)
                                            If IsAlphaNumeric(HICN) = False Then
                                                ICNDraft = PICN
                                                ICNDraft = PICN
                                                HICN = PICN
                                            End If
                                            If ICNDraft.Length > 11 Then
                                                Draft = Right(ICNDraft, 10)
                                                'Draft = ICNDraft.Substring(11, 10)
                                            End If

                                            If Left(HICN, 1) = "0" And ICNDraft.Length < 11 Then
                                                Draft = HICN
                                                HICN = ""
                                            Else
                                                HICN = HICN
                                            End If
                                            FindIcn = HICN
                                            If InStr(ICNList, HICN) > 0 OrElse Trim(HICN) <> "" Then
                                                FinalICN = HICN + "," + FinalICN
                                                Exit For
                                            Else
                                                CheckHardCopy = "Hard copy ICN is not found"
                                            End If
                                            ' Return FinalICN
                                            'arrMainLine = strMainLine.Split(New Char() {":"c, " "c}, StringSplitOptions.RemoveEmptyEntries)

                                        End If
                                        LineCount = LineCount + 1
                                    End While

                                End If
                            Next
                        End If
                        checkCount += 1


                    End If
                    If FindIcn = "" Then
                        Dim DraftComments As String = ""
                        GetHRHI(PICN, HistoryRHI, DraftComments)
                        If DraftComments = "Draft list is empty" Then
                            Continue For
                        End If
                        If InStr(strResearchComments, "W062NO CLAIMS ON FILE") > 0 Then
                            Continue For
                        End If
                        Dim kk As Integer = 0
                        Dim ll As Integer = 0
                        Dim count1 As Integer = 0
                        Dim TotalCount1 As Integer = 0
                        Dim Divresult1 As Integer = 0
                        Dim PercentageResult1 As Integer = 0
                        For rr As Integer = LBound(HistoryRHI) To UBound(HistoryRHI)
                            For uu As Integer = LBound(HistoryRHI(rr).RHIServiceLineCount.RevCode) To UBound(HistoryRHI(rr).RHIServiceLineCount.RevCode)
                                TotalCount1 += 1
                                HEDS1RevCode = HistoryRHI(rr).RHIServiceLineCount.RevCode(uu)
                                HEDS1Charge = HistoryRHI(rr).RHIServiceLineCount.Charge(uu)
                                HEDS1Date = HistoryRHI(rr).RHIServiceLineCount.DateOfService(uu)
                                HEDS1HCPC = HistoryRHI(rr).RHIServiceLineCount.ProcCode(uu)
                                For kk = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                    For ll = LBound(CurrentICNEDS1(kk).ServiceLine.RevCode) To UBound(CurrentICNEDS1(kk).ServiceLine.RevCode)
                                        EDS1RevCode = CurrentICNEDS1(kk).ServiceLine.RevCode(ll)
                                        EDS1Charge = CurrentICNEDS1(kk).ServiceLine.BillCharge(ll)
                                        EDS1HCPC = CurrentICNEDS1(kk).ServiceLine.CPTCode(ll)
                                        EDS1Date = CurrentICNEDS1(kk).ServiceLine.DateofService(ll)
                                        If HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC AndAlso HEDS1Date = EDS1Date Then
                                            count1 += 1
                                        End If
                                    Next
                                Next
                            Next
                            kk = 0
                            ll = 0
                        Next
                        Divresult1 = count1 / TotalCount1
                        PercentageResult1 = Divresult1 * 100
                        If PercentageResult1 >= 80 Then
                            If InStr(FinalICN, PICN) = 0 Then
                                FinalICN = PICN + "," + FinalICN
                            End If

                        End If
                    End If

                ElseIf CheckHardCopy = "Hard copy is not open" OrElse CheckHardCopy = "Hard copy ICN is not found" OrElse InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Then
                    Dim DraftComments As String = ""
                    'If InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Then
                    GetHRHI(PICN, HistoryRHI, DraftComments)
                    If DraftComments = "Draft list is empty" Then
                        Exit For
                    End If
                    Dim kk As Integer = 0
                    Dim ll As Integer = 0
                    Dim count1 As Integer = 0
                    Dim TotalCount1 As Integer = 0
                    Dim Divresult1 As Integer = 0
                    Dim PercentageResult1 As Integer = 0
                    For rr As Integer = LBound(HistoryRHI) To UBound(HistoryRHI)
                        For uu As Integer = LBound(HistoryRHI(rr).RHIServiceLineCount.RevCode) To UBound(HistoryRHI(rr).RHIServiceLineCount.RevCode)
                            TotalCount1 += 1
                            HEDS1RevCode = HistoryRHI(rr).RHIServiceLineCount.RevCode(uu)
                            HEDS1Charge = HistoryRHI(rr).RHIServiceLineCount.Charge(uu)
                            HEDS1Date = HistoryRHI(rr).RHIServiceLineCount.DateOfService(uu)
                            HEDS1HCPC = HistoryRHI(rr).RHIServiceLineCount.ProcCode(uu)
                            For kk = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                For ll = LBound(CurrentICNEDS1(kk).ServiceLine.RevCode) To UBound(CurrentICNEDS1(kk).ServiceLine.RevCode)
                                    EDS1RevCode = CurrentICNEDS1(kk).ServiceLine.RevCode(ll)
                                    EDS1Charge = CurrentICNEDS1(kk).ServiceLine.BillCharge(ll)
                                    EDS1HCPC = CurrentICNEDS1(kk).ServiceLine.CPTCode(ll)
                                    EDS1Date = CurrentICNEDS1(kk).ServiceLine.DateofService(ll)
                                    If HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC AndAlso HEDS1Date = EDS1Date Then
                                        count1 += 1
                                    End If
                                Next
                            Next
                        Next
                        kk = 0
                        ll = 0
                    Next
                    Divresult1 = count1 / TotalCount1
                    PercentageResult1 = Divresult1 * 100
                    If PercentageResult1 = 80 Then
                        FinalICN = PICN + "," + FinalICN
                    End If
                    'End If
                Else

                    If Trim(GetText(6, 2, 5)) = "EDS 1" Then
                        GetEDS1(HistoryICNEDS1, PICN)
                        Dim hh As Integer = 0
                        Dim nn As Integer = 0
                        Dim count As Integer = 0
                        Dim TotalCount As Integer = 0
                        Dim PercentageResult As Integer = 0
                        Dim DivResult As Integer = 0
                        For bb As Integer = LBound(HistoryICNEDS1) To UBound(HistoryICNEDS1)
                            For ll As Integer = LBound(HistoryICNEDS1(bb).ServiceLine.RevCode) To UBound(HistoryICNEDS1(bb).ServiceLine.RevCode)
                                TotalCount += 1
                                HEDS1RevCode = HistoryICNEDS1(bb).ServiceLine.RevCode(ll)
                                HEDS1Charge = HistoryICNEDS1(bb).ServiceLine.BillCharge(ll)
                                HEDS1Date = HistoryICNEDS1(bb).ServiceLine.DateofService(ll)
                                HEDS1HCPC = HistoryICNEDS1(bb).ServiceLine.CPTCode(ll)
                                For hh = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                    For nn = LBound(CurrentICNEDS1(hh).ServiceLine.RevCode) To UBound(CurrentICNEDS1(hh).ServiceLine.RevCode)
                                        EDS1RevCode = CurrentICNEDS1(hh).ServiceLine.RevCode(nn)
                                        EDS1Charge = CurrentICNEDS1(hh).ServiceLine.BillCharge(nn)
                                        EDS1HCPC = CurrentICNEDS1(hh).ServiceLine.CPTCode(nn)
                                        EDS1Date = CurrentICNEDS1(hh).ServiceLine.DateofService(nn)
                                        If HEDS1RevCode = EDS1RevCode AndAlso HEDS1Charge = EDS1Charge AndAlso HEDS1HCPC = EDS1HCPC AndAlso HEDS1Date = EDS1Date Then
                                            count += 1
                                        End If
                                    Next
                                Next
                            Next
                            hh = 0
                            nn = 0
                        Next
                        DivResult = count / TotalCount
                        PercentageResult = DivResult * 100

                        If PercentageResult >= 80 Then
                            FindIcn = PICN
                            If InStr(FinalICN, PICN) = 0 Then
                                FinalICN = PICN + "," + FinalICN
                            End If

                        End If

                    End If
                    If FindIcn = "" Then
                        If checkCount = 0 Then
                            Dim HICN As String = String.Empty
                            Dim LineCount As Integer = 0
                            Dim strMainLine As String
                            Dim LineCount1 As Integer = 0
                            Dim arrMainLine() As String
                            Dim FLN As String = CurrentClaimDetails.FLN
                            Dim DCC As String = CurrentClaimDetails.DCC
                            Dim getDoc360 As String = DOC360(FLN, DCC)
                            Dim ICNDraft As String = String.Empty
                            Dim ICNDraft1 As String = String.Empty
                            Dim Draft As String = String.Empty
                            If getDoc360 = "NA" Then
                                CheckHardCopy = "Hard copy is not open"
                            Else
                                UserCtrlFrm.rtfEdss.Text = getDoc360
                                For Each strLine As String In UserCtrlFrm.rtfEdss.Lines
                                    If InStr(strLine, "TYPE BILL:") > 0 Then
                                        LineCount = LineCount + 1
                                        strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                                        While InStr(strMainLine, "|PAT STATUS CD:") = 0
                                            strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                                            If InStr(strMainLine, "ORIG CLM REF NUMBER:") > 0 Then
                                                Erase arrMainLine
                                                arrMainLine = strMainLine.Split(":")
                                                ICNDraft1 = arrMainLine(1).Replace("|", "")
                                                ICNDraft = Trim(ICNDraft1)
                                                HICN = Left(ICNDraft, 10)
                                                If ICNDraft.Length > 11 Then
                                                    Draft = Right(ICNDraft, 10)
                                                End If

                                                If Left(HICN, 1) = "0" And ICNDraft.Length < 11 Then
                                                    Draft = HICN
                                                    HICN = ""
                                                Else
                                                    HICN = HICN
                                                End If

                                                If HICN = "" Then
                                                    strResearchComments = "Hard copy ICN is not found"
                                                    Return FinalICN = HICN
                                                    Exit Function
                                                ElseIf InStr(ICNList, HICN) > 0 OrElse Trim(HICN) <> "" Then
                                                    FinalICN = HICN + "," + FinalICN
                                                    Exit For
                                                Else
                                                    CheckHardCopy = "Hard copy ICN is not found"
                                                End If
                                                ' Return FinalICN
                                                'arrMainLine = strMainLine.Split(New Char() {":"c, " "c}, StringSplitOptions.RemoveEmptyEntries)

                                            End If
                                            LineCount = LineCount + 1
                                        End While

                                    End If
                                Next
                            End If
                            checkCount += 1


                        End If
                    End If

                End If

            Next
        Catch ex As Exception
            strResearchComments = GetCurrentMethod.Name() + " exception: " + ex.Message.ToString
        End Try
        Return FinalICN

    End Function

    Public Function RepositionArrayElements(ByVal icn() As String) As String()
        Dim len As Integer = icn.Length
        Dim arr(len) As String

        For i As Integer = 0 To icn.Length - 1
            arr(i + 1) = icn(i)
        Next
        arr(0) = CurrentClaimDetails.ICN
        Return arr
    End Function

    Public Function DOC360(ByVal FLN As String, ByVal DCC As String) As String
        Dim st As String
        Dim strFLNDCC As String
        Try
            strFLNDCC = FLN + DCC
            Dim wsShell = CreateObject("WScript.Shell")
            'Dim citrixUsername As String = wsShell.ExpandEnvironmentStrings("%USERNAME%")
            Dim citrixUsername As String = "rroy7"
            Dim objDoc As DOC360API.DOC360 = New DOC360API.DOC360()

            st = objDoc.GetDOC360(strFLNDCC, citrixUsername)
            If st = "NA" Then
                st = objDoc.GetDOC360(strFLNDCC, citrixUsername)
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return st
    End Function
    Public Sub PutJulianDate()
        Try
            Dim value As String
            Dim vl = String.Empty
            value = String.Empty
            Dim strDate As String = String.Empty
            Dim DateArray() As String
            value = CurrentICNEDS1(0).FDate
            'value = CurrentICNEDS9(0).JulianDate
            'strDate = ClaculateDateMonth(value)            Changed on 09/03/2022 to use from date from EDS1 in the RS DT field of eds6
            strDate = Left(value, 2) + "/" + Left(Right(value, 4), 2) + "/" + "20" + Right(value, 2)
            DateArray = strDate.Split("/")
            Dim LeftDate As String = String.Empty
            Dim MidDate As String = String.Empty
            Dim RightDate As String = String.Empty
            LeftDate = DateArray(0)
            MidDate = DateArray(1)
            RightDate = DateArray(2)
            SendKeyes("08", 22, 49)
            SendKeyes("R", 22, 61)
            SendKeyes(LeftDate, 22, 69)
            SendKeyes(MidDate, 22, 72)
            SendKeyes(RightDate, 22, 75)
            SendEnter()
            apiChk()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub

    Public Sub MakeAllowablePullOnMpc(sfx As Integer)
        Try

            'For SFX As Integer = 1 To MaxSuffix
            '    'SendClear()
            '    'SendClear()
            'Next
            pullEDSScreen(CurrentClaimDetails.ICN, 1)
            SendKeyes("A" & SFX, 1, 76)
            SendEnter()
            SendOverrideInScreen("01")
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If

            RemoveNotCoveredFromEDS6Point5()
            If strResearchComments.Contains("ENTER UNITS ON REPEATED LINE") Or strResearchComments.Contains("Cannot change pricing until 3C Is split") Then
                Exit Sub
            End If
            SendKeyes("A", 1, 75)
            SendEnter()
            SendOverrideInScreen("01")
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ReadPaymentScreen(ByRef MyPayment() As PaymentScreen, SuffixCount As Integer)
        Dim Allowable As String = String.Empty
        Dim Deductiable As String = String.Empty
        Dim Payment As String = String.Empty
        ReDim MyPayment(SuffixCount - 1)
        Dim count As Integer = 1
        'Dim I As Integer = 0
        Dim LineCount As Integer = 0
        Try
            WriteLog("==================ENTERING" + GetCurrentMethod.Name + "========================")

            For I As Integer = 1 To SuffixCount

                SendClear()
                SendClear()
                pullEDSScreen(CurrentClaimDetails.ICN, 1)
                SendKeyes("A" & I, 1, 76)
                SendEnter()
                'apiChk()
                'apiChk()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                If (GetText(10, 41, 5) = GetText(10, 50, 5) And GetText(10, 47, 2) = GetText(10, 56, 2)) And (GetText(10, 29, 2) = "11" Or GetText(10, 29, 2) = "--") Then
                End If
                MakeAllowablePullOnMpc(I)
                CheckBypassHistoryField()
                If GetText(2, 2, 2) <> "MP" Then
                    strResearchComments = "Cannot A Over to payment screen"
                    Exit Sub
                End If
                WriteLog(GetTextRect)
                'Loop added on 26th Sept,2022: to send 1 in NBR field for NCOPAY before collecting data from MPC
                For ff As Integer = 10 To 22 Step 2
                    If GetText(ff, 4, 6) = "------" Then Exit For
                    If GetText(ff, 4, 6) = "NCOPAY" Then
                        If Trim(GetText(ff, 25, 3)) <> "1" Then
                            SendKeyes("001", ff, 25)
                            apiChk()
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                    End If
                Next
                LineCount = 0
                count = 1
                For dd As Integer = 11 To 23 Step 2
                    Dim CheckLine As String = String.Empty

                    CheckLine = Trim(GetText(dd - 1, 11, 6))
                    If CheckLine = "------" Then Exit For
                    If LineCount = 0 Then
                        ReDim MyPayment(I - 1).Allowable(LineCount)
                        ReDim MyPayment(I - 1).Deductiable(LineCount)
                        ReDim MyPayment(I - 1).CoInsurance(LineCount)
                        ReDim MyPayment(I - 1).Payment(LineCount)
                        ReDim MyPayment(I - 1).LineCount(LineCount)
                    Else
                        ReDim Preserve MyPayment(I - 1).Allowable(LineCount)
                        ReDim Preserve MyPayment(I - 1).Deductiable(LineCount)
                        ReDim Preserve MyPayment(I - 1).CoInsurance(LineCount)
                        ReDim Preserve MyPayment(I - 1).Payment(LineCount)
                        ReDim Preserve MyPayment(I - 1).LineCount(LineCount)
                    End If
                    'Added If part on 27th June,2022  ---Sourav
                    If Trim(GetText(dd - 1, 4, 6)) = "NCOPAY" Then
                        Allowable = Trim(GetText(dd - 1, 50, 5)) + "." + Trim(GetText(dd - 1, 56, 2))
                    Else
                        Allowable = Trim(GetText(dd, 30, 5)) + "." + Trim(GetText(dd, 36, 2))
                    End If
                    If Allowable = ".00" Then
                        Allowable = "0.00"
                    End If
                    If InStr(Allowable, "-") > 0 Then
                        Allowable = Allowable.Replace("-", "0")
                    End If
                    'MyPayment(LineCount).Allowable(I) = Trim(GetText(dd, 30, 5)) + "." + Trim(GetText(dd, 36, 2))
                    MyPayment(I - 1).Allowable(LineCount) = Allowable
                    Deductiable = Trim(GetText(dd, 41, 5)) + "." + Trim(GetText(dd, 47, 2))
                    If Deductiable = ".00" Then
                        Deductiable = "0.00"
                    End If
                    If InStr(Deductiable, "-") > 0 Then
                        Deductiable = Deductiable.Replace("-", "0")
                    End If
                    'MyPayment(LineCount).Deductiable(I) = Trim(GetText(dd, 41, 5)) + "." + Trim(GetText(dd, 47, 2))
                    MyPayment(I - 1).Deductiable(LineCount) = Deductiable
                    MyPayment(I - 1).CoInsurance(LineCount) = Trim(GetText(dd, 55, 3)).Replace("-", "0")
                    Payment = Trim(GetText(dd, 61, 5)) + "." + Trim(GetText(dd, 67, 2))
                    If Payment = ".00" Then
                        Payment = "0.00"
                    End If
                    If InStr(Payment, "-") > 0 Then
                        Payment = Payment.Replace("-", "0")
                    End If
                    'MyPayment(LineCount).Payment(I) = Trim(GetText(dd, 61, 5)) + "." + Trim(GetText(dd, 67, 2))
                    MyPayment(I - 1).Payment(LineCount) = Payment
                    MyPayment(I - 1).LineCount(LineCount) = count
                    LineCount += 1
                    count += 1

                Next
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        Finally
            WriteLog(GetTextRect)
            WriteLog("==================EXIT" + GetCurrentMethod.Name + "========================")
        End Try
    End Sub
    'Public Sub HistoryMHIRead(ByRef HMHI() As MHI, ByVal HICN As String)
    '    SendKeyes("ret.i" + HICN + ",m", 2, 2)
    '    SendEnter()
    '    apiChk()
    '    ReDim HMHI(0)
    '    Dim LineCount As Integer = 0
    '    Dim x As Integer = 0
    '    For dd As Integer = 11 To 19 Step 4

    '    Next
    'End Sub
    Public Sub GetHRHI(ByVal icn As String, ByRef R() As HRHIScreen, ByRef DraftListComments As String)
        Try
            WriteLog("==================ENTERING " + GetCurrentMethod.Name + "========================")
            Dim DraftNoList As String = String.Empty
            Dim MHILine As String = String.Empty
            DraftNoList = HistoryDraft(icn)
            If String.IsNullOrEmpty(DraftNoList) Then
                DraftListComments = "Draft list is empty"
                Exit Sub
            End If
            Dim DraftList() As String = DraftNoList.Split(",")
            MHILine = GetMHIControlLine()
            ReDim R(0)

            Dim LineCount As Integer = -1
            Dim X As Integer = 0
            Dim I As Integer = 0
            Dim count As Integer = 0

            For dd As Integer = LBound(DraftList) To UBound(DraftList) - 1

                SendClear()
                SendClear()
                SendKeyes(MHILine + DraftList(dd) + ",I" + icn, 2, 2)
                SendKeyes("R", 2, 2)
                SendEnter()
                apiChk()

                'If Dcount = 0 Then
                '    Dcount += 1
                '    For ff As Integer = 5 To 79 Step 12
                '        If Trim(GetText(5, ff, 4)) = String.Empty Then Exit For
                '        If I = 0 Then
                '            ReDim D(0).DXCode(ff)
                '        Else
                '            ReDim Preserve D(0).DXCode(ff)
                '        End If
                '        D(0).DXCode(ff) = Trim(GetText(5, ff, 4))
                '        I += 1
                '    Next
                'End If
                If InStr(Trim(GetText(24, 1, 80)), "W062NO CLAIMS ON FILE") > 0 Then
                    strResearchComments = "RHI EDIT " + icn + ": W062NO CLAIMS ON FILE"
                    Exit Sub
                End If
                WriteLog($"RHI of {icn}:")
                WriteLog(GetTextRect())

                For hh As Integer = 8 To 23
                    If InStr(Trim(GetText(hh, 2, 3)), "SVC") > 0 Then
                        LineCount += 1
                        X = 0
                        count += 1

                        If LineCount = 0 Then
                            ReDim R(LineCount)
                        Else
                            ReDim Preserve R(LineCount)
                        End If
                        R(LineCount).RHIServiceLineCount.RHITranslation = Trim(GetText(hh, 7, 4))
                        R(LineCount).RHILineCount = count

                    End If
                    If IsNumeric(Trim(GetText(hh, 3, 4))) Then
                        If X = 0 Then

                            ReDim R(LineCount).RHIServiceLineCount.RevCode(X)
                            ReDim R(LineCount).RHIServiceLineCount.ProcCode(X)
                            ReDim R(LineCount).RHIServiceLineCount.DateOfService(X)
                            ReDim R(LineCount).RHIServiceLineCount.LineCount(X)
                            ReDim R(LineCount).RHIServiceLineCount.M1(X)
                            ReDim R(LineCount).RHIServiceLineCount.M2(X)
                            ReDim R(LineCount).RHIServiceLineCount.Charge(X)
                            ReDim R(LineCount).RHIServiceLineCount.Unit(X)
                            ReDim R(LineCount).RHIServiceLineCount.ServiceCode(X)
                        Else
                            ReDim Preserve R(LineCount).RHIServiceLineCount.RevCode(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.ProcCode(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.DateOfService(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.LineCount(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.M1(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.M2(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.Charge(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.Unit(X)
                            ReDim Preserve R(LineCount).RHIServiceLineCount.ServiceCode(X)
                        End If
                        R(LineCount).RHIServiceLineCount.RevCode(X) = Trim(GetText(hh, 3, 4))
                        R(LineCount).RHIServiceLineCount.ProcCode(X) = Trim(GetText(hh, 8, 5))
                        R(LineCount).RHIServiceLineCount.DateOfService(X) = Trim(GetText(hh, 28, 4))
                        R(LineCount).RHIServiceLineCount.Unit(X) = Trim(GetText(hh, 34, 5))
                        R(LineCount).RHIServiceLineCount.LineCount(X) = X
                        R(LineCount).RHIServiceLineCount.M1(X) = Trim(GetText(hh, 15, 2))
                        R(LineCount).RHIServiceLineCount.M2(X) = Trim(GetText(hh, 18, 2))
                        R(LineCount).RHIServiceLineCount.Charge(X) = Trim(GetText(hh, 43, 8))
                        R(LineCount).RHIServiceLineCount.ServiceCode(X) = Trim(GetText(hh - 2, 7, 5))
                        X += 1
                    End If


                    If hh = 23 Then
                        'SendF8()
                        If InStr(autECLPSObj.GetText(24, 3, 80), "RHI MAX HX LINES EXCEEDED") > 0 OrElse InStr(autECLPSObj.GetText(24, 3, 80), "W758 NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        SendF8()
                        hh = 7

                    End If

                Next

            Next
            WriteLog("==================Exit" + GetCurrentMethod.Name + "=====================")
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    Public Function GetMHIControlLine() As String
        Dim MHILine As String = Trim(GetText(2, 2, 36))
        Return MHILine
    End Function
    Public Function HistoryDraft(ByVal HICN As String) As String
        Dim DraftNo As String = String.Empty
        Dim TempDraftNO As String = String.Empty
        Try
            If HICN <> "" Then
                SendClear()
                SendClear()
                SendKeyes("ret,i" + HICN + ",m", 2, 2)
                SendEnter()
                apiChk()
                If InStr(Trim(GetText(24, 1, 80)), "E028NOT AUTHORIZED") > 0 Then Return "E028NOT AUTHORIZED"
                GetMHIControlLine()
                For dd As Integer = 12 To 20 Step 4

                    If InStr(Trim(GetText(dd, 1, 3)), "ICN") > 0 Then
                        If Trim(GetText(dd - 2, 28, 10)) <> "0000000000" Then

                            TempDraftNO = Trim(GetText(dd - 2, 28, 10))
                            If InStr(DraftNo, TempDraftNO) = 0 Then
                                DraftNo = TempDraftNO + "," + DraftNo
                            End If

                        End If
                    End If

                    If dd = 20 Then
                        'SendF8()
                        If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        SendF8()
                        dd = 8

                    End If
                Next
                strDraFT = DraftNo
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return DraftNo
    End Function
    Public Sub HistoryClaimDetails(ByRef H() As PaymentScreen, ByVal HistoryICN As String)
        Dim ICNList() As String = HistoryICN.Split(",")
        Dim ICN As String = String.Empty
        Dim Allowable As String = String.Empty
        Dim Deductiable As String = String.Empty
        Dim Payment As String = String.Empty
        ReDim H(0)
        Dim Linecount As Integer = 0
        Dim I As Integer = 0
        For dd As Integer = LBound(ICNList) To UBound(ICNList) - 1
            ICN = ICNList(dd)
            SendClear()
            SendClear()
            SendKeyes("ret,i" + ICN + ",m", 2, 2)
            SendEnter()
            apiChk()
            For tt As Integer = 11 To 19 Step 4

                If Trim(GetText(tt, 1, 1)) = "%" Then

                    If I = 0 Then
                        ReDim H(Linecount).LineCount(I)
                        ReDim H(Linecount).Allowable(I)
                        ReDim H(Linecount).Deductiable(I)
                        ReDim H(Linecount).CoInsurance(I)
                        ReDim H(Linecount).Payment(I)
                    Else
                        ReDim Preserve H(Linecount).LineCount(I)
                        ReDim Preserve H(Linecount).Allowable(I)
                        ReDim Preserve H(Linecount).Deductiable(I)
                        ReDim Preserve H(Linecount).CoInsurance(I)
                        ReDim Preserve H(Linecount).Payment(I)
                    End If
                    If InStr(Trim(GetText(tt, 1, 1)), "%") > 0 Then
                        H(Linecount).LineCount(I) = I
                        Allowable = Trim(GetText(tt, 31, 8))
                        If Allowable = ".00" Then
                            Allowable = "0.00"
                        End If
                        If InStr(Allowable, "-") > 0 Then
                            Allowable = "0.00"
                        End If
                        H(Linecount).Allowable(I) = Allowable
                        Deductiable = Trim(GetText(tt, 42, 7))
                        If Deductiable = ".00" Then
                            Deductiable = "0.00"
                        End If
                        If InStr(Deductiable, "-") > 0 Then
                            Deductiable = "0.00"
                        End If
                        H(Linecount).Deductiable(I) = Deductiable
                        H(Linecount).CoInsurance(I) = Trim(GetText(tt, 55, 3))
                        Payment = Trim(GetText(tt, 62, 8))
                        If Payment = ".00" Then
                            Payment = "0.00"
                        End If
                        If InStr(Payment, "-") > 0 Then
                            Payment = "0.00"
                        End If
                        H(Linecount).Payment(I) = Payment
                        I += 1
                    End If
                End If

                If tt = 19 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                    SendF8()
                    tt = 7

                End If

            Next

        Next

    End Sub

    Public Sub BasicHistoryClaimDetails(ByRef BH() As PaymentScreen, ByVal HistoryICN As String)
        Dim ICNList() As String = HistoryICN.Split(",")
        Dim ICN As String = String.Empty
        Dim Allowable As String = String.Empty
        Dim Deductiable As String = String.Empty
        Dim Payment As String = String.Empty
        ReDim BH(0)
        Dim Linecount As Integer = 0
        Dim I As Integer = 0

        For dd As Integer = 10 To 18 Step 4
            If Trim(GetText(dd + 1, 1, 1)) <> "%" Then Exit For
            If I = 0 Then
                ReDim BH(Linecount).LineCount(I)
                ReDim BH(Linecount).Allowable(I)
                ReDim BH(Linecount).Deductiable(I)
                ReDim BH(Linecount).CoInsurance(I)
                ReDim BH(Linecount).Payment(I)
            Else
                ReDim Preserve BH(Linecount).LineCount(I)
                ReDim Preserve BH(Linecount).Allowable(I)
                ReDim Preserve BH(Linecount).Deductiable(I)
                ReDim Preserve BH(Linecount).CoInsurance(I)
                ReDim Preserve BH(Linecount).Payment(I)
            End If
            BH(Linecount).LineCount(I) = I
            Allowable = Trim(GetText(dd, 60, 7))
            If Allowable = ".00" Then
                Allowable = "0.00"
            End If
            BH(Linecount).Allowable(I) = Allowable
            Deductiable = Trim(GetText(dd, 69, 7))
            If Deductiable = ".00" Then
                Deductiable = "0.00"
            End If
            BH(Linecount).Deductiable(I) = Deductiable
            BH(Linecount).CoInsurance(I) = Trim(GetText(dd + 1, 3, 3))
            Payment = Trim(GetText(dd + 1, 10, 7))
            If Payment = ".00" Then
                Payment = "0.00"
            End If
            BH(Linecount).Payment(I) = Payment
            I += 1
            If dd = 18 Then
                'SendF8()
                If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                SendF8()
                dd = 6

            End If

        Next


    End Sub
    Public Function ClaimDetermination(ByVal PaidAmount As String) As String
        Dim CPaymentAmount As Decimal = 0.0
        Dim HPaymentAmount As Decimal = 0.0
        Dim result As String = String.Empty
        Try
            If ICESDenialFlag = True Then
                For dd As Integer = LBound(CurrentIcnMPCScreenIfICES) To UBound(CurrentIcnMPCScreenIfICES)
                    For ee As Integer = LBound(CurrentIcnMPCScreenIfICES(dd).LineCount) To UBound(CurrentIcnMPCScreenIfICES(dd).LineCount)
                        CPaymentAmount = CDec(CurrentIcnMPCScreenIfICES(dd).Payment(ee)) + CPaymentAmount
                    Next
                Next
            Else
                For dd As Integer = LBound(CurreentICNPaymentScreen) To UBound(CurreentICNPaymentScreen)
                    For ee As Integer = LBound(CurreentICNPaymentScreen(dd).LineCount) To UBound(CurreentICNPaymentScreen(dd).LineCount)
                        CPaymentAmount = CDec(CurreentICNPaymentScreen(dd).Payment(ee)) + CPaymentAmount
                    Next
                Next
            End If
            currentIcnPaymentAmount = 0.0
            currentIcnPaymentAmount = CPaymentAmount
            'For hh As Integer = LBound(HistoryICNDetails) To UBound(HistoryICNDetails)
            '    For ll As Integer = LBound(HistoryICNDetails(hh).LineCount) To UBound(HistoryICNDetails(hh).LineCount)
            '        HPaymentAmount = CDec(HistoryICNDetails(hh).Payment(ll)) + HPaymentAmount
            '    Next
            'Next

            If PaidAmount > CPaymentAmount Then

                result = "Over Payment"
            End If
            If PaidAmount < CPaymentAmount Then

                result = "Under Payment"
            End If
            If PaidAmount = CPaymentAmount Then

                result = "Zero Payment"
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return result
    End Function
    Public Sub getOiOIM(ByRef isOI As Boolean, ByRef isOIM As Boolean, ByRef CurrentOII As OIIDetails)
        Try
            isOI = False
            isOIM = False
            If CurrentOII.OI IsNot Nothing Then
                For I As Integer = LBound(CurrentOII.OI) To UBound(CurrentOII.OI)
                    If CurrentOII.OI(I).Indicator <> "N" AndAlso CurrentOII.OI(I).Indicator <> "-" Then
                        isOI = True
                        Exit For
                    End If
                Next
            End If
            If CurrentOII.Medicare IsNot Nothing Then
                For I As Integer = LBound(CurrentOII.Medicare) To UBound(CurrentOII.Medicare)
                    If CurrentOII.Medicare(I).Indicator <> "N" AndAlso CurrentOII.Medicare(I).Indicator <> "-" Then
                        isOIM = True
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub

    Public Sub HistoryDetail(ByRef H() As HistoryDetails, ByRef HC() As HistoryClaimDeils, ByVal icn As String)
        Try
            SendClear()
            SendClear()
            SendKeyes("ret,i" + icn + ",m")
            SendEnter()
            apiChk()
            ReDim H(0)
            ReDim HC(0)

            Dim Linecount As Integer = 0

            Dim I As Integer = 0

            For dd As Integer = 10 To 18 Step 4

                Dim count As Integer = 0

                If Trim(GetText(dd + 2, 1, 3)) = "ICN" And count < 1 Then
                    count += 1
                    HC(0).TIN = Trim(GetText(dd, 5, 9))
                    HC(0).DCC = Trim(GetText(dd + 2, 49, 3))
                    HC(0).FLN = Trim(GetText(dd + 2, 34, 10))
                End If

                If Trim(GetText(dd + 1, 1, 1)) = "%" Then

                    If I = 0 Then

                        ReDim H(Linecount).POS(I)
                        ReDim H(Linecount).Draft(I)
                        ReDim H(Linecount).FDate(I)
                        ReDim H(Linecount).LDate(I)
                        ReDim H(Linecount).OV(I)
                        ReDim H(Linecount).RC(I)
                        ReDim H(Linecount).Line(I)

                    Else

                        ReDim Preserve H(Linecount).POS(I)
                        ReDim Preserve H(Linecount).Draft(I)
                        ReDim Preserve H(Linecount).FDate(I)
                        ReDim Preserve H(Linecount).LDate(I)
                        ReDim Preserve H(Linecount).OV(I)
                        ReDim Preserve H(Linecount).RC(I)
                        ReDim Preserve H(Linecount).Line(I)

                    End If

                    H(Linecount).POS(I) = Trim(GetText(dd, 1, 2))
                    H(Linecount).FDate(I) = Trim(GetText(dd, 11, 6))
                    H(Linecount).LDate(I) = Trim(GetText(dd, 18, 6))
                    H(Linecount).OV(I) = Trim(GetText(dd, 29, 2))
                    H(Linecount).RC(I) = Trim(GetText(dd, 36, 2))
                    I += 1
                End If
                If dd = 18 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                    SendF8()
                    dd = 6

                End If

            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub
    ''-----New Developemt Need To Change---------------
    Public Sub VoidHistoryClaim(ByRef FlagError As Boolean, ByRef HistoryList As String, ByVal CheckHistoryClaim As String, ByVal PaidAmount As String)
        strResearchComments = ""
        Dim ICN As String = String.Empty
        is30LineToBeused = False
        'For i As Integer = LBound(ICNList) To UBound(ICNList) - 1
        ICN = HistoryList
        Try
            HistoryDetail(HistoryAllDetail, HistoryDetailsICN, ICN)
            DxCodeFunctionForHistory(HistoryDxCode, ICN, FlagError)
            If FlagError And InStr(strResearchStatus, "ERROR in pulling RHI") > 0 Then
                strResearchComments = strResearchStatus
                Exit Sub
            End If
            MatchDxCode()
            'If CheckHistoryClaim = "Under Payment" OrElse CheckHistoryClaim = "Zero Payment" Then
            If DXCodeMatchFlag = True Then
                If HistoryDetailsICN(0).FLN <> CurrentClaimDetails.FLN AndAlso IsGZHWB5B6DeniedClaimWith951DCC(HistoryDetailsICN(0).DCC) = False Then
                    ProcessReversal(FlagError, ICN, PaidAmount) 'FalgHistoryVoid

                    If FlagError Then
                        Exit Sub
                    End If
                    is30LineToBeused = True
                ElseIf getSpecCD(HistoryDetailsICN(0).TIN) = getSpecCD(CurrentClaim(0).Tin) AndAlso IsGZHWB5B6DeniedClaimWith951DCC(HistoryDetailsICN(0).DCC) = False Then
                    ProcessReversal(FlagError, ICN, PaidAmount) 'FalgHistoryVoid
                    If FlagError Then
                        Exit Sub
                    End If
                    is30LineToBeused = True
                Else
                    is30LineToBeused = is30LineToBeused Or False
                End If
            Else
                strResearchComments = "Dxcode do not match"
                FlagError = True
                Exit Sub
            End If
            'Else
            '    strResearchComments = "claim is over payment"
            '    Exit Sub
            'End If
        Catch e As IndexOutOfRangeException
            WriteLog("Index out of range exception")
            strResearchComments = GetCurrentMethod.Name() + " method: Index out of range"
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try

        ' Next
    End Sub

    Public Function IsGZHWB5B6DeniedClaimWith951DCC(ByVal DCC As String) As Boolean 'ByRef MyClaim() As ClaimFromPaymentScreen,
        'If DCC <> "951" Then Return False
        'For I As Integer = LBound(MyClaim) To UBound(MyClaim)
        '    For J As Integer = LBound(MyClaim(I).POS) To UBound(MyClaim(I).POS)
        '        If InStr("GZ,HW,B5,B6", MyClaim(I).RemarkCode(J)) > 0 AndAlso MyClaim(I).RemarkCode(J).Trim <> "" Then
        '            Return True
        '        End If
        '    Next
        'Next
        'Return False
        If DCC <> "951" Then Return False
        For I As Integer = LBound(HistoryAllDetail) To UBound(HistoryAllDetail)
            For J As Integer = LBound(HistoryAllDetail(I).POS) To UBound(HistoryAllDetail(I).POS)
                If InStr("GZ,HW,B5,B6,C", HistoryAllDetail(I).RC(J)) > 0 AndAlso HistoryAllDetail(I).RC(J).Trim <> "" Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function
    Public Function IsGZHWB5B6DeniedClaimWith951DCC(ByRef MyClaim() As ClaimFromPaymentScreen, ByVal DCC As String) As Boolean 'ByRef MyClaim() As ClaimFromPaymentScreen,
        If DCC <> "951" Then Return False
        For I As Integer = LBound(MyClaim) To UBound(MyClaim)
            For J As Integer = LBound(MyClaim(I).POS) To UBound(MyClaim(I).POS)
                If InStr("GZ,HW,B5,B6", MyClaim(I).RemarkCode(J)) > 0 AndAlso MyClaim(I).RemarkCode(J).Trim <> "" Then
                    Return True
                End If
            Next
        Next
        Return False

    End Function


    Sub ProcessReversal(ByRef FlagError As Boolean, ByVal ICN As String, ByVal PaidAmount As String)
        Try
            Dim I As Integer, J As Integer, MyClaim(100) As Claim
            I = 0
            SendClear()
            SendClear()
            PullMHI(ICN)
            J = ReadMHI(MyClaim)

            Dim sb1 As New StringBuilder
            While I < J
                If MyClaim(I).Voided = False And CDbl(MyClaim(I).Draft) <> 0 And MyClaim(I).Closed(0) <> "1" Then
                    PerformVoid(FlagError, MyClaim(I).Draft, MyClaim(I).ICN, MyClaim(I).PaidAmount, MyClaim(I).DateOfProcessing)

                    'HistoryPaidAmount = PaidAmount + "," + PaidAmount
                    'If FlagError Then Exit Sub
                    If FlagError Then
                        If strResearchComments = "" Then
                            strResearchComments = "Skip from " + GetCurrentMethod.Name
                        Else
                            strResearchComments = "HxICN Void failed: " + strResearchComments
                        End If
                        Exit Sub
                    End If
                End If
                'HistoryPaidAmount = PaidAmount + "," + HistoryPaidAmount

                Dim draft As String = MyClaim(I).Draft

                If InStr(sb1.ToString, draft) = 0 Then
                    If sb1.ToString = "" Then
                        sb1.Append(draft)
                    Else
                        sb1.Append("," + draft)
                    End If
                End If

                'If HistoryPaidAmount = String.Empty Then
                '    HistoryPaidAmount = PaidAmount
                'Else
                '    HistoryPaidAmount = "," + PaidAmount
                'End If

                'strDraFT = draft + "," + strDraFT


                I += 1
            End While

            strDraFT = sb1.ToString
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
    End Sub

    Public Function getSpecCD(ByVal Tin As String) As String
        Dim strValue As String = ""
        SendClear()
        SendClear()
        SendKeyes("PMI,", 1, 2)
        SendEnter()
        SendKeyes(Left(Tin, 1), 3, 6)
        SendKeyes(Mid(Tin, 3, 9), 3, 8)
        SendKeyes(Right(Tin, 5), 3, 22)
        SendEnter()
        Dim SpecCode As String = GetText(13, 40, 3, True).Replace("-", "")
        Select Case SpecCode
            Case "A"
                strValue = "02"
            Case "ABS"
                strValue = "17"
            Case "ACU"
                strValue = "122"
            Case "ADC"
                strValue = "62"
            Case "ADE"
                strValue = "15"
            Case "ADL"
                strValue = "12"
            Case "ADP"
                strValue = "15"
            Case "AHF"
                strValue = "25"
            Case "AI"
                strValue = "02"
            Case "AM"
                strValue = "49"
            Case "AMB"
                strValue = "81"
            Case "AMS"
                strValue = "62"
            Case "AN"
                strValue = "20"
            Case "ANA"
                strValue = "20"
            Case "APL"
                strValue = "11"
            Case "ATH"
                strValue = "117"
            Case "ATP"
                strValue = "11"
            Case "AUD"
                strValue = "45"
            Case "BHA"
                strValue = "62"
            Case "BTM"
                strValue = "11"
            Case "C"
                strValue = "25"
            Case "CAR"
                strValue = "25"
            Case "CBG"
                strValue = "11"
            Case "CBM"
                strValue = "11"
            Case "CCG"
                strValue = "11"
            Case "CCM"
                strValue = "89"
            Case "CD "
                strValue = "25"
            Case "CDE"
                strValue = "109"
            Case "CDS"
                strValue = "37"
            Case "CGT"
                strValue = "11"
            Case "CLP"
                strValue = "11"
            Case "CM"
                strValue = "44"
            Case "CME"
                strValue = "116"
            Case "CMG"
                strValue = "11"
            Case "CMP"
                strValue = "11"
            Case "CMW"
                strValue = "126"
            Case "CNS"
                strValue = "116"
            Case "CNY"
                strValue = "05"
            Case "CPY"
                strValue = "11"
            Case "CRD"
                strValue = "62"
            Case "CRE"
                strValue = "25"
            Case "CRS"
                strValue = "14"
            Case "CRT"
                strValue = "48"
            Case "CSE"
                strValue = "116"
            Case "CSP"
                strValue = "118"
            Case "CSW"
                strValue = "62"
            Case "CVS"
                strValue = "111"
            Case "D"
                strValue = "03"
            Case "DCT"
                strValue = "31"
            Case "DD"
                strValue = "29"
            Case "DIA"
                strValue = "33"
            Case "DIR"
                strValue = "49"
            Case "DLI"
                strValue = "34"
            Case "DM"
                strValue = "80"
            Case "DME"
                strValue = "82"
            Case "DMP"
                strValue = "11"
            Case "DNP"
                strValue = "123"
            Case "DNT"
                strValue = "127"
            Case "DOP"
                strValue = "82"
            Case "DR"
                strValue = "16"
            Case "DRN"
                strValue = "16"
            Case "EAP"
                strValue = "62"
            Case "EC"
                strValue = "67"
            Case "EIN"
                strValue = "62"
            Case "EM"
                strValue = "39"
            Case "END"
                strValue = "33"
            Case "ENT"
                strValue = "77"
            Case "FMP"
                strValue = "74"
            Case "FNP"
                strValue = "116"
            Case "FOP"
                strValue = "11"
            Case "FP"
                strValue = "01"
            Case "FPN"
                strValue = "07"
            Case "FPS"
                strValue = "73"
            Case "FPY"
                strValue = "15"
            Case "FQH"
                strValue = "125"
            Case "FSE"
                strValue = "39"
            Case "GE"
                strValue = "29"
            Case "GEP"
                strValue = "15"
            Case "GER"
                strValue = "100"
            Case "GNC"
                strValue = "117"
            Case "GO"
                strValue = "106"
            Case "GP"
                strValue = "01"
            Case "GPM"
                strValue = "73"
            Case "GPS"
                strValue = "01"
            Case "GS"
                strValue = "17"
            Case "GYN"
                strValue = "07"
            Case "HAI"
                strValue = "60"
            Case "HEM"
                strValue = "30"
            Case "HH"
                strValue = "60"
            Case "HHA"
                strValue = "60"
            Case "HHD"
                strValue = "60"
            Case "HHI"
                strValue = "60"
            Case "HNS"
                strValue = "17"
            Case "HO"
                strValue = "36"
            Case "HOM"
                strValue = "118"
            Case "HOS"
                strValue = "128"
            Case "HPM"
                strValue = "107"
            Case "HRS"
                strValue = "45"
            Case "HS"
                strValue = "103"
            Case "HSP"
                strValue = "107"
            Case "ID"
                strValue = "92"
            Case "IG"
                strValue = "02"
            Case "IM"
                strValue = "04"
            Case "IMS"
                strValue = "76"
            Case "INC"
                strValue = "74"
            Case "IND"
                strValue = "73"
            Case "IP"
                strValue = "11"
            Case "IR"
                strValue = "112"
            Case "IVC"
                strValue = "25"
            Case "LAB"
                strValue = "34"
            Case "LAB"
                strValue = "35"
            Case "LAC"
                strValue = "117"
            Case "LAR"
                strValue = "10"
            Case "LIT"
                strValue = "49"
            Case "LM"
                strValue = "49"
            Case "LPC"
                strValue = "62"
            Case "LPN"
                strValue = "28"
            Case "LVN"
                strValue = "28"
            Case "MDS"
                strValue = "05"
            Case "MET"
                strValue = "11"
            Case "MFM"
                strValue = "74"
            Case "MFS"
                strValue = "114"
            Case "MFT"
                strValue = "62"
            Case "MFW"
                strValue = "62"
            Case "MGP"
                strValue = "11"
            Case "MGT"
                strValue = "11"
            Case "MHP"
                strValue = "62"
            Case "MM"
                strValue = "11"
            Case "MPD"
                strValue = "04"
            Case "MRE"
                strValue = "23"
            Case "MRI"
                strValue = "16"
            Case "MSW"
                strValue = "62"
            Case "MTH"
                strValue = "123"
            Case "N"
                strValue = "05"
            Case "NA"
                strValue = "20"
            Case "NCR"
                strValue = "25"
            Case "NDD"
                strValue = "05"
            Case "NDM"
                strValue = "05"
            Case "NE"
                strValue = "11"
            Case "NEP"
                strValue = "31"
            Case "NER"
                strValue = "16"
            Case "NET"
                strValue = "81"
            Case "NM"
                strValue = "38"
            Case "NPD"
                strValue = "36"
            Case "NPM"
                strValue = "47"
            Case "NPS"
                strValue = "05"
            Case "NR"
                strValue = "38"
            Case "NS"
                strValue = "06"
            Case "NSP"
                strValue = "06"
            Case "NSY"
                strValue = "61"
            Case "NTR"
                strValue = "63"
            Case "NW"
                strValue = "126"
            Case "OBG"
                strValue = "07"
            Case "OBS"
                strValue = "07"
            Case "OM"
                strValue = "23"
            Case "OMM"
                strValue = "113"
            Case "ON"
                strValue = "104"
            Case "OOF"
                strValue = "13"
            Case "OPH"
                strValue = "08"
            Case "OPR"
                strValue = "82"
            Case "OPT"
                strValue = "21"
            Case "OPY"
                strValue = "22"
            Case "ORD"
                strValue = "41"
            Case "ORP"
                strValue = "11"
            Case "ORS"
                strValue = "09"
            Case "OS"
                strValue = "49"
            Case "OT"
                strValue = "10"
            Case "OTH"
                strValue = "102"
            Case "OTO"
                strValue = "10"
            Case "OUT"
                strValue = "01"
            Case "P"
                strValue = "15"
            Case "PA"
                strValue = "11"
            Case "PAD"
                strValue = "12"
            Case "PAM"
                strValue = "115"
            Case "PAN"
                strValue = "20"
            Case "PAS"
                strValue = "117"
            Case "PCC"
                strValue = "89"
            Case "PCS"
                strValue = "18"
            Case "PD"
                strValue = "12"
            Case "PDA"
                strValue = "02"
            Case "PDC"
                strValue = "25"
            Case "PDE"
                strValue = "33"
            Case "PDG"
                strValue = "11"
            Case "PDL"
                strValue = "15"
            Case "PDM"
                strValue = "03"
            Case "PDN"
                strValue = "05"
            Case "PDO"
                strValue = "102"
            Case "PDP"
                strValue = "48"
            Case "PDR"
                strValue = "16"
            Case "PDS"
                strValue = "17"
            Case "PDT"
                strValue = "78"
            Case "PDU"
                strValue = "19"
            Case "PEM"
                strValue = "39"
            Case "PEN"
                strValue = "63"
            Case "PGS"
                strValue = "29"
            Case "PHA"
                strValue = "115"
            Case "PHE"
                strValue = "37"
            Case "PHL"
                strValue = "34"
            Case "PHM"
                strValue = "70"
            Case "PHO"
                strValue = "36"
            Case "PID"
                strValue = "92"
            Case "PLS"
                strValue = "13"
            Case "PM"
                strValue = "23"
            Case "PMG"
                strValue = "101"
            Case "PN"
                strValue = "120"
            Case "PNP"
                strValue = "31"
            Case "PNS"
                strValue = "120"
            Case "PNT"
                strValue = "116"
            Case "POD"
                strValue = "27"
            Case "POT"
                strValue = "10"
            Case "PPA"
                strValue = "11"
            Case "PPH"
                strValue = "73"
            Case "PPS"
                strValue = "13"
            Case "PPY"
                strValue = "61"
            Case "PRC"
                strValue = "76"
            Case "PRM"
                strValue = "23"
            Case "PRT"
                strValue = "42"
            Case "PRU"
                strValue = "32"
            Case "PS"
                strValue = "13"
            Case "PSA"
                strValue = "15"
            Case "PSH"
                strValue = "46"
            Case "PSN"
                strValue = "110"
            Case "PSP"
                strValue = "75"
            Case "PST"
                strValue = "79"
            Case "PSW"
                strValue = "62"
            Case "PT"
                strValue = "64"
            Case "PTH"
                strValue = "11"
            Case "PUD"
                strValue = "48"
            Case "PXR"
                strValue = "16"
            Case "PYA"
                strValue = "61"
            Case "PYG"
                strValue = "61"
            Case "R"
                strValue = "16"
            Case "RCT"
                strValue = "16"
            Case "REG"
                strValue = "62"
            Case "REN"
                strValue = "33"
            Case "RFA"
                strValue = "83"
            Case "RFC"
                strValue = "84"
            Case "RHC"
                strValue = "73"
            Case "RHI"
                strValue = "10"
            Case "RHU"
                strValue = "32"
            Case "RIR"
                strValue = "38"
            Case "RM"
                strValue = "23"
            Case "RN"
                strValue = "119"
            Case "RNA"
                strValue = "20"
            Case "RNP"
                strValue = "116"
            Case "ROC"
                strValue = "108"
            Case "RPY"
                strValue = "24"
            Case "SA"
                strValue = "62"
            Case "SAN"
                strValue = "121"
            Case "SCC"
                strValue = "89"
            Case "SCI"
                strValue = "06"
            Case "SCT"
                strValue = "68"
            Case "SLS"
                strValue = "34"
            Case "SO"
                strValue = "105"
            Case "SPM"
                strValue = "110"
            Case "SPT"
                strValue = "46"
            Case "TCS"
                strValue = "18"
            Case "THP"
                strValue = "29"
            Case "TR"
                strValue = "24"
            Case "TRS"
                strValue = "17"
            Case "TS"
                strValue = "18"
            Case "U"
                strValue = "19"
            Case "UCA"
                strValue = "66"
            Case "UHM"
                strValue = "49"
            Case "UNP"
                strValue = "122"
            Case "USM"
                strValue = "49"
            Case "USP"
                strValue = "49"
            Case "VIR"
                strValue = "112"
            Case "VIS"
                strValue = "60"
            Case "WNS"
                strValue = "119"
            Case "WSP"
                strValue = "124"

        End Select
        Return strValue
    End Function

    Public Function FindIndexInDraftDetailsByDraftNumber(ByRef myDraftDetails() As DraftDetails, ByVal Draft As String) As Integer
        For I As Integer = LBound(myDraftDetails) To UBound(myDraftDetails)
            If myDraftDetails(I).Draft = Draft Then Return I
        Next
        Return -1
    End Function

    Sub PerformVoid(ByRef FlagError As Boolean, ByVal Draft As String, ByVal ICN As String, ByVal strPaidAmount As String, ByVal ProcDate As String)
        WriteLog("=================Inside " + GetCurrentMethod.Name + "===============")
        WriteLog("Voiding " + ICN)
        Dim FlagFound As Boolean, K As Integer, ValidLline As Integer, Count2 As Integer, EditWarning As String = ""
        'MsgBox DateDiff("d", CDate(ProcDate), CDate(Format(Now, "MM/DD,YY")))
        'MsgBox(DateTime.Today)
        FlagError = False
        'WriteLog("Void starts for draft " & Draft & "and ICN " & ICN)
        'WriteLog("Processed Date " & ProcDate & " And Date.Today: " & CStr(CDate(System.TimeZoneInfo.ConvertTime(Now, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")))))
        Try
            If CDate(ProcDate) = CDate(System.TimeZoneInfo.ConvertTime(Now, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"))) Then 'CDate(DateTime.Today) Then   ''''''74 Void
                WriteLog($"inside If")
                SendClear()
                SendClear()
                SendKeyes("MPI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName & "," & CurrentClaimDetails.PatientRelationship & "," & Draft & ",I" & ICN, 2, 2)
                SendEnter()
                'WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                If InStr(GetText(24, 1, 80), "REVERSAL ON FILE") = 0 Then 'Already 74 void is performed
                    SendKeyes("GZN", 2, 2)
                    SendKeyes("74", 3, 78)
                    SendEnter()
                    'WriteLog("Before 74 void MPP for draft " & Draft & " and ICN " & ICN, True)
                    'WriteLog(GetTextRect())
                    Count2 = 0
                    EditWarning = ""
                    Do
                        SendEnter()
                        If Count2 > 50 Then Exit Do
                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        EditWarning = EditWarning & GetText(24, 1, 80, True)
                        Count2 = Count2 + 1
                    Loop
                    If Count2 > 50 Then
                        strResearchComments = EditWarning
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If
                    If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                        strResearchComments = "W342 NEW COINS REVERSAL"
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If
                    If FlagError = True Then
                        Exit Sub
                        'Return strResearchComments
                    End If
                    If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                        SendMPP()
                        WriteLog(GetTextRect)
                        If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                            strResearchComments = GetText(24, 1, 80, True)
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then

                            isHistoryICNVoided = True
                            flagHistoryVoid = True
                        End If
                    Else
                        strResearchComments = EditWarning
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If

                    'WriteLog("After 74 void MPP for draft " & Draft & " and ICN " & ICN, True)
                End If
            ElseIf DateDiff("d", CDate(ProcDate), CDate(DateTime.Today)) < 3 And CDbl(strPaidAmount) > 0 Then  'ELSEIF 1
                WriteLog("Inside ElseIf 1")
                '''''If Check not issued then 71 Void else 69/70
                SendKeyes("SCI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName & "," & CurrentClaimDetails.PatientRelationship & "," & Draft, 2, 2)
                SendEnter()
                'WriteLog("Control Line for SCI = " + GetText(2, 1, 80, True))
                If InStr(GetText(24, 1, 80), "E1006") > 0 Or InStr(GetText(24, 1, 80), "E1007") > 0 Then
                    SendClear()
                    SendClear()
                    SendKeyes("MPI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName & "," & CurrentClaimDetails.PatientRelationship & "," & Draft & ",I" & ICN, 2, 2)
                    SendEnter()
                    'WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                    If GetText(22, 36, 2, True) <> "71" Then 'Already 71 void has been performed
                        SendKeyes("GZ3", 2, 2)
                        SendKeyes("71", 3, 78)
                        SendEnter()
                        'WriteLog("Before 71 void MPP for draft " & Draft & " and ICN " & ICN, True)
                        'WriteLog(GetTextRect())
                        Count2 = 0
                        EditWarning = ""
                        Do
                            SendEnter()
                            If Count2 > 50 Then Exit Do
                            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                            EditWarning = EditWarning & GetText(24, 1, 80, True)
                            Count2 = Count2 + 1
                        Loop
                        If Count2 > 50 Then
                            strResearchComments = EditWarning
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        End If
                        If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                            strResearchComments = "W342 NEW COINS REVERSAL"
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        End If
                        If FlagError = True Then
                            Exit Sub
                            'Return strResearchComments
                        End If
                        If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                            SendMPP()
                            WriteLog(GetTextRect)

                            If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                strResearchComments = GetText(24, 1, 80, True)
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                isHistoryICNVoided = True
                                flagHistoryVoid = True
                            End If
                        Else
                            strResearchComments = EditWarning
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        End If
                        'WriteLog("After 71 void MPP for draft " & Draft & " and ICN " & ICN, True)
                    End If
                Else
                    SendClear()
                    SendClear()
                    SendKeyes("MPI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName & "," & CurrentClaimDetails.PatientRelationship & "," & Draft & ",I" & ICN, 2, 2)
                    SendEnter()
                    'WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                    If GetText(22, 36, 2, True) <> "70" Then
                        If InStr(GetText(24, 1, 80), "REVERSAL NOT FOUND") = 0 AndAlso InStr(GetText(24, 1, 80), "REVERSAL ON FILE") = 0 Then                 'Changes made if draft is closed or already 74 voided
                            SendKeyes("GZH", 2, 2)
                            SendKeyes("69", 3, 78)
                            SendEnter()
                            FlagFound = False
                            For K = 22 To 10 Step -2
                                If FlagFound = False Then
                                    If Trim(autECLPSObj.GetText(K, 4, 6)) <> "------" And Trim(autECLPSObj.GetText(K, 4, 6)) <> "NCOPAY" Then
                                        ValidLline = K
                                        FlagFound = True
                                    ElseIf Trim(autECLPSObj.GetText(K, 4, 6)) = "NCOPAY" Then
                                        SendKeyes("000", K, 25)
                                        SendKeyes("13", K, 29)
                                        SendKeyes("00000", K, 41)
                                        SendKeyes("00", K, 47)
                                        SendEnter()
                                    End If
                                Else
                                    SendKeyes("--", K, 1)
                                    SendKeyes("------", K, 4)
                                End If
                            Next
                            SendEnter()

                            ''''''''Check If RX

                            If CDbl(strPaidAmount) = 0 Then              ''''''''RX Void
                                SendKeyes("RX", ValidLline, 1)
                                SendKeyes("RX----", ValidLline, 4)
                                SendKeyes("13", ValidLline, 29)

                                SendKeyes("00000", ValidLline, 41)       ''''''''Bill Charge
                                SendKeyes("01", ValidLline, 47)

                                SendKeyes("00000", ValidLline, 50)      ''''''''Not Covered
                                SendKeyes("01", ValidLline, 56)

                                SendKeyes("00000", ValidLline + 1, 30)     ''''''''Allowed Amount
                                SendKeyes("00", ValidLline + 1, 36)

                                SendKeyes("00000", ValidLline + 1, 61) ''''''''Paid Amount
                                SendKeyes("00", ValidLline + 1, 67)

                                SendKeyes("00000", ValidLline + 1, 41)     ''''''''Deductible
                                SendKeyes("00", ValidLline + 1, 47)

                                SendKeyes("100", ValidLline + 1, 55)       ''''''''Coinsurance

                                SendKeyes("00000", ValidLline, 59)       ''''''''Basic Allowed
                                SendKeyes("00", ValidLline, 65)

                                SendKeyes("00000", ValidLline, 68)       ''''''''basic Deductible
                                SendKeyes("00", ValidLline, 74)

                                SendKeyes("000", ValidLline + 1, 3)       ''''''''basic coins

                                SendKeyes("00000", ValidLline + 1, 9)       ''''''''basic paid
                                SendKeyes("00", ValidLline + 1, 15)

                                SendEnter()
                                'WriteLog("Before RX void MPP for draft " & Draft & " and ICN " & ICN, True)
                                'WriteLog(GetTextRect())
                                Count2 = 0
                                EditWarning = ""
                                Do
                                    SendEnter()
                                    If Count2 > 50 Then Exit Do
                                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                    EditWarning = EditWarning & GetText(24, 1, 80, True)
                                    Count2 = Count2 + 1
                                Loop
                                If Count2 > 50 Then
                                    strResearchComments = EditWarning
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                                    strResearchComments = "W342 NEW COINS REVERSAL"
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                If FlagError = True Then
                                    Exit Sub
                                    'Return strResearchComments
                                End If
                                If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                    SendMPP()
                                    WriteLog(GetTextRect)
                                    If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                        strResearchComments = GetText(24, 1, 80, True)
                                        strResearchStatus = "skipped"
                                        FlagError = True
                                        Exit Sub
                                    ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                        isHistoryICNVoided = True
                                        flagHistoryVoid = True
                                    End If
                                Else
                                    'If InStr(EditWarning, "E740 VOID NOT COMPLETED") > 0 Then Exit Sub
                                    strResearchComments = EditWarning
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                'WriteLog("After RX void MPP for draft " & Draft & " and ICN " & ICN, True)
                            Else

                                SendKeyes("13", ValidLline, 29)          ''''''''Override
                                SendKeyes("00000", ValidLline, 50)
                                SendKeyes("00", ValidLline, 56)
                                SendKeyes("00000", ValidLline + 1, 30)
                                SendKeyes("00", ValidLline + 1, 36)
                                SendKeyes("00000", ValidLline + 1, 41)
                                SendKeyes("00", ValidLline + 1, 47)
                                SendKeyes("2", ValidLline, 32)
                                SendKeyes("1", ValidLline, 34)

                                If CDbl(GetText(ValidLline, 59, 5).Replace(" ", "0") + "." + GetText(ValidLline, 65, 2).Replace(" ", "0")) > 0 Then
                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 41)    ''''''''Bill Charge
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 47)

                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 59)       ''''''''Basic Allowed
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 65)

                                    SendKeyes("00000", ValidLline, 68)       ''''''''basic Deductible
                                    SendKeyes("00", ValidLline, 74)

                                    SendKeyes("100", ValidLline + 1, 3)       ''''''''basic coins

                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 9)       ''''''''basic paid
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 15)
                                Else
                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 41)    ''''''''Bill Charge
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 47)

                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 30)   ''''''''Allowed Amount
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 36)

                                    SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 61)     ''''''''Paid Amount
                                    SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 67)

                                    SendKeyes("100", ValidLline + 1, 55)        ''''''''Coinsurance
                                End If
                                SendEnter()
                                'WriteLog("Before 69-70 void MPP for draft " & Draft & " and ICN " & ICN, True)
                                'WriteLog(GetTextRect())
                                Count2 = 0
                                EditWarning = ""
                                Do
                                    SendEnter()
                                    If Count2 > 50 Then Exit Do
                                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                    EditWarning = EditWarning & GetText(24, 1, 80, True)
                                    Count2 = Count2 + 1
                                Loop
                                If Count2 > 50 Then
                                    strResearchComments = EditWarning
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                                    strResearchComments = "W342 NEW COINS REVERSAL"
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                If FlagError = True Then
                                    Exit Sub
                                    'Return strResearchComments
                                End If
                                If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                    SendMPP()
                                    WriteLog(GetTextRect)
                                    If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                        strResearchComments = GetText(24, 1, 80, True)
                                        strResearchStatus = "skipped"
                                        FlagError = True
                                        Exit Sub
                                    ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                        isHistoryICNVoided = True
                                        flagHistoryVoid = True
                                    End If
                                Else
                                    strResearchComments = EditWarning
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                End If
                                'WriteLog("After 69-70 void MPP for draft " & Draft & " and ICN " & ICN, True)
                            End If
                            If InStr(GetText(24, 1, 80), "ALL FILES") = 0 Then
                                FlagError = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Else
                WriteLog(">>>>>>>>>>> Inside Else (69/70 void) <<<<<<<<<<<<")
                '''''''''''''69/70 void
                SendClear()
                SendClear()
                SendKeyes("MPI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName & "," & CurrentClaimDetails.PatientRelationship & "," & Draft & ",I" & ICN, 2, 2)
                SendEnter()
                apiChk()
                apiChk()
                WriteLog(GetTextRect)
                'WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                If GetText(22, 36, 2, True) <> "70" AndAlso GetText(22, 36, 2, True) <> "OV" Then
                    If InStr(GetText(24, 1, 80), "REVERSAL NOT FOUND") = 0 AndAlso InStr(GetText(24, 1, 80), "REVERSAL ON FILE") = 0 Then                 'Changes made if draft is closed or already 74 voided
                        SendKeyes("GZH", 2, 2)
                        SendKeyes("69", 3, 78)
                        SendEnter()
                        FlagFound = False
                        For K = 22 To 10 Step -2
                            If FlagFound = False Then
                                If Trim(autECLPSObj.GetText(K, 4, 6)) <> "------" And Trim(autECLPSObj.GetText(K, 4, 6)) <> "NCOPAY" Then
                                    ValidLline = K
                                    FlagFound = True
                                ElseIf Trim(autECLPSObj.GetText(K, 4, 6)) = "NCOPAY" Then
                                    SendKeyes("000", K, 25)
                                    SendKeyes("13", K, 29)
                                    SendKeyes("00000", K, 41)
                                    SendKeyes("00", K, 47)
                                    SendEnter()
                                End If
                            Else
                                SendKeyes("--", K, 1)
                                SendKeyes("------", K, 4)
                            End If
                        Next
                        SendEnter()
                        Do
                            SendEnter()
                            If Count2 > 50 Then Exit Do
                            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                            EditWarning = EditWarning & GetText(24, 1, 80, True)
                            Count2 = Count2 + 1
                        Loop
                        If Count2 > 50 Then
                            strResearchComments = EditWarning
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        End If
                        If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                            strResearchComments = "W342 NEW COINS REVERSAL"
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        End If
                        If FlagError = True Then
                            Exit Sub
                            'Return strResearchComments
                        End If
                        If InStr(EditWarning, "WE689 INVALID HOSP DISC") > 0 Then
                            SendKeyes("-", 8, 25)
                            SendEnter()
                        End If
                        ''''''''Check If RX
                        For l = 22 To 10 Step -2
                            If Trim(autECLPSObj.GetText(l, 4, 6)) = "NCOPAY" Then
                                SendKeyes("--", l, 29)
                                SendEnter()
                            End If
                        Next

                        If CDbl(strPaidAmount) = 0 Then              ''''''''RX Void
                            SendKeyes("RX", ValidLline, 1)
                            SendKeyes("RX----", ValidLline, 4)
                            SendKeyes("13", ValidLline, 29)

                            SendKeyes("00000", ValidLline, 41)       ''''''''Bill Charge
                            SendKeyes("01", ValidLline, 47)

                            SendKeyes("00000", ValidLline, 50)       ''''''''Not Covered
                            SendKeyes("01", ValidLline, 56)

                            SendKeyes("00000", ValidLline + 1, 30)     ''''''''Allowed Amount
                            SendKeyes("00", ValidLline + 1, 36)

                            SendKeyes("00000", ValidLline + 1, 61)     ''''''''Paid Amount
                            SendKeyes("00", ValidLline + 1, 67)

                            SendKeyes("00000", ValidLline + 1, 41)     ''''''''Deductible
                            SendKeyes("00", ValidLline + 1, 47)

                            SendKeyes("100", ValidLline + 1, 55)       ''''''''Coinsurance

                            SendKeyes("00000", ValidLline, 59)       ''''''''Basic Allowed
                            SendKeyes("00", ValidLline, 65)

                            SendKeyes("00000", ValidLline, 68)       ''''''''basic Deductible
                            SendKeyes("00", ValidLline, 74)

                            SendKeyes("000", ValidLline + 1, 3)       ''''''''basic coins

                            SendKeyes("00000", ValidLline + 1, 9)       ''''''''basic paid
                            SendKeyes("00", ValidLline + 1, 15)

                            SendEnter()
                            'WriteLog("Before RX void MPP for draft " & Draft & " and ICN " & ICN, True)
                            'WriteLog(GetTextRect())
                            Count2 = 0
                            EditWarning = ""
                            Do
                                SendEnter()
                                If Count2 > 50 Then Exit Do
                                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                EditWarning = EditWarning & GetText(24, 1, 80, True)
                                Count2 = Count2 + 1
                            Loop
                            If Count2 > 50 Then
                                strResearchComments = EditWarning
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                                strResearchComments = "W342 NEW COINS REVERSAL"
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            'PasteHistoryICN(ICN, FlagError, strResearchStatus, strResearchComments)
                            If FlagError = True Then
                                Exit Sub
                                'Return strResearchComments
                            End If
                            If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                SendMPP()
                                WriteLog(GetTextRect)

                                If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                    strResearchComments = GetText(24, 1, 80, True)
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                    isHistoryICNVoided = True
                                    flagHistoryVoid = True
                                End If
                            Else
                                'If InStr(EditWarning, "E740 VOID NOT COMPLETED") > 0 Then Exit Sub
                                strResearchComments = EditWarning
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            'WriteLog("After RX void MPP for draft " & Draft & " and ICN " & ICN, True)
                        Else

                            SendKeyes("13", ValidLline, 29)          ''''''''Override
                            SendKeyes("00000", ValidLline, 50)
                            SendKeyes("00", ValidLline, 56)
                            SendKeyes("00000", ValidLline + 1, 30)
                            SendKeyes("00", ValidLline + 1, 36)
                            SendKeyes("00000", ValidLline + 1, 41)
                            SendKeyes("00", ValidLline + 1, 47)
                            SendKeyes("2", ValidLline, 32)
                            SendKeyes("1", ValidLline, 34)
                            If GetText(ValidLline, 59, 5) <> "-----" AndAlso CDbl(GetText(ValidLline, 59, 5).Replace(" ", "0") + "." + GetText(ValidLline, 65, 2).Replace(" ", "0")) > 0 Then
                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 41)    ''''''''Bill Charge
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 47)

                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 59)       ''''''''Basic Allowed
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 65)

                                SendKeyes("00000", ValidLline, 68)       ''''''''basic Deductible
                                SendKeyes("00", ValidLline, 74)

                                SendKeyes("100", ValidLline + 1, 3)       ''''''''basic coins

                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 9)       ''''''''basic paid
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 15)
                            Else
                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline, 41)    ''''''''Bill Charge
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline, 47)

                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 30)   ''''''''Allowed Amount
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 36)

                                SendKeyes(ConvertToFormat(Split(strPaidAmount, ".")(0), 5, True), ValidLline + 1, 61)     ''''''''Paid Amount
                                SendKeyes(Split(strPaidAmount, ".")(1), ValidLline + 1, 67)

                                SendKeyes("100", ValidLline + 1, 55)        ''''''''Coinsurance
                            End If

                            SendEnter()
                            apiChk()
                            apiChk()
                            'WriteLog("Before 69-70 void MPP for draft " & Draft & " and ICN " & ICN, True)
                            ' WriteLog(GetTextRect())
                            Count2 = 0
                            EditWarning = ""
                            Do
                                SendEnter()
                                If Count2 > 50 Then Exit Do
                                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                                EditWarning = EditWarning & GetText(24, 1, 80, True)
                                Count2 = Count2 + 1
                            Loop
                            If Count2 > 50 Then
                                strResearchComments = EditWarning
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            If InStr(EditWarning, "W342 NEW COINS REVERSAL") > 0 Then
                                strResearchComments = "W342 NEW COINS REVERSAL"
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            'PasteHistoryICN(ICN, FlagError, strResearchStatus, strResearchComments)
                            If FlagError = True Then
                                Exit Sub
                                'Return strResearchComments
                            End If
                            If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                                SendMPP()
                                WriteLog(GetTextRect)

                                If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                                    strResearchComments = GetText(24, 1, 80, True)
                                    strResearchStatus = "skipped"
                                    FlagError = True
                                    Exit Sub
                                ElseIf InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                                    isHistoryICNVoided = True
                                    flagHistoryVoid = True
                                End If
                            Else
                                strResearchComments = Trim(GetText(24, 1, 80))
                                strResearchStatus = "skipped"
                                FlagError = True
                                Exit Sub
                            End If
                            'WriteLog("After 69-70 void MPP for draft " & Draft & " and ICN " & ICN, True)
                        End If
                        If InStr(GetText(24, 1, 80), "ALL FILES") = 0 Then
                            FlagError = True
                            Exit Sub
                        End If
                    End If
                End If
                'If GetText(22, 36, 2, True) = "70" Or GetText(22, 36, 2, True) = "OV" Then
                '    WriteLog("History ICN already voided " + GetTextRect())
                'End If
            End If
        Catch ex As Exception
            FlagError = True
            WriteLog("-------------------EXCEPTION: " + GetCurrentMethod.Name() + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name() + " method: " + ex.Message.ToString
            Exit Sub
        End Try
        'WriteLog("Void ends for draft " & Draft & " and ICN " & ICN)
        'PullMHI(CurrentClaim(0).ICN)
        'CurrentClaimCount = ReadMHI(CurrentClaim)
        'WriteLog("MHI READ START After voiding draft " & Draft & " in ICN " & ICN)
        'WriteLog(GetMHI(CurrentClaim, CurrentClaimCount))
    End Sub
    'Public Sub ReadMHI1(ByRef C() As Claim)
    '    Dim Index As Integer, EditWarning As String, TempClosure As String = "", Position As Integer, Position2 As Integer, TempRemark As String
    '    Dim TempIsClosed As String, controlLine As String, LineCount As Integer, X As Integer ',Counter As Integer,
    '    Dim Temp, FindObject, Row, ELE, XX
    '    Dim DiagCode As String
    '    Dim CurrentScreen As String = ""
    '    Index = 0
    '    'MsgBox "MHI"
    '    C(Index).LineCount = 0
    '    Temp = ""
    '    TempRemark = ""
    '    TempIsClosed = ""
    '    controlLine = Trim(autECLPSObj.GetText(2, 1, 80))
    '    LineCount = 1
    '    Do
    '        For Row = 10 To 18 Step 4

    '            autECLPSObj.auteclfieldlist.Refresh
    '            apiChk()
    '            FindObject = autECLPSObj.auteclfieldlist.FindFieldByRowCol(Row, 1)
    '            If autECLPSObj.GetText(Row + 2, 1, 3) = "ICN" And FindObject.display Then
    '                'If Len(TempIsClosed) - Len(Replace(TempIsClosed, ",", "")) > 1 Then TempIsClosed = Left(TempIsClosed, Len(TempIsClosed) - 1)
    '                'If Len(TempRemark) - Len(Replace(TempRemark, "?", "")) > 1 Then TempRemark = Left(temremark, Len(TempRemark) - 2)
    '                ReDim C(Index).RHI(7)
    '                If TempIsClosed <> "" Then TempIsClosed = Microsoft.VisualBasic.Left(TempIsClosed, Len(TempIsClosed) - 1)
    '                If TempRemark <> "" And TempRemark <> "?" Then TempRemark = Microsoft.VisualBasic.Left(TempRemark, Len(TempRemark) - 1)
    '                C(Index).ICN = Trim(autECLPSObj.GetText(Row + 2, 5, 10))
    '                C(Index).Suffix = Trim(autECLPSObj.GetText(Row + 2, 16, 2))
    '                C(Index).FLN = Trim(autECLPSObj.GetText(Row + 2, 34, 10))
    '                C(Index).DCC = Trim(autECLPSObj.GetText(Row + 2, 49, 3))
    '                C(Index).ClaimNumber = Trim(autECLPSObj.GetText(8, 8, 3))
    '                C(Index).CauseCode = Trim(autECLPSObj.GetText(8, 42, 1))
    '                C(Index).Draft = Trim(autECLPSObj.GetText(Row, 28, 10))
    '                C(Index).Tin = Trim(autECLPSObj.GetText(Row, 3, 1)) & "-" &
    '                               Trim(autECLPSObj.GetText(Row, 5, 9)) & "-" &
    '                               Trim(autECLPSObj.GetText(Row, 15, 5))
    '                C(Index).BillCharge = autECLPSObj.GetText(Row + 1, 3, 9).replace(" ", "0")
    '                C(Index).PaidAmount = autECLPSObj.GetText(Row + 1, 16, 9).replace(" ", "0")
    '                C(Index).DateOfProcessing = autECLPSObj.GetText(Row, 39, 2) & "/" &
    '                                            autECLPSObj.GetText(Row, 41, 2) & "/" &
    '                                            autECLPSObj.GetText(Row, 43, 2)
    '                C(Index).AdjusterID = autECLPSObj.GetText(Row, 50, 6)
    '                C(Index).ServiceLine = Split(Temp, "?")
    '                'MsgBox c(0).ServiceLine(0)

    '                '''''''Changes made by Anirudha for Voided Check''''''''''

    '                If InStr(TempRemark, "69") > 0 Or InStr(TempRemark, "74") > 0 Or InStr(TempRemark, "71") > 0 Or InStr(TempRemark, "70") > 0 Then
    '                    C(Index).Voided = True
    '                Else
    '                    C(Index).Voided = False
    '                End If
    '                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '                C(Index).RemarkCode = Split(TempRemark, "?")

    '                If UBound(C(Index).RemarkCode) + 1 < C(Index).LineCount And UBound(C(Index).RemarkCode) > -1 Then
    '                    Temp = ""
    '                    For X = LBound(C(Index).RemarkCode) To UBound(C(Index).RemarkCode)
    '                        For Each ELE In Split(C(Index).RemarkCode(X), ",")
    '                            If Temp = "" Then
    '                                Temp = Temp & ELE & ","
    '                            Else
    '                                If InStr(Temp, ELE) = 0 Then
    '                                    Temp = Temp & ELE & ","
    '                                End If
    '                            End If
    '                        Next
    '                    Next
    '                    'MsgBox ""
    '                    'temp = C(Index).RemarkCode(0)
    '                    ReDim C(Index).RemarkCode(C(Index).LineCount)
    '                    TempRemark = ""
    '                    Temp = Microsoft.VisualBasic.Left(Temp, Len(Temp) - 1)
    '                    For X = 1 To C(Index).LineCount
    '                        TempRemark = TempRemark & Temp & "?"
    '                    Next
    '                    If TempRemark <> "" And TempRemark <> "?" Then TempRemark = Microsoft.VisualBasic.Left(TempRemark, Len(TempRemark) - 1)
    '                    C(Index).RemarkCode = Split(TempRemark, "?")
    '                End If
    '                '--------------POS fill up-------------------
    '                ReDim C(Index).POS(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).POS(II) = Mid(C(Index).ServiceLine(II), 1, 2)
    '                Next
    '                ReDim C(Index).ProcCode(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).ProcCode(II) = Trim(Mid(C(Index).ServiceLine(II), 4, 6))
    '                Next
    '                ReDim C(Index).FDOS(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).FDOS(II) = Mid(C(Index).ServiceLine(II), 11, 6)
    '                Next
    '                ReDim C(Index).LDOS(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).LDOS(II) = Mid(C(Index).ServiceLine(II), 18, 6)
    '                Next
    '                ReDim C(Index).NBR(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).NBR(II) = Mid(C(Index).ServiceLine(II), 25, 3).Replace(" ", "0")
    '                Next
    '                ReDim C(Index).OV(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).OV(II) = Mid(C(Index).ServiceLine(II), 29, 2).Trim
    '                Next

    '                ReDim C(Index).Copay(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).Copay(II) = "0.00"
    '                Next

    '                ReDim C(Index).Charge(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).Charge(II) = Mid(C(Index).ServiceLine(II), 41, 8).Replace(" ", "0")
    '                Next
    '                ReDim C(Index).NotCovered(UBound(C(Index).ServiceLine))

    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1

    '                    C(Index).NotCovered(II) = Mid(C(Index).ServiceLine(II), 50, 8).Replace(" ", "0")
    '                Next
    '                ReDim C(Index).Allowable(UBound(C(Index).ServiceLine))

    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1

    '                    C(Index).Allowable(II) = Mid(Split(C(Index).ServiceLine(II), "\\")(1), 30, 8).Replace(" ", "0")
    '                Next

    '                ReDim C(Index).Deductible(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).Deductible(II) = Mid(Split(C(Index).ServiceLine(II), "\\")(1), 41, 8).Replace(" ", "0")
    '                Next
    '                ReDim C(Index).CoIns(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).CoIns(II) = Mid(Split(C(Index).ServiceLine(II), "\\")(1), 55, 3).Trim
    '                Next
    '                ReDim C(Index).Paid(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).Paid(II) = Mid(Split(C(Index).ServiceLine(II), "\\")(1), 61, 8).Replace(" ", "0")
    '                Next
    '                ReDim C(Index).Modifier(UBound(C(Index).ServiceLine))
    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).Modifier(II) = ""
    '                Next
    '                'For II = LBound(c(Index).OV) To UBound(c(Index).OV) - 1
    '                '    c(Index).Line30 = False
    '                '    If c(Index).OV(II) = "30" Then
    '                '        c(Index).PaidAmount = CStr(CDbl(c(Index).PaidAmount) + CDbl(c(Index).Paid(II)))
    '                '        If InStr(c(Index).PaidAmount, ".") = 0 Then c(Index).PaidAmount = c(Index).PaidAmount & ".00"
    '                '        c(Index).PaidAmount = ConvertToFormat(c(Index).PaidAmount, 5, True, "0")
    '                '        c(Index).Line30 = True
    '                '        Exit For
    '                '    End If
    '                'Next
    '                C(Index).TotalAllowable = "0"
    '                C(Index).TotalDeductible = "0"
    '                C(Index).TotalNotCovered = "0"

    '                For II = LBound(C(Index).ServiceLine) To UBound(C(Index).ServiceLine) - 1
    '                    C(Index).TotalAllowable = CStr(CDbl(C(Index).TotalAllowable) + CDbl(C(Index).Allowable(II)))
    '                    C(Index).TotalDeductible = CStr(CDbl(C(Index).TotalDeductible) + CDbl(C(Index).Deductible(II)))
    '                    C(Index).TotalNotCovered = CStr(CDbl(C(Index).TotalNotCovered) + CDbl(C(Index).NotCovered(II)))
    '                Next
    '                C(Index).Closed = Split(TempIsClosed, ",")
    '                C(Index).Policy = autECLPSObj.GetText(2, 6, 6)
    '                C(Index).SSN = autECLPSObj.GetText(2, 13, 10)
    '                C(Index).EmployeeName = autECLPSObj.GetText(2, 24, 10)
    '                C(Index).Relationship = autECLPSObj.GetText(2, 35, 2)
    '                Index = Index + 1
    '                C(Index).LineCount = 0
    '                Temp = ""
    '                TempRemark = ""
    '                TempIsClosed = ""
    '                'LineCount = 1
    '            ElseIf FindObject.display Then
    '                'C(Index).ServiceLine(LineCount) = Trim(autECLPSObj.gettext(Row, 1, 80)) & "\\" & Trim(autECLPSObj.gettext(Row + 1, 1, 80))
    '                If Trim(autECLPSObj.GetText(Row, 29, 2)) = "C" Then
    '                    '\\ If closed
    '                    'C(Index).Closed(C(Index).LineCount) = 1 'True
    '                    TempIsClosed = TempIsClosed & "1" & ","
    '                    If Trim(autECLPSObj.GetText(Row, 36, 2)) = "7Y" And
    '                       Trim(autECLPSObj.GetText(Row, 4, 6)) <> "NCOPAY" Then
    '                        '\\ If EDS closed
    '                        EditWarning = autECLPSObj.GetText(24, 1, 80)
    '                        If LineCount = 1 Then '\\ First Line
    '                            If InStr(EditWarning, "L1:") Then
    '                                Position = InStr(EditWarning, "L1:") + 3
    '                                Position2 = InStr(EditWarning, "  ,")
    '                                TempClosure = Mid(EditWarning, Position, Position2 - Position)
    '                                If Len(Trim(TempClosure)) > 2 Then
    '                                    TempClosure = Trim(TempClosure)
    '                                End If
    '                                If InStr(TempClosure, " ") Then '\\ More than One closure code
    '                                    TempClosure = Replace(TempClosure, " ", ",")
    '                                Else

    '                                End If
    '                                TempRemark = TempRemark & TempClosure & "?"
    '                            End If
    '                        ElseIf LineCount = 2 Then  '\\ Second Line
    '                            If InStr(EditWarning, "L2:") Then
    '                                Position = InStr(EditWarning, "L2:") + 3
    '                                Position2 = InStr(Mid(EditWarning, InStr(EditWarning, "  ,") + 3, Len(EditWarning) - InStr(EditWarning, "  ,") - 3), "  ,") + InStr(EditWarning, "  ,") + 2
    '                                TempClosure = Mid(EditWarning, Position, Position2 - Position)
    '                                If Len(Trim(TempClosure)) > 2 Then
    '                                    TempClosure = Trim(TempClosure)
    '                                End If
    '                                If InStr(TempClosure, " ") Then  '\\ More than One closure code
    '                                    TempClosure = Replace(TempClosure, " ", ",")
    '                                Else

    '                                End If
    '                            End If
    '                            TempRemark = TempRemark & TempClosure & "?"

    '                        ElseIf LineCount = 3 Then  '\\ Third Line
    '                            If InStr(EditWarning, "L3:") Then
    '                                Position = InStr(EditWarning, "L3:") + 3
    '                                Position2 = InStr(Mid(EditWarning, InStr(Mid(EditWarning, InStr(EditWarning, "  ,") + 3, Len(EditWarning) - InStr(EditWarning, "  ,") - 3), "  ,") + InStr(EditWarning, "  ,") + 2 + 3, (Len(EditWarning) - (InStr(Mid(EditWarning, InStr(EditWarning, "  ,") + 3, Len(EditWarning) - InStr(EditWarning, "  ,") - 3), "  ,") + InStr(EditWarning, "  ,") + 2))), "  ,") + InStr(Mid(EditWarning, InStr(EditWarning, "  ,") + 3, Len(EditWarning) - InStr(EditWarning, "  ,") - 3), "  ,") + InStr(EditWarning, "  ,") + 2 + 2
    '                                'Position2 = InStr(EditWarning, "  ,")+
    '                                TempClosure = Mid(EditWarning, Position, Position2 - Position)
    '                                If Len(Trim(TempClosure)) > 2 Then
    '                                    TempClosure = Trim(TempClosure)
    '                                End If
    '                                If InStr(TempClosure, " ") Then '\\ More than One closure code
    '                                    TempClosure = Replace(TempClosure, " ", ",")
    '                                Else

    '                                End If
    '                            End If
    '                            TempRemark = TempRemark & TempClosure & "?"

    '                        End If
    '                    Else
    '                        '\\ If not EDS closed
    '                        TempRemark = TempRemark & Trim(autECLPSObj.GetText(Row, 36, 2)) & "?"
    '                    End If
    '                Else
    '                    '\\ If not Closed
    '                    TempIsClosed = TempIsClosed & "0" & ","
    '                    TempRemark = TempRemark & Trim(autECLPSObj.GetText(Row, 36, 2)) & "?"
    '                End If
    '                If Trim(autECLPSObj.GetText(Row, 1, 80)) <> "" And Trim(autECLPSObj.GetText(Row + 1, 1, 80)) <> "" Then
    '                    Temp = Temp & autECLPSObj.GetText(Row, 1, 80) &
    '                                  "\\" & autECLPSObj.GetText(Row + 1, 1, 80) & "?"
    '                    C(Index).LineCount = C(Index).LineCount + 1

    '                End If
    '                LineCount = LineCount + 1
    '            Else
    '                '    GoTo ExitLoop
    '            End If

    '            If Row = 18 Then
    '                If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") OrElse InStr(autECLPSObj.GetText(24, 1, 80), "W062NO CLAIMS ON FILE") Then '"W062NO CLAIMS ON FILE"
    '                    Exit Do
    '                Else
    '                    CurrentScreen = GetTextRect(10, 1, 20, 79)
    '                    Do
    '                        SendF8()
    '                    Loop While CurrentScreen = GetTextRect(10, 1, 20, 79)
    '                    'autECLPSObj.SendKeys("[pf8]")
    '                    'apiChk()

    '                    If InStr(autECLPSObj.GetText(24, 1, 80), "NEGATIVE AMOUNT") Then
    '                        autECLPSObj.SendKeys("[pf8]")
    '                        apiChk()
    '                    End If
    '                    Row = 6
    '                    LineCount = 1
    '                End If
    '            End If
    '        Next
    '    Loop

    'End Sub

    'Public Sub HistoryClaimDXCode(ByVal icn As String)

    '    SendClear()
    '    SendClear()
    '    SendKeyes("mei,i" + icn + ",1", 2, 2)
    '    SendEnter()
    '    apiChk()
    '    SendKeyes("dx", 1, 76)
    '    SendEnter()
    '    apiChk()
    '    GetDXCode(HistoryDxCode)

    'End Sub
    Public Function MatchDxCode() As Boolean
        Try
            Dim HDxCode As String = String.Empty
            Dim CDxCode As String = String.Empty
            Dim MaxCount As Integer = 0
            Dim hd As Integer = 0
            Dim ee As Integer = 0
            Dim DxCodeMatch As Boolean = False
            For cc As Integer = LBound(HistoryDxCode) To UBound(HistoryDxCode)
                For dd As Integer = LBound(HistoryDxCode(cc).DXCode) To UBound(HistoryDxCode(cc).DXCode)
                    HDxCode = HistoryDxCode(cc).DXCode(dd)
                    For hd = LBound(CurrentDXCode) To UBound(CurrentDXCode)
                        For ee = LBound(CurrentDXCode(hd).DXCode) To UBound(CurrentDXCode(hd).DXCode)
                            CDxCode = CurrentDXCode(hd).DXCode(ee)
                            If HDxCode = CDxCode Then
                                MaxCount += 1

                            End If
                        Next
                    Next
                    hd = 0
                    ee = 0
                Next
            Next
            If MaxCount > 0 Then
                DXCodeMatchFlag = True
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try
        Return DXCodeMatchFlag
    End Function
    Public Function DetermineIfClaimIsInVoided71Or74(ByRef Mydrafts() As DraftDetails, ByVal MaxSuffix As Integer) As Boolean
        Dim Flag As Boolean = False
        For I As Integer = 1 To MaxSuffix
            'Flag = False
            If UBound(Mydrafts) = 0 And Mydrafts(0).Suffix = "" Then Return False
            For J = LBound(Mydrafts) To UBound(Mydrafts)
                If CInt(Mydrafts(J).Suffix) = I Then
                    'If InStr("74,71,Not Voided", Mydrafts(J).ReversalType) > 0 Then
                    '    Flag = False
                    'Else
                    '    Flag = True
                    'End If
                    'Flag = Mydrafts(J).Reversal
                    'Exit For
                    If Mydrafts(J).ReversalType = "71" OrElse Mydrafts(J).ReversalType = "71 Void" OrElse Mydrafts(J).ReversalType = "74" Then
                        Return True
                    End If
                    Exit For
                Else
                    'Flag = True
                End If
            Next
            'If Flag = True Then Return False
        Next
        Return False
    End Function

    Public Sub ReadPaymentScreenForLog(ByRef Myclaim() As ClaimFromPaymentScreen, ByVal MaxSuffix As Integer)
        Try
            WriteLog("===============ENTERING " + GetCurrentMethod.Name + "==============")
            WriteLog(GetTextRect)
            Dim c As Integer = 0
            Dim CountOfLine As Integer = -1
            Dim EditWarning As String = String.Empty
            Dim count As Integer
            ReDim DenialRC(100)
            'dt_denialRC.Columns.Clear()

            Initialize_dt_denialRC(dt_denialRC)
            ReDim Myclaim(MaxSuffix - 1)
            For I As Integer = 1 To MaxSuffix
                CountOfLine = -1
                SendClear()
                SendClear()
                pullEDSScreen(CurrICN, 2)
                SendKeyes("A" & I, 1, 76)
                SendEnter()
                apiChk()
                initialOperationsToBeDoneWhenPaymentScreenIsPulled(I)
                CheckBypassHistoryField(I)
                For J As Integer = 10 To 22 Step 2
                    If GetText(J, 1, 23, True).Replace("-", "").Trim = "" Then Exit For
                    Dim POS As String = GetText(J, 1, 2)
                    Dim ProcCode As String = GetText(J, 4, 6)
                    Dim Fdos As String = GetText(J, 11, 6)
                    Dim Ldos As String = GetText(J, 18, 6)
                    Dim NBR As String = GetText(J, 25, 3)
                    Dim OV As String = GetText(J, 29, 2)
                    Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                    Dim Remark As String = GetText(J, 36, 2)
                    Dim charge As String = GetText(J, 41, 8, True).Replace(" ", ".")
                    Dim NotCove As String = GetText(J, 50, 8, True).Replace(" ", ".")
                    Dim Allwable As String = GetText(J + 1, 30, 8, True).Replace(" ", ".")
                    Dim Deductible As String = GetText(J + 1, 41, 8, True).Replace(" ", ".")
                    Dim Coins As String = GetText(J + 1, 55, 3, True)
                    Dim Paid As String = GetText(J + 1, 61, 8, True).Replace(" ", ".")
                    If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                        CountOfLine += 1
                    End If
                Next


                ReDim Myclaim(I - 1).POS(CountOfLine)
                ReDim Myclaim(I - 1).ProcCode(CountOfLine)
                ReDim Myclaim(I - 1).FDOS(CountOfLine)
                ReDim Myclaim(I - 1).LDOS(CountOfLine)
                ReDim Myclaim(I - 1).NBR(CountOfLine)
                ReDim Myclaim(I - 1).Override(CountOfLine)
                ReDim Myclaim(I - 1).PayeeCode(CountOfLine)
                ReDim Myclaim(I - 1).RemarkCode(CountOfLine)
                ReDim Myclaim(I - 1).Charge(CountOfLine)
                ReDim Myclaim(I - 1).NotCovered(CountOfLine)
                ReDim Myclaim(I - 1).Allowable(CountOfLine)
                ReDim Myclaim(I - 1).Deductible(CountOfLine)
                ReDim Myclaim(I - 1).Coins(CountOfLine)
                ReDim Myclaim(I - 1).Paid(CountOfLine)
                ReDim Myclaim(I - 1).Sanction(CountOfLine)
                ReDim Myclaim(I - 1).HxCindex(CountOfLine)
                ReDim Myclaim(I - 1).HxLIndex(CountOfLine)
                ReDim Myclaim(I - 1).Modifier(CountOfLine)

                Myclaim(I - 1).CauseCode = GetText(7, 4, 1)
                Myclaim(I - 1).Suffix = I
                Myclaim(I - 1).Draft = "0000000000"
                Myclaim(I - 1).DateOfProcessing = "99/99/99"
                Dim TotalPaid As String = "0"
                CountOfLine = 0
                For J As Integer = 10 To 22 Step 2
                    If GetText(J, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                    Dim POS As String = GetText(J, 1, 2)
                    Dim ProcCode As String = GetText(J, 4, 6)
                    Dim Fdos As String = GetText(J, 11, 6)
                    Dim Ldos As String = GetText(J, 18, 6)
                    Dim NBR As String = GetText(J, 25, 3).Replace(" ", "0")
                    Dim OV As String = GetText(J, 29, 2)
                    Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                    Dim Remark As String = GetText(J, 36, 2)
                    Dim charge As String = GetText(J, 41, 5, True).Replace("-", "0") & "." & GetText(J, 47, 2, True) 'GetText(J, 41, 8, True).Replace(" ", ".")
                    Dim NotCove As String = GetText(J, 50, 5, True).Replace("-", "0") & "." & GetText(J, 56, 2, True) 'GetText(J, 50, 8, True).Replace(" ", ".")
                    Dim Allwable As String = GetText(J + 1, 30, 5, True).Replace("-", "0") & "." & GetText(J + 1, 36, 2, True) 'GetText(J + 1, 30, 8, True).Replace(" ", ".")
                    Dim Deductible As String = GetText(J + 1, 41, 5, True).Replace("-", "0") & "." & GetText(J + 1, 47, 2, True) 'GetText(J + 1, 41, 8, True).Replace(" ", ".")
                    Dim Coins As String = GetText(J + 1, 55, 3, True)
                    Dim Paid As String = GetText(J + 1, 61, 5, True).Replace("-", "0") & "." & GetText(J + 1, 67, 2, True) 'GetText(J + 1, 61, 8, True).Replace(" ", ".")
                    Dim Sanction As String = GetText(J, 39, 1)
                    If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                        Myclaim(I - 1).POS(CountOfLine) = POS
                        Myclaim(I - 1).ProcCode(CountOfLine) = ProcCode.Trim
                        Myclaim(I - 1).FDOS(CountOfLine) = Fdos
                        Myclaim(I - 1).LDOS(CountOfLine) = Ldos
                        Myclaim(I - 1).NBR(CountOfLine) = NBR.Replace("-", "0")
                        Myclaim(I - 1).Override(CountOfLine) = OV
                        Myclaim(I - 1).PayeeCode(CountOfLine) = Payee
                        Myclaim(I - 1).RemarkCode(CountOfLine) = Remark
                        Myclaim(I - 1).Charge(CountOfLine) = charge.Replace("-", "0")
                        Myclaim(I - 1).NotCovered(CountOfLine) = NotCove.Replace("-", "0")
                        Myclaim(I - 1).Allowable(CountOfLine) = Allwable.Replace("-", "0")
                        Myclaim(I - 1).Deductible(CountOfLine) = Deductible.Replace("-", "0")
                        Myclaim(I - 1).Coins(CountOfLine) = Coins.Replace("-", "0")
                        Myclaim(I - 1).Paid(CountOfLine) = Paid.Replace("-", "0")
                        'TotalPaid = CStr(CDbl(TotalPaid) + CDbl(Myclaim(I - 1).Paid(CountOfLine)))
                        TotalPaid = CStr(Val(TotalPaid) + Val(Myclaim(I - 1).Paid(CountOfLine)))
                        Myclaim(I - 1).Sanction(CountOfLine) = Sanction
                        Myclaim(I - 1).HxCindex(CountOfLine) = I - 1
                        Myclaim(I - 1).HxLIndex(CountOfLine) = CountOfLine
                        CountOfLine += 1
                    End If
                    Myclaim(I - 1).TotalPaid = TotalPaid
                Next


                If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                    SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                    If GetText(2, 2, 2) <> "MP" Then
                        Dim s = GetText(2, 2, 3)
                        strResearchComments = "Unable to remove PZ1, pulling " + s + " screen"
                        Exit Sub
                    End If
                    SendOverrideInScreen("01")
                    CaptureScreenLog()
                End If
                Do
                    If EditWarning = String.Empty Then
                        EditWarning = String.Empty
                        count = 0
                    End If
                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                    If count > 50 Then Exit Do
                    EditWarning += GetText(24, 1, 80, True)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                Loop
                If InStr(EditWarning, "W3057CRE RULES REVW REQD") > 0 Or InStr(EditWarning, "WE3056CRE RULES REVW REQD") > 0 Then
                    strResearchComments = "W3057CRE RULES REVIEW REQD"
                End If
                EditWarningCheck(EditWarning)
                Myclaim(I - 1).EditWarning = EditWarning

                For K As Integer = 10 To 22 Step 2
                    If GetText(K, 1, 2) = "--" Then Exit For
                    If GetText(K, 36, 2) = "QC" Then flag_QC = True
                    If InStr(DenialRemarkCodeList, GetText(K, 36, 2)) > 0 Then
                        flagDenial = True
                        DenialRC(c) = GetText(K, 36, 2)
                        dt_denialRC.Rows.Add(GetText(K, 36, 2), GetText(K, 41, 5) + "." + GetText(K, 47, 2), I)
                        c += 1
                    End If
                Next
            Next

            'For I As Integer = 1 To MaxSuffix
            '    Myclaim(I - 1).EditWarning = EditWarning
            'Next

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " =>" + ex.Message.ToString)
        Finally
            WriteLog("===============EXITING " + GetCurrentMethod.Name + "==============")
        End Try
    End Sub
    ''' Need to integrated''''''''''
    Public Sub ReadClaimPaymentScreen(ByRef Myclaim() As ClaimFromPaymentScreen, ByRef MyDrafts() As DraftDetails, ByRef CurrentClaim() As Claim, ByVal CureentClaimCount As Integer, ByRef FlagError As Boolean, ByVal ICN As String, ByVal MaxSuffix As Integer, Optional ByVal IsCurrentClaim As Boolean = False)
        Try
            WriteLog("===============ENTERING " + GetCurrentMethod.Name + "==============")
            Dim Suffixes As String = "", drafts As String = ""
            Dim CountOfLine As Integer = -1
            Dim ReversalDraft As String = ""
            If IsCurrentClaim Then
                For I As Integer = 1 To MaxSuffix
                    For J As Integer = LBound(MyDrafts) To UBound(MyDrafts)
                        If IsNumeric(MyDrafts(J).Suffix) AndAlso CInt(MyDrafts(J).Suffix) = I Then
                            If MyDrafts(J).Reversal = False Then
                                Suffixes += MyDrafts(J).Suffix & ","
                                drafts += MyDrafts(J).Draft & ","
                            End If
                            Exit For
                        End If
                    Next
                Next
            Else
                For I As Integer = 1 To MaxSuffix
                    For J As Integer = LBound(MyDrafts) To UBound(MyDrafts)
                        If IsNumeric(MyDrafts(J).Suffix) AndAlso CInt(MyDrafts(J).Suffix) = I Then
                            If MyDrafts(J).Reversal = False OrElse MyDrafts(J).ReversalType = "" OrElse InStr("74,71", MyDrafts(J).ReversalType) = 0 Then
                                Suffixes += MyDrafts(J).Suffix & ","
                                drafts += MyDrafts(J).Draft & ","
                            End If
                            If InStr(Suffixes, ConvertToFormat(CStr(I), 2, True, "0")) > 0 Then Exit For
                        End If
                    Next
                    'For J As Integer = 0 To GetClaimCountFromStruct(HistoryClaim) - 1
                    '    If HistoryClaim(J).Voided = True Then

                    '    Else

                    '    End If
                    'Next
                Next
            End If

            'For I As Integer = LBound(MyDrafts) To UBound(MyDrafts)
            '    If IsNumeric(MyDrafts(I).Draft) AndAlso (MyDrafts(I).ReversalType = "Not Voided" OrElse InStr("74,71", MyDrafts(I).ReversalType) = 0) Then
            '        If InStr(Suffixes, MyDrafts(I).Suffix) = 0 Then
            '            Suffixes += MyDrafts(I).Suffix & ","
            '            drafts += MyDrafts(I).Draft & ","
            '        End If
            '    End If
            'Next
            ReDimension(Myclaim, MaxSuffix - 1)
            For I As Integer = 1 To MaxSuffix
                Dim EditWarning As String
                Dim Count2 As Integer = 0
                CountOfLine = -1
                If ICN = CurrICN Then
                    SendClear()
                    SendClear()
                    pullEDSScreen(ICN, 2)
                    SendKeyes("A" & I, 1, 76)
                    SendEnter()
                    apiChk()
                    EditWarning = initialOperationsToBeDoneWhenPaymentScreenIsPulled(I)
                    If (GetText(10, 41, 5) = GetText(10, 50, 5) And GetText(10, 47, 2) = GetText(10, 56, 2)) And (GetText(10, 29, 2) = "11" Or GetText(10, 29, 2) = "--") Then
                    End If
                    MakeAllowablePullOnMpc(I)
                    CheckBypassHistoryField(I)
                End If
                If strResearchComments.Contains("ENTER UNITS ON REPEATED LINE") Or strResearchComments.Contains("Cannot change pricing until 3C Is split") Then
                    Exit Sub
                End If
                Do
                    SendEnter()
                    If Count2 > 50 Then Exit Do
                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                    EditWarning = EditWarning & GetText(24, 1, 80, True)
                    Count2 = Count2 + 1
                Loop
                If InStr(Suffixes, ConvertToFormat(CStr(I), 2, True, "0")) > 0 Then
                    'copy from mhi structure
                    For J = 0 To CureentClaimCount - 1
                        If CInt(CurrentClaim(J).Suffix) = I And InStr(drafts, CurrentClaim(J).Draft) > 0 And CurrentClaim(J).Voided = False And (CurrentClaim(J).RemarkCode(0).Trim = "" OrElse InStr("77,75,OV,42,87,AC", CurrentClaim(J).RemarkCode(0)) = 0) Then
                            ReDim Myclaim(I - 1).POS(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).ProcCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).FDOS(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).LDOS(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).NBR(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Override(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).PayeeCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).RemarkCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Charge(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).NotCovered(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Allowable(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Deductible(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Coins(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Paid(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Sanction(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).HxCindex(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).HxLIndex(UBound(CurrentClaim(J).ServiceLine) - 1)
                            ReDim Myclaim(I - 1).Modifier(UBound(CurrentClaim(J).ServiceLine) - 1)
                            Myclaim(I - 1).CauseCode = CurrentClaim(J).CauseCode
                            Myclaim(I - 1).Suffix = CInt(CurrentClaim(J).Suffix)
                            Myclaim(I - 1).Draft = CurrentClaim(J).Draft
                            Myclaim(I - 1).DateOfProcessing = CurrentClaim(J).DateOfProcessing
                            Myclaim(I - 1).TotalPaid = CurrentClaim(J).PaidAmount
                            For K = 0 To UBound(CurrentClaim(J).ServiceLine) - 1
                                Myclaim(I - 1).POS(K) = CurrentClaim(J).POS(K)
                                Myclaim(I - 1).ProcCode(K) = CurrentClaim(J).ProcCode(K)
                                Myclaim(I - 1).FDOS(K) = CurrentClaim(J).FDOS(K)
                                Myclaim(I - 1).LDOS(K) = CurrentClaim(J).LDOS(K)
                                Myclaim(I - 1).NBR(K) = CurrentClaim(J).NBR(K)
                                Myclaim(I - 1).Override(K) = CurrentClaim(J).OV(K)
                                Myclaim(I - 1).PayeeCode(K) = "21"
                                Myclaim(I - 1).RemarkCode(K) = CurrentClaim(J).RemarkCode(K)
                                Myclaim(I - 1).Charge(K) = CurrentClaim(J).Charge(K)
                                Myclaim(I - 1).NotCovered(K) = CurrentClaim(J).NotCovered(K)
                                Myclaim(I - 1).Allowable(K) = CurrentClaim(J).Allowable(K)
                                Myclaim(I - 1).Deductible(K) = CurrentClaim(J).Deductible(K)
                                Myclaim(I - 1).Coins(K) = CurrentClaim(J).CoIns(K)
                                Myclaim(I - 1).Sanction(K) = Mid(CurrentClaim(J).ServiceLine(K), 39, 1)
                                Myclaim(I - 1).Paid(K) = CurrentClaim(J).Paid(K)
                                Myclaim(I - 1).HxCindex(K) = J
                                Myclaim(I - 1).HxLIndex(K) = K
                            Next
                            Exit For
                        End If
                    Next
                Else
                    'Copy frm payment screen
                    If Trim(GetText(2, 2, 2)) <> "MP" Then
                        pullEDSScreen(ICN, "9")
                        If InStr(GetText(24, 1, 80), "NO CLAIMS FOUND WITH EMPLOYEE ID SPECIFIED") > 0 Then
                            GoTo LoopAgain
                        End If
                        SendKeyes("A" & I, 1, 76)
                        SendEnter()
                        IncreementMPCCount()
                    End If
                    If InStr(GetText(24, 1, 80), "CLAIM AWAITING AUTO ADJUDICATION") > 0 Then
                        FlagError = True
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "CLAIM AWAITING AUTO ADJUDICATION"
                        Exit Sub
                    End If
                    If GetText(2, 2, 2) = "AR" Then
                        FlagError = True
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "Notification Related Claim"
                        Exit Sub
                    End If
                    If GetText(2, 2, 2) = "MR" Then
                        FlagError = True
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "Patient information mismatch"
                        Exit Sub
                    End If
                    If GetText(2, 2, 3) <> "MPC" Then 'AndAlso ICN = CurrICN Then
                        FlagError = True
                        'WriteLog(GetTextRect())
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "unable to pull payment screen"
                        Exit Sub
                    End If
                    'in case no lines pended
                    If InStr(GetText(24, 1, 80, True), "WE1191NO LINES PENDED") > 0 Or InStr(GetText(24, 1, 80, True), "WE1196DRAFT NBR REQUIRED") > 0 Or InStr(GetText(24, 1, 80, True), "E944INCOMPLETE DATA") > 0 Then 'Or InStr(GetText(24, 1, 80, True), "FLN NOT NUMERIC") > 0 Then
                        For J = LBound(MyDrafts) To UBound(MyDrafts)
                            If CInt(MyDrafts(J).Suffix) = I Then
                                SendClear()
                                SendClear()
                                SendClear()
                                Dim AdjConLine As String = "ADJ," + CurrentClaim(0).Policy + "," + CurrentClaim(0).SSN + "," +
                                                   CurrentClaim(0).EmployeeName + "," + CurrentClaim(0).Relationship + "," +
                                                   MyDrafts(J).Draft + ",I" + CurrentClaim(0).ICN
                                SendKeyes(AdjConLine, 2, 2)
                                SendEnter()
                                Exit For
                            End If
                        Next
                    End If
                    If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                        SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                        If GetText(2, 2, 2) <> "MP" Then
                            Dim s = GetText(2, 2, 3)
                            strResearchComments = "Unable to remove PZ1, pulling " + s + " screen"
                            Exit Sub
                        End If
                        SendOverrideInScreen("01")
                        CaptureScreenLog()
                    End If
                    RemoveBlankAndPISLLine()
                    'ICESCheck(CurrICN)
                    If strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE" Then
                        Exit Sub
                    End If
                    If strResearchComments = "unable to remove pisl line" Then
                        Exit Sub
                    End If
                    SendOverrideInScreen("01")
                    Dim EditWarning1 As String = String.Empty
                    Dim Count As Integer
                    Do
                        If EditWarning1 = String.Empty Then
                            EditWarning1 = String.Empty
                            Count = 0
                        End If
                        If InStr(EditWarning1, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        If Count > 50 Then Exit Do
                        EditWarning1 += GetText(24, 1, 80, True)
                        SendEnter()
                        Count += 1
                    Loop
                    apiChk()
                    apiChk()
                    EditWarningCheck(EditWarning1)
                    If InStr(EditWarning1, "RSC APLY JK") > 0 Then      'Added by Sourav 22/05/24 so that mod50 claims don't get skipped
                        EditWarning1 = EditWarning1.Replace("WE2570 RSC APLY JK", "")
                    End If
                    'If InStr(EditWarning1, "WE") > 0 Then
                    '    Dim count1 As Integer = 0
                    '    Do
                    '        If InStr(GetText(24, 1, 80), "WE") > 0 Or InStr(GetText(24, 1, 80), ",E") > 0 Then
                    '            Exit Do
                    '        ElseIf count1 > 50 Then
                    '            Dim indx1 As Integer = InStr(EditWarning1, "WE")
                    '            Dim len As Integer = 25
                    '            Dim txt As String = EditWarning1.Substring(indx1, len)
                    '            strResearchComments = "HARD EDIT - E# or WE Found: " + txt
                    '            Exit Sub
                    '        Else
                    '            count1 += 1
                    '            SendEnter()
                    '        End If
                    '    Loop
                    '    Dim position As Integer = GetText(24, 1, 80).IndexOf(",")
                    '    strResearchComments = "HARD EDIT - E# or WE Found: " + GetText(24, 1, position).Trim
                    '    'strResearchComments = "HARD EDIT - E# or WE Found"
                    '    Exit Sub
                    'End If
                    'If InStr(EditWarning1, ",E") > 0 Then
                    '    Dim count1 As Integer = 0
                    '    Do
                    '        If InStr(GetText(24, 1, 80), ",E") > 0 Then
                    '            Exit Do
                    '        ElseIf count1 > 50 Then
                    '            Dim indx1 As Integer = InStr(EditWarning1, ",E")
                    '            Dim len As Integer = 25
                    '            Dim txt As String = EditWarning1.Substring(indx1, len)
                    '            strResearchComments = "HARD EDIT - E# or WE Found: " + txt
                    '            Exit Sub
                    '        Else
                    '            count1 += 1
                    '            SendEnter()
                    '        End If
                    '    Loop
                    '    Dim position As Integer = GetText(24, 1, 80).IndexOf(",")
                    '    strResearchComments = "HARD EDIT - E# or WE Found: " + GetText(24, 1, position).Trim
                    '    'strResearchComments = "HARD EDIT - E# or WE Found"
                    '    Exit Sub
                    'End If
                    'If CheckWhetherAllEDS1LinesPulledInPaymentScreen(CurrentClaimModifier, I) = False Then
                    '    strResearchStatus = "SKIPPED"
                    '    strResearchComments = "All Eds1 Lines are not pulled in payment screen"
                    '    FlagError = True
                    '    Exit Sub
                    'End If
                    If checkForPendPMCMADCV() Then
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "Pend PM/CM/CV/AD present in claim"
                        FlagError = True
                        Exit Sub
                    End If
                    If CheckIfMultipleCopayLineWithNoNBR(I) Then
                        FlagError = True
                        strResearchStatus = "skipped"
                        strResearchComments = "more than 2 copay line present with 0 units. manual intervention required."
                        Exit Sub
                    End If
                    'EditWarning = initialOperationsToBeDoneWhenPaymentScreenIsPulled(I)

                    If InStr(EditWarning1, "W512") > 0 Then 'Chk Man Spi
                        FlagError = True
                        strResearchStatus = "skipped"
                        strResearchComments = "Check manual spi exists."
                        Exit Sub
                        'ExitAllAndReturnWithException()
                        Exit Sub
                    End If
                    'count lines in payment screen
                    For J As Integer = 10 To 22 Step 2
                        If GetText(J, 1, 23, True).Replace("-", "").Trim = "" Then Exit For
                        Dim POS As String = GetText(J, 1, 2)
                        Dim ProcCode As String = GetText(J, 4, 6)
                        Dim Fdos As String = GetText(J, 11, 6)
                        Dim Ldos As String = GetText(J, 18, 6)
                        Dim NBR As String = GetText(J, 25, 3)
                        Dim OV As String = GetText(J, 29, 2)
                        Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                        Dim Remark As String = GetText(J, 36, 2)
                        Dim charge As String = GetText(J, 41, 8, True).Replace(" ", ".")
                        Dim NotCove As String = GetText(J, 50, 8, True).Replace(" ", ".")
                        Dim Allwable As String = GetText(J + 1, 30, 8, True).Replace(" ", ".")
                        Dim Deductible As String = GetText(J + 1, 41, 8, True).Replace(" ", ".")
                        Dim Coins As String = GetText(J + 1, 55, 3, True)
                        Dim Paid As String = GetText(J + 1, 61, 8, True).Replace(" ", ".")
                        If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                            CountOfLine += 1
                        End If
                    Next


                    ReDim Myclaim(I - 1).POS(CountOfLine)
                    ReDim Myclaim(I - 1).ProcCode(CountOfLine)
                    ReDim Myclaim(I - 1).FDOS(CountOfLine)
                    ReDim Myclaim(I - 1).LDOS(CountOfLine)
                    ReDim Myclaim(I - 1).NBR(CountOfLine)
                    ReDim Myclaim(I - 1).Override(CountOfLine)
                    ReDim Myclaim(I - 1).PayeeCode(CountOfLine)
                    ReDim Myclaim(I - 1).RemarkCode(CountOfLine)
                    ReDim Myclaim(I - 1).Charge(CountOfLine)
                    ReDim Myclaim(I - 1).NotCovered(CountOfLine)
                    ReDim Myclaim(I - 1).Allowable(CountOfLine)
                    ReDim Myclaim(I - 1).Deductible(CountOfLine)
                    ReDim Myclaim(I - 1).Coins(CountOfLine)
                    ReDim Myclaim(I - 1).Paid(CountOfLine)
                    ReDim Myclaim(I - 1).Sanction(CountOfLine)
                    ReDim Myclaim(I - 1).HxCindex(CountOfLine)
                    ReDim Myclaim(I - 1).HxLIndex(CountOfLine)
                    ReDim Myclaim(I - 1).Modifier(CountOfLine)
                    Myclaim(I - 1).EditWarning = EditWarning1
                    Myclaim(I - 1).CauseCode = GetText(7, 4, 1)
                    Myclaim(I - 1).Suffix = I
                    Myclaim(I - 1).Draft = "0000000000"
                    Myclaim(I - 1).DateOfProcessing = "99/99/99"
                    Dim TotalPaid As String = "0"
                    CountOfLine = 0
                    For J As Integer = 10 To 22 Step 2
                        If GetText(J, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                        Dim POS As String = GetText(J, 1, 2)
                        Dim ProcCode As String = GetText(J, 4, 6)
                        Dim Fdos As String = GetText(J, 11, 6)
                        Dim Ldos As String = GetText(J, 18, 6)
                        Dim NBR As String = GetText(J, 25, 3).Replace(" ", "0")
                        Dim OV As String = GetText(J, 29, 2)
                        Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                        Dim Remark As String = GetText(J, 36, 2)
                        Dim charge As String = GetText(J, 41, 5, True).Replace("-", "0") & "." & GetText(J, 47, 2, True) 'GetText(J, 41, 8, True).Replace(" ", ".")
                        Dim NotCove As String = GetText(J, 50, 5, True).Replace("-", "0") & "." & GetText(J, 56, 2, True) 'GetText(J, 50, 8, True).Replace(" ", ".")
                        Dim Allwable As String = GetText(J + 1, 30, 5, True).Replace("-", "0") & "." & GetText(J + 1, 36, 2, True) 'GetText(J + 1, 30, 8, True).Replace(" ", ".")
                        Dim Deductible As String = GetText(J + 1, 41, 5, True).Replace("-", "0") & "." & GetText(J + 1, 47, 2, True) 'GetText(J + 1, 41, 8, True).Replace(" ", ".")
                        Dim Coins As String = GetText(J + 1, 55, 3, True)
                        Dim Paid As String = GetText(J + 1, 61, 5, True).Replace("-", "0") & "." & GetText(J + 1, 67, 2, True) 'GetText(J + 1, 61, 8, True).Replace(" ", ".")
                        Dim Sanction As String = GetText(J, 39, 1)
                        If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                            Myclaim(I - 1).POS(CountOfLine) = POS
                            Myclaim(I - 1).ProcCode(CountOfLine) = ProcCode.Trim
                            Myclaim(I - 1).FDOS(CountOfLine) = Fdos
                            Myclaim(I - 1).LDOS(CountOfLine) = Ldos
                            Myclaim(I - 1).NBR(CountOfLine) = NBR.Replace("-", "0")
                            Myclaim(I - 1).Override(CountOfLine) = OV
                            Myclaim(I - 1).PayeeCode(CountOfLine) = Payee
                            Myclaim(I - 1).RemarkCode(CountOfLine) = Remark
                            Myclaim(I - 1).Charge(CountOfLine) = charge.Replace("-", "0")
                            Myclaim(I - 1).NotCovered(CountOfLine) = NotCove.Replace("-", "0")
                            Myclaim(I - 1).Allowable(CountOfLine) = Allwable.Replace("-", "0")
                            Myclaim(I - 1).Deductible(CountOfLine) = Deductible.Replace("-", "0")
                            Myclaim(I - 1).Coins(CountOfLine) = Coins.Replace("-", "0")
                            Myclaim(I - 1).Paid(CountOfLine) = Paid.Replace("-", "0")
                            TotalPaid = CStr(Val(TotalPaid) + Val(Myclaim(I - 1).Paid(CountOfLine)))
                            Myclaim(I - 1).Sanction(CountOfLine) = Sanction
                            Myclaim(I - 1).HxCindex(CountOfLine) = I - 1
                            Myclaim(I - 1).HxLIndex(CountOfLine) = CountOfLine
                            CountOfLine += 1
                        End If
                        Myclaim(I - 1).TotalPaid = TotalPaid
                    Next
                    'Goto EDS Screen and remove 13 override
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    SendOverrideInScreen("01")
                    'For KK As Integer = 6 To 18 Step 3
                    '    SendKeyes("opn", KK, 2)
                    '    SendEnter()
                    '    apiChk()
                    '    For nn As Integer = 8 To 17 Step 3
                    '        If Trim(GetText(nn, 7, 4)) = String.Empty Then Exit For
                    '        SendKeyes("------", nn, 57)
                    '        SendKeyes("--", nn, 64)
                    '        SendKeyes("--", nn, 67)
                    '        SendKeyes("------", nn, 71)
                    '        SendKeyes("--", nn, 78)
                    '        If nn = 17 Then
                    '            EDS6Point5ValueMatch(PreviosValue)
                    '            SendF8()
                    '            EDS6Point5ValueMatch(CurrentValue)

                    '            If EDS6Point5IndexMatch() = False Then
                    '                nn = 5
                    '            Else
                    '                Exit For
                    '            End If

                    '        End If



                    '    Next

                    '    SendKeyes("Y", 2, 49)
                    '    SendEnter()
                    '    SendF3()
                    'Next
                    SendKeyes("N", 4, 79)
                    SendEnter()

                    If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                        For cc As Integer = 7 To 19 Step 3
                            If GetText(cc - 1, 8, 9) <> "         " Then
                                autECLPSObj.SendKeys("--------", cc, 6)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 20)
                                apiChk()
                                autECLPSObj.SendKeys("--------", cc, 28)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 42)
                                apiChk()
                                autECLPSObj.SendKeys("------", cc, 50)
                                apiChk()
                            Else
                                SendEnter()
                                SendF7()
                                RemoveNotCoveredFromEDS6Point5()
                                Exit For
                            End If
                            If cc = 19 Then
                                SendEnter()
                                Dim Myscreen6Point5 As String = GetTextRect(8, 1, 17, 79)
                                SendF8()
                                If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                    SendF3()
                                    Exit For
                                ElseIf GetTextRect(8, 1, 17, 79) = Myscreen6Point5 Then
                                    SendF3()
                                    Exit For
                                Else
                                    cc = 4
                                End If
                                'SendF8()
                                'If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then Exit For
                                'cc = 4
                            End If
                        Next
                        apiChk()
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()

                    End If

                    SendKeyes("A", 1, 75)
                    SendEnter()
                    'apiChk()
                    'apiChk()
                    'apiChk()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                    If GetText(2, 2, 2, True) = "AR" Then
                        SendKeyes("S", 3, 5)
                        SendEnter()
                    End If
                    If GetText(2, 2, 2) <> "MP" Then
                        strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                        ExitAllAndReturnWithException()
                    End If
                End If
LoopAgain:
            Next


            WriteLog("===============EXITING " + GetCurrentMethod.Name + "==============")
        Catch ex As Exception

            WriteLog(GetCurrentMethod.Name + " =>" + ex.Message.ToString)
            strResearchComments = "EXCEPTION=>" + ex.Message.ToString
        End Try
    End Sub

    Sub RemoveNotCoveredFromEDS6Point5()
        ''My Code'''''''''''''''''''''''''''''''''''''''''''''''''''
        WriteLog("Inside " + GetCurrentMethod.Name())
        Try
            Dim SafetyPageCount As Integer = 0
            For KK As Integer = 6 To 18 Step 3

                If GetText(KK, 11, 6, True) = "" Then
                    Exit For
                End If
                SendKeyes("----------------------", KK + 1, 6)
                SendEnter()
                SendKeyes("opn", KK, 2)
                SendEnter()
                apiChk()

                For nn As Integer = 8 To 17 Step 3
                    If Trim(GetText(nn, 7, 4)) = String.Empty Then
                        WriteLog(GetTextRect)
                        SendEnter()
                        Exit For
                    End If
                    SendKeyes("------", nn, 57)
                    SendKeyes("01", nn, 44)
                    SendKeyes("--", nn, 64)
                    SendKeyes("--", nn, 67)
                    SendKeyes("------", nn, 71)
                    SendKeyes("--", nn, 78)
                    SendKeyes("-", nn + 1, 61)
                    SendKeyes("------", nn + 1, 67)
                    SendKeyes("--", nn + 1, 74)
                    If nn = 17 Then
                        WriteLog(GetTextRect)
                        Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "ENTER UNITS ON REPEATED LINE") > 0 Then
                            strResearchComments = "ENTER UNITS ON REPEATED LINE found in EDS 6.5"
                            Exit Sub
                        End If
                        If InStr(GetText(24, 1, 80, True), "Cannot change pricing until 3C is split".ToUpper) > 0 Then
                            strResearchComments = "Cannot change pricing until 3C is split found in EDS 6.5"
                            Exit Sub
                        End If
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                SendF3()
                                Exit For
                            ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                                SendF3()
                                Exit For
                            Else
                                nn = 5
                            End If
                        End If
                Next

                SendKeyes("Y", 2, 49)
                SendEnter()
                SendF3()
                SendKeyes("--------", KK, 59)
                autECLPSObj.SendKeys("[ENTER]")
                If KK = 18 Then
                    SendEnter()
                    If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                        SendKeyes("N", 4, 79)
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                    Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                        Exit For
                    ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                        Exit For
                    Else
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                        KK = 3
                    End If
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
        End Try
    End Sub
    Function initialOperationsToBeDoneWhenPaymentScreenIsPulled(ByVal CurrentSuffix As Integer) As String
        WriteLog(">>============Entering " + GetCurrentMethod.Name + "==============<<")
        RemovePZ1(CurrentSuffix)
        'PullLNLYinPaymentScreen()
        'GetClaimCountFromStruct(HistoryClaim)

        RemoveDenial()
        SendOverrideInScreen("01")
        RemoveDenialWithoutRemarkCode()
        SendOverrideInScreen("01")
        PutCopayNBR()
        RemoveDenial()
        RemoveHxDraftAndICNFromHalfScreen()
        'CheckBypassHistoryField()
        If GetText(2, 2, 2) = "MP" Then
            SendKeyes("EDS", 2, 2)
            SendEnter()
        End If
        RemoveNotCoveredFromEDS6Point5()
        If GetText(3, 2, 5) = "EDS 6" Then
            SendKeyes("A", 1, 75)
            apiChk()
            SendEnter()
        End If
        WriteLog("After RemoveNotCoveredFromEDS6Point5 inside initialOperations method")
        PutCopayNBR()
        'RemoveRPEDecisionDeclinedEdit()
        Dim EditWarning = ""
        Dim Count2 = 0
        Do
            SendEnter()
            If Count2 > 50 Then Exit Do
            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
            EditWarning = EditWarning & GetText(24, 1, 80, True)
            Count2 = Count2 + 1
        Loop
        'If any edit comes in payment screen before this point, the ICN must be skipped

        If InStr(EditWarning, "E1454") > 0 Then
            RemoveRPEDecisionDeclinedEdit()
            EditWarning = ""
            Count2 = 0
            Do
                SendEnter()
                If Count2 > 50 Then Exit Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                EditWarning = EditWarning & GetText(24, 1, 80, True)
                Count2 = Count2 + 1
            Loop
        End If
        If InStr(EditWarning, "E1519") > 0 OrElse InStr(EditWarning, "E1049") > 0 OrElse InStr(EditWarning, "E2884") > 0 Then
            For IJ As Integer = 10 To 22 Step 2
                If InStr(InvalidProcCode, GetText(IJ, 4, 5)) = 0 Then SendKeyes("D1", IJ, 36)
            Next
            SendEnter()
            EditWarning = ""
            Count2 = 0
            EditWarning = GetText(24, 1, 80, True)
            Do
                SendEnter()
                If Count2 > 50 Then Exit Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                EditWarning = EditWarning & GetText(24, 1, 80, True)
                Count2 = Count2 + 1
            Loop
        End If
        If InStr(EditWarning, "WE2560") > 0 AndAlso (InStr(EditWarning, "E2497") = 0 And InStr(EditWarning, "W512") = 0) Then
            SendOverrideInScreen("09")
            EditWarning = ""
            Count2 = 0
            Do
                SendEnter()
                If Count2 > 50 Then Exit Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                EditWarning = EditWarning & GetText(24, 1, 80, True)
                Count2 = Count2 + 1
            Loop
        End If
        WriteLog(">>============Exit " + GetCurrentMethod.Name + "==============<<")
        WriteLog("Returning EditWarning =" + EditWarning)
        Return EditWarning

    End Function
    Public Sub RemoveRPEDecisionDeclinedEdit(Optional ByVal Flag As Boolean = False)
        Try
            Dim Count As Integer = 0, pagecount As Integer = 1
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If
            For I As Integer = 6 To 18 Step 3 'Infinite Loop exception is handled
                If GetText(I, 1, 80, True) = "" Then Exit For

                If InStr(InvalidProcCode, GetText(I, 11, 5)) = 0 Then
                    If GetText(I, 38, 2) <> "30" Then
                        SendKeyes("E9", I, 69)
                        If GetText(I, 69, 2) = "E9" Then
                            SendEnter()
                            'Exit For
                        End If
                    Else
                        SendKeyes("BLANK", I, 11)
                        SendEnter()
                    End If

                End If
                If I = 18 Then
                    If pagecount > 2 Then Exit For
                    pagecount += 1
                    SendF8()
                    I = 3
                End If
            Next
            pagecount = 1
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                ExitAllAndReturnWithException()
            End If
            If Flag Then Exit Sub
            If IsDolClosure() Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                For I As Integer = 6 To 18 Step 3
                    If GetText(I, 1, 80, True) = "" Then Exit For
                    If InStr(InvalidProcCode, GetText(I, 11, 5)) = 0 Then
                        SendKeyes("01", I, 38)
                        SendKeyes("---", I, 45)
                        SendKeyes("--", I, 69)
                        SendKeyes("--", I, 72)
                        SendKeyes("--", I, 75)
                        SendKeyes("--", I, 78)
                    End If
                    If I = 18 Then
                        SendEnter()
                        If pagecount > 2 Then Exit For
                        pagecount += 1
                        SendF8()
                        I = 3
                    End If
                Next
                SendEnter()
            Else
                Count = GetCountOfVaildLineInPaymentScreen()
                If Count >= 7 Then
                    flagExit = True
                    strResearchComments = "7 Lines in payment screen, unable to insert after for removing E1454RPE DECISION DCLINED"
                    strResearchStatus = "Skipped"
                    Dim a As Integer = 0
                    Dim b As Integer = 5 / a
                Else
                    SendKeyes("EDS", 2, 2)
                    SendEnter()
                    For I As Integer = 6 To 18 Step 3
                        If ((I - 6) / 3) + 1 = Count Then
                            SendKeyes("IA", I, 2)
                            SendEnter()
                            If I = 18 Then
                                SendF8()
                                SendKeyes("E9", 6, 69)
                                SendEnter()
                            Else
                                I += 3
                                SendKeyes("E9", I, 69)
                                SendEnter()
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                ExitAllAndReturnWithException()
            End If
            If InStr(GetEditFromPaymentScreen, "E1098") > 0 Then
                SendKeyes(GetTotalChargeFromPaymentScreen().Split(".")(0), 6, 4)
                SendKeyes(GetTotalChargeFromPaymentScreen().Split(".")(1), 6, 11)
                SendEnter()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function GetTotalChargeFromPaymentScreen() As String
        ' WriteLog("------------ inside " + GetCurrentMethod().Name + " ------------")
        Dim TotalCharge As Decimal = 0
        'If InStr(GetEditFromPaymentScreen(), "E1098TOTAL CHARGE REQ") > 0 Then
        For J As Integer = 10 To 22 Step 2
            If InStr(InvalidProcCode, GetText(J, 4, 5)) = 0 AndAlso GetText(J, 29, 2) <> "30" Then
                TotalCharge += CDec(GetText(J, 41, 5).Replace("-", "0").Replace(" ", "0") + "." + GetText(J, 47, 2).Replace("-", "0").Replace(" ", "0"))
            End If
        Next
        'End If
        Dim Charge As String = ""
        Charge = CStr(TotalCharge)
        If InStr(Charge, ".") = 0 Then Charge += ".00"
        Charge = ConvertToFormat(Split(Charge, ".")(0), 6, True, "0") + "." + ConvertToFormat(Split(Charge, ".")(1), 2, , "0")
        Return Charge
    End Function
    Public Function GetCountOfVaildLineInPaymentScreen() As Integer
        Dim Count As Integer = 0
        For I As Integer = 10 To 22 Step 2
            If InStr(InvalidProcCode, GetText(I, 4, 5)) = 0 Then
                Count += 1
            End If
        Next
        Return Count
    End Function
    Public Function IsDolClosure() As Boolean
        If GetText(2, 2, 2) = "MP" Then
            For I As Integer = 10 To 22 Step 2
                If GetText(I, 36, 2) = "7Y" AndAlso GetText(I, 29, 1) = "C" Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function
    Public Sub RemoveHxDraftAndICNFromHalfScreen()
        WriteLog(">>>>>>>>>>>>>>>>>Entering " + GetCurrentMethod.Name)
        Try
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
                apiChk()
            End If
            Dim tempData As String = GetText(4, 15, 27)
            SendKeyes("----------", 4, 15)
            SendKeyes("----------", 4, 32)
            SendEnter()
            apiChk()
            apiChk()
            If InStr(Trim(GetText(24, 2, 80)), "ORG ICN AND DRAFT SHOULD NOT BLANK FOR CORRECTED") > 0 Then
                Dim origIcn As String = Left(tempData, 10)
                Dim drft As String = Right(tempData, 10)
                SendKeyes(origIcn, 4, 15)
                SendKeyes(drft, 4, 32)
                SendEnter()
                apiChk()
            End If

            SendKeyes("N", 4, 79)
            SendEnter()
            SendKeyes("A", 1, 75)
            autECLPSObj.SendKeys("[ENTER]")
            'SendEnter()
            apiChk()
            apiChk()
            If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                WriteLog(GetTextRect)
                For cc As Integer = 7 To 19 Step 3
                    If GetText(cc - 1, 8, 9) <> "         " Then
                        autECLPSObj.SendKeys("--------", cc, 6)
                        apiChk()
                        autECLPSObj.SendKeys("---", cc, 20)
                        apiChk()
                        autECLPSObj.SendKeys("--------", cc, 28)
                        apiChk()
                        autECLPSObj.SendKeys("---", cc, 42)
                        apiChk()
                        autECLPSObj.SendKeys("------", cc, 50)
                        apiChk()
                    End If
                Next
                'SendEnter()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
            End If
            If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                RemoveNotCoveredFromEDS6Point5()
                SendKeyes("N", 4, 79)
                SendEnter()
                apiChk()
                SendKeyes("A", 1, 75)
                'SendEnter()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()

                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                    For cc As Integer = 7 To 19 Step 3
                        If GetText(cc - 1, 8, 9) <> "         " Then
                            autECLPSObj.SendKeys("--------", cc, 6)
                            apiChk()
                            autECLPSObj.SendKeys("---", cc, 20)
                            apiChk()
                            autECLPSObj.SendKeys("--------", cc, 28)
                            apiChk()
                            autECLPSObj.SendKeys("---", cc, 42)
                            apiChk()
                            autECLPSObj.SendKeys("------", cc, 50)
                            apiChk()
                        End If
                    Next
                    'SendEnter()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                End If
                If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                    SendKeyesFor01(OriginalDraft, 4, 32)
                    SendKeyesFor01(OriginalICN, 4, 15)
                    SendEnter()
                    apiChk()
                    SendKeyes("A", 1, 75)
                    'SendEnter()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    If GetText(2, 2, 2) = "MP" Then
                        autECLPSObj.SendKeys("EDS", 2, 2)
                        'SendEnter()
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                End If
            End If
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                ExitAllAndReturnWithException()
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        WriteLog(">>>>>>>>>>>>>>>>>Exit " + GetCurrentMethod.Name)
    End Sub
    Public Function CheckIfMultipleCopayLineWithNoNBR(ByVal CurrentSuffix As Integer) As Boolean
        Try
            Dim Count As Integer = 0
            RemovePZ1(CurrentSuffix)

            For I As Integer = 10 To 22 Step 2
                If GetText(I, 4, 6, True) = "NCOPAY" AndAlso IsNumeric(GetText(I, 25, 3, True).Replace("-", "0")) AndAlso CInt(GetText(I, 25, 3, True).Replace("-", "0")) = 0 Then
                    If Count > 0 Then Return True
                    Count += 1
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return False
    End Function
    Public Function checkForPendPMCMADCV() As Boolean
        Dim TempStr As String = ""
        Try
            If GetText(2, 2, 2) <> "MP" Then Return False
            If InStr(GetText(2, 1, 80), "PZ1") = 0 Then
                Do
                    TempStr = TempStr & GetText(24, 1, 80)
                    SendEnter()
                    'MsgBox TempStr
                Loop Until (InStr(TempStr, GetText(24, 1, 80)) > 0 And TempStr <> "") Or InStr(GetText(2, 1, 80), "PZ1") > 0
            Else
                SendOverrideInScreen("01")
                Do
                    TempStr = TempStr & GetText(24, 1, 80)
                    SendEnter()
                    'MsgBox TempStr
                Loop Until (InStr(TempStr, GetText(24, 1, 80)) > 0 And TempStr <> "") Or InStr(GetText(2, 1, 80), "PZ1") > 0
            End If
            If InStr(TempStr, "E1454") > 0 Then RemoveRPEDecisionDeclinedEdit()
            RemoveBlankAndPISLLine()
            'ICESCheck(CurrICN)
            If strResearchComments = "EDS6.5: ENTER UNITS ON REPEATED LINE" Then
                strResearchStatus = "SKIP"
                Return False
                Exit Function
            End If
            For I As Integer = 10 To 22 Step 2
                'If GetText(I, 1, 80, True).Replace("-", "") = "" Then Exit For
                If GetText(I, 29, 1) = "P" AndAlso InStr("PM,CM,AD,CV", GetText(I, 36, 2)) > 0 Then
                    Return True
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return False
    End Function
    Public Function CheckWhetherAllEDS1LinesPulledInPaymentScreen(ByRef EDS1() As ModifierList, ByVal Suffix As Integer) As Boolean
        Try
            Dim FlagMatch As Boolean = False
            Dim LineCount As String = ""
            'WriteLog("====================== inside Check Whether All EDS1 Lines Pulled In Payment Screen ======================")
CheckAgain:
            Dim arr() As String = Split(GetText(2, 2, 78, True), ",")
            Dim charge As String
            If Right(arr(6), 10) <> CurrICN Then Return True
            For K As Integer = LBound(EDS1) To UBound(EDS1)
                'WriteLog("K = " + CStr(K))
                If CInt(EDS1(K).Suffix) = Suffix Then
                    FlagMatch = False
                    'WriteLog("Matching Suffix " & EDS1(K).Suffix)
                    For J As Integer = 10 To 22 Step 2
                        'WriteLog("J = " + CStr(J))
                        'WriteLog(GetText(J, 4, 5) + " = " + EDS1(K).ProcedureCode)
                        charge = CStr(CDbl(GetText(J, 41, 5).Replace("-", "0").Replace(" ", "0") + "." + GetText(J, 47, 2).Replace("-", "0").Replace(" ", "0")))
                        charge = IIf(CDbl(charge) = 0, "0.01", charge)
                        'WriteLog(charge + " = " + CStr(IIf(CDbl(EDS1(K).Charge) = 0, 0.01, CDbl(EDS1(K).Charge))))
                        If GetText(J, 4, 5) = EDS1(K).ProcedureCode AndAlso CDbl(charge) = IIf(CDbl(EDS1(K).Charge) = 0, 0.01, CDbl(EDS1(K).Charge)) Then
                            If InStr(LineCount, CStr(J)) = 0 Then
                                LineCount += CStr(J) + ","
                                FlagMatch = True
                                Exit For
                            End If
                        End If
                    Next
                    If FlagMatch = False Then
                        ' WriteLog("====================== outside Check Whether All EDS1 Lines Pulled In Payment Screen ======================")
                        Return False
                    End If
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        'WriteLog("====================== outside Check Whether All EDS1 Lines Pulled In Payment Screen ======================")
        'If FlagUndo = False Then
        '    If isEDSUpdateY(CurrICN) Then
        '        FlagUndo = True
        '        PerFormUndo(CurrICN)
        '        SendClear()
        '        sendcl
        '        GoTo CheckAgain
        '    End If
        'End If
        Return True
    End Function
    Public Sub EDS6Point5ValueMatch(ByRef C() As PreviousValueCheck)
        Try
            ReDim C(0)
            Dim LineCount As Integer = 0
            Dim I As Integer = 0
            For dd As Integer = 8 To 17 Step 3
                If Trim(GetText(dd, 7, 4)) = String.Empty Then Exit For

                If I = 0 Then
                    ReDim C(LineCount).RevCode(I)
                    ReDim C(LineCount).DOS(I)
                    ReDim C(LineCount).Charge(I)
                    ReDim C(LineCount).OV(I)
                    ReDim C(LineCount).Units(I)
                Else
                    ReDim Preserve C(LineCount).RevCode(I)
                    ReDim Preserve C(LineCount).DOS(I)
                    ReDim Preserve C(LineCount).Charge(I)
                    ReDim Preserve C(LineCount).OV(I)
                    ReDim Preserve C(LineCount).Units(I)
                End If
                C(LineCount).RevCode(I) = Trim(GetText(dd, 7, 4))
                C(LineCount).DOS(I) = Trim(GetText(dd, 33, 7))
                C(LineCount).Charge(I) = Trim(GetText(dd, 47, 6)) + "." + Trim(GetText(dd, 54, 2))
                C(LineCount).OV(I) = Trim(GetText(dd, 44, 2))
                C(LineCount).Units(I) = Trim(GetText(dd, 38, 5))
                I += 1

            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function EDS6Point5IndexMatch() As Boolean
        Dim checkFlag As Boolean = True
        Dim Index As Integer = 0
        Dim count As Integer = 0
        Dim Ncount As Integer = 0
        Try
            For dd As Integer = LBound(PreviosValue(0).RevCode) To UBound(PreviosValue(0).RevCode)
                If PreviosValue(0).RevCode(dd) = CurrentValue(0).RevCode(dd) Then
                Else
                    checkFlag = False
                    Exit For
                End If
                If PreviosValue(0).Charge(dd) = CurrentValue(0).Charge(dd) Then
                Else
                    checkFlag = False
                    Exit For
                End If
                If PreviosValue(0).Units(dd) = CurrentValue(0).Units(dd) Then
                Else
                    checkFlag = False
                    Exit For
                End If
                If PreviosValue(0).OV(dd) = CurrentValue(0).OV(dd) Then
                Else
                    checkFlag = False
                    Exit For
                End If
                If PreviosValue(0).DOS(dd) = CurrentValue(0).DOS(dd) Then
                Else
                    checkFlag = False
                    Exit For
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try

        Return checkFlag

    End Function
    Public Sub CurrentPaymentScreen(ByRef C() As CurrentICNPaymentScreen, SuffixCount As Integer)
        Try
            ReDim C(SuffixCount - 1)
            Dim Linecount As Integer
            Dim count As Integer = 1
            For I As Integer = 1 To SuffixCount
                SendClear()
                SendClear()
                pullEDSScreen(CurrentClaimDetails.ICN, 1)
                SendKeyes("A" & I, 1, 76)
                SendEnter()
                apiChk()
                Linecount = 0
                count = 1
                For dd As Integer = 10 To 22 Step 2
                    If Trim(GetText(dd, 1, 2)) = "--" Then Exit For

                    If Linecount = 0 Then

                        ReDim C(I - 1).POS(Linecount)
                        ReDim C(I - 1).SVC(Linecount)
                        ReDim C(I - 1).FDate(Linecount)
                        ReDim C(I - 1).LDAte(Linecount)
                        ReDim C(I - 1).NBR(Linecount)
                        ReDim C(I - 1).OV(Linecount)
                        ReDim C(I - 1).RC(Linecount)
                        ReDim C(I - 1).Charge(Linecount)

                    Else

                        ReDim Preserve C(I - 1).POS(Linecount)
                        ReDim Preserve C(I - 1).SVC(Linecount)
                        ReDim Preserve C(I - 1).FDate(Linecount)
                        ReDim Preserve C(I - 1).LDAte(Linecount)
                        ReDim Preserve C(I - 1).NBR(Linecount)
                        ReDim Preserve C(I - 1).OV(Linecount)
                        ReDim Preserve C(I - 1).RC(Linecount)
                        ReDim Preserve C(I - 1).Charge(Linecount)

                    End If

                    C(I - 1).POS(Linecount) = Trim(GetText(dd, 1, 2))
                    C(I - 1).SVC(Linecount) = Trim(GetText(dd, 4, 6))
                    C(I - 1).FDate(Linecount) = Trim(GetText(dd, 11, 6))
                    C(I - 1).LDAte(Linecount) = Trim(GetText(dd, 18, 6))
                    C(I - 1).NBR(Linecount) = Trim(GetText(dd, 25, 3))
                    C(I - 1).OV(Linecount) = Trim(GetText(dd, 29, 2))
                    C(I - 1).RC(Linecount) = Trim(GetText(dd, 36, 2))
                    C(I - 1).Charge(Linecount) = Trim(GetText(dd, 41, 5)) + "." + Trim(GetText(dd, 47, 2))
                    'C(I - 1).LineCount(Linecount) = count

                    Linecount += 1
                    count += 1
                Next
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function GetTotalPaidAmount(ByVal ICN As String) As Double
        SendClear()
        SendClear()
        'SendClear()
        PullMHI(ICN)
        If strResearchStatus = "W703DOC PROCESSED" Then
            Return 0
        ElseIf strResearchStatus = "Unable to pull MHI" Then
            Return 0
        End If
        If GetText(2, 2, 2) <> "MH" Then Return "0"
        Dim myClaim(100) As Claim
        'PullMHI(ICN)
        Dim J As Integer = ReadMHI(myClaim)
        Dim I As Integer = J - 1
        Dim draftToBeConsidered As String = ""
        Dim TotalPaid As Double = 0
        Dim FlagROV As Boolean = False
        While I >= 0
            If myClaim(I).Voided = True Then
                If myClaim(I).RemarkCode(0) = "69" Then
                    'No need to change
                ElseIf myClaim(I).RemarkCode(0) = "74" Then
                    If InStr(draftToBeConsidered, myClaim(I).Draft + ",") > 0 Then
                        draftToBeConsidered = Replace(draftToBeConsidered, myClaim(I).Draft + ",", "")
                    ElseIf InStr(draftToBeConsidered, myClaim(I).Draft) > 0 Then
                        draftToBeConsidered = Replace(draftToBeConsidered, myClaim(I).Draft, "")
                    End If
                ElseIf myClaim(I).RemarkCode(0) = "71" Then
                    If myClaim(I).OV(0).Trim = "R" Then
                        'No need to change
                    Else
                        If InStr(draftToBeConsidered, myClaim(I).Draft + ",") > 0 Then
                            draftToBeConsidered = Replace(draftToBeConsidered, myClaim(I).Draft + ",", "")
                        ElseIf InStr(draftToBeConsidered, myClaim(I).Draft) > 0 Then
                            draftToBeConsidered = Replace(draftToBeConsidered, myClaim(I).Draft, "")
                        End If
                    End If
                End If
            Else
                If myClaim(I).RemarkCode(0) = "77" Or myClaim(I).RemarkCode(0) = "75" Then
                    FlagROV = True
                    'draftToBeConsidered = Replace(draftToBeConsidered, myClaim(I).Draft, "")
                Else
                    If CDbl(myClaim(I).Draft) <> 0 AndAlso InStr(draftToBeConsidered, myClaim(I).Draft) = 0 Then draftToBeConsidered += myClaim(I).Draft + ","
                End If

            End If
            I -= 1
        End While

        For Each Draft In draftToBeConsidered.Split(",")
            If Draft.Trim <> "" Then
                I = J - 1
                While I >= 0
                    If myClaim(I).Draft = Draft Then
                        'If isROVInDraft(myClaim, J, Draft) Then
                        '    If FlagROV Then
                        '        TotalPaid += CDbl(myClaim(I).PaidAmount) - GetTotalROVAmountInDraft(myClaim, J, Draft)
                        '    Else
                        '        TotalPaid += CDbl(myClaim(I).PaidAmount)
                        '    End If
                        'Else
                        '    TotalPaid += CDbl(myClaim(I).PaidAmount)
                        'End If
                        If IsDraftAlreadyVoided(myClaim, Draft) Then 'CDbl(myClaim(I).PaidAmount) > 0 Then '7415117241
                            TotalPaid += CDbl(myClaim(I).PaidAmount)

                            'For k As Integer = LBound(myClaim(I).ServiceLine) To UBound(myClaim(I).ServiceLine) - 1
                            '    Dim ProcCode As String = myClaim(I).ProcCode(k)
                            '    Dim Fdos As String = myClaim(I).FDOS(k)
                            '    Dim ldos As String = myClaim(I).LDOS(k)
                            '    If IsProcAndDosPresentInHx(ICNList, ProcCode, Fdos, ldos) Then
                            '        TotalPaid += CDbl(myClaim(I).Paid(k))
                            '    End If
                            'Next
                        End If
                        Exit While
                    End If
                    I -= 1
                End While
            End If
        Next
        Return TotalPaid
    End Function

    Public Function IsDraftAlreadyVoided(ByRef MyClaim() As Claim, ByVal Draft As String) As Boolean

        For I As Integer = 0 To GetClaimCountFromStruct(MyClaim) - 1
            If MyClaim(I).Draft = Draft AndAlso MyClaim(I).Voided = True Then
                If MyClaim(I).RemarkCode(0) = "69" Then Return True
                If MyClaim(I).RemarkCode(0) = "74" Then Return False
            End If
        Next
        Return True

    End Function
    Public Function AllHistoryPerform(ByVal ICNList As String) As String
        Dim iCN As String = String.Empty
        Dim PaidAmount As String = String.Empty
        Dim claimDetarmination As String = String.Empty
        Dim sb As New StringBuilder
        Dim iCN1 As String = String.Empty
        Dim ICNCount1 As Integer = 0
        Dim IsMHIRead As Integer
        Dim flagOV As Boolean = False
        ReDim ListOfPossibleICNActual(0)
        Try
            Dim HICNArray() As String = ICNList.Split(",")
            For mm As Integer = LBound(HICNArray) To UBound(HICNArray)
                strResearchStatus = ""
                iCN1 = HICNArray(mm)
                If iCN1.Trim <> "" Then
                    PullMHI(IIf(iCN1.Length > 10, iCN1, iCN1))
                    If strResearchStatus = "W703DOC PROCESSED" Then
                        Continue For
                    End If
                    If strResearchStatus = "Unable to pull MHI" Then
                        strResearchStatus = "Hist ICN: Unable to pull MHI of " + iCN1
                        Continue For
                    End If
                    IsMHIRead = ReadMHI(HistoryClaim)
                    For J As Integer = LBound(HistoryClaim) To UBound(HistoryClaim)        'Added this check to skip claims with Hx voided by other ID  @ 20/08/25
                        If HistoryClaim(J).AdjusterID Is Nothing Then Exit For
                        Dim is69 As Boolean = False
                        Dim is70 As Boolean = False
                        For K As Integer = LBound(HistoryClaim(J).RemarkCode) To UBound(HistoryClaim(J).RemarkCode)
                            If HistoryClaim(J).RemarkCode(K) = "69" Then
                                is69 = True
                            End If
                        Next
                        For K As Integer = LBound(HistoryClaim(J).RemarkCode) To UBound(HistoryClaim(J).RemarkCode)
                            If HistoryClaim(J).RemarkCode(K) = "70" Then
                                is70 = True
                            End If
                        Next
                        If is69 And is70 Then
                            If HistoryClaim(J).AdjusterID <> "398676" And HistoryClaim(J).AdjusterID <> "430664" Then
                                strResearchComments = "History voided by different Adjuster"
                                Return strResearchComments
                            End If
                        End If
                    Next

                    If IsMHIRead = -1 Then
                        Continue For
                    End If
                    If ICNCount1 = 0 Then

                        CopyPossibleIcn(ListOfPossibleICNActual(ICNCount1), IIf(iCN1.Length > 10, iCN1, iCN1))
                        ICNCount1 += 1

                    Else

                        ReDim Preserve ListOfPossibleICNActual(ICNCount1)
                        CopyPossibleIcn(ListOfPossibleICNActual(ICNCount1), IIf(iCN1.Length > 10, iCN1, iCN1))
                        ICNCount1 += 1

                    End If

                End If
            Next
            strResearchStatus = ""
            strResearchComments = ""
            For II As Integer = LBound(HICNArray) To UBound(HICNArray)
                Dim FalgError As Boolean = False
                Dim IsRCCA As Integer
                iCN = HICNArray(II)

                If iCN.Trim <> "" Then
                    PaidAmount = GetTotalPaidAmount(iCN).ToString()
                    If strResearchStatus = "W703DOC PROCESSED" AndAlso PaidAmount = 0 Then
                        Continue For
                    ElseIf strResearchStatus = "Unable to pull MHI" AndAlso PaidAmount = 0 Then
                        Continue For
                    End If
                    WriteLog("Paid amt for HxICN:" + iCN + " is= " + PaidAmount)

                    claimDetarmination = ClaimDetermination(PaidAmount)
                    PaymentComment = claimDetarmination
                    If PaymentComment = "Over Payment" And (billType = "7" Or billType = "1") Then
                        If SuffixCountOfCurrentICN > 1 Then
                            strResearchComments = "Overpayment in multiple suffix"
                            Return strResearchComments
                        End If

                        Dim ubOV As UBOverpayment = New UBOverpayment(Office, ICNList, CurrICN)
                        flagOV = ubOV.UBOverpaymentProcessFlow(True)
                        If strResearchStatus = "Overpayment adjustment performed" And flagOV = True Then
                            Return strResearchComments
                        Else
                            Return strResearchComments
                        End If
                    ElseIf PaymentComment = "Over Payment" Then
                        strResearchComments = "Claim is Over Payment"
                        Return strResearchComments
                    End If
                    WriteLog("ICN: " + iCN + " determined as " + PaymentComment)
                    If DetermineIfCAApplicable(Map, CurrentClaimPaymentScreen, IsRCCA, HistoryClaim(0).CauseCode, IIf(iCN.Length > 10, iCN, iCN)) Then
                        strResearchComments = "Need to changes:CA applicable"
                        'Exit For
                    Else
                        VoidHistoryClaim(FalgError, iCN, claimDetarmination, PaidAmount)
                        WriteLog($"History ICN voided: {isHistoryICNVoided}")
                        If FalgError Then
                            Return strResearchComments
                        End If
                        If InStr(strResearchComments, "Index out of range") > 0 Then
                            Return strResearchComments
                            Exit Function
                        ElseIf InStr(strResearchComments, "E1007SUMMRY CHK NOT ISSD") > 0 Then
                            Return strResearchComments
                            Exit Function
                        End If
                        If strResearchComments <> "claim is over payment" And strResearchComments <> "Dxcode not match" OrElse strResearchComments <> "Skip" Then
                            If sb.ToString = "" Then
                                sb.Append(PaidAmount)
                            Else
                                sb.Append("," + PaidAmount)
                            End If
                        End If
                    End If
                End If

            Next
            If strResearchComments = "claim is over payment" OrElse strResearchComments = "Skip" OrElse strResearchComments = "Dxcode not match" OrElse
                InStr(strResearchComments, "Need to changes") > 0 Then
                Return strResearchComments
            End If
            Dim ICNCount As Integer = 0
            ReDim ListOfPossibleICNActual(0)
            For mm As Integer = LBound(HICNArray) To UBound(HICNArray)
                strResearchStatus = ""
                iCN1 = HICNArray(mm)
                If iCN1.Trim <> "" Then
                    PullMHI(IIf(iCN1.Length > 10, iCN1, iCN1))
                    If strResearchStatus = "W703DOC PROCESSED" Then
                        Continue For
                    ElseIf strResearchStatus = "Unable to pull MHI" Then
                        Continue For
                    End If
                    ReadMHI(HistoryClaim)
                    If ICNCount = 0 Then

                        CopyPossibleIcn(ListOfPossibleICNActual(ICNCount), IIf(iCN1.Length > 10, iCN1, iCN1))
                        ICNCount += 1

                    Else

                        ReDim Preserve ListOfPossibleICNActual(ICNCount)
                        CopyPossibleIcn(ListOfPossibleICNActual(ICNCount), IIf(iCN1.Length > 10, iCN1, iCN1))
                        ICNCount += 1

                    End If
                    Try
                        For I As Integer = LBound(HistoryClaim) To UBound(HistoryClaim)
                            If HistoryClaim(I).Draft Is Nothing Then Exit For
                            For J As Integer = LBound(HistoryClaim(I).RemarkCode) To UBound(HistoryClaim(I).RemarkCode)
                                If HistoryClaim(I).RemarkCode(J) Is Nothing Then Continue For
                                If HistoryClaim(I).RemarkCode(J) = "OV" Then
                                    If HistoryClaim(I).OV(J) = "R" Then

                                        If HistoryClaim(I).AdjusterID <> "398676" And HistoryClaim(I).AdjusterID <> "430664" Then
                                            strResearchComments = "ROV done by other ID"
                                            Return strResearchComments
                                        End If
                                    End If
                                End If
                            Next
                        Next

                    Catch ex As Exception
                        WriteLog($"Error from {GetCurrentMethod.Name} while checkng ROV done by any other adjuster")
                    End Try

                End If
            Next
            If IsOIOIM(ListOfPossibleICNActual, CurrentClaimPaymentScreen) Then
                strResearchComments = "OI/OIM PRESENT"
                Return strResearchComments
            End If

            HistoryPaidAmount = sb.ToString
            WriteLog("History paid amount: " + HistoryPaidAmount)
            If HistoryPaidAmount = Nothing Then
                strResearchComments = "Paid amount is Nothing"
                Return strResearchComments
            End If
            Dim ArrayPaidAmount() As String = HistoryPaidAmount.Split(",")
            Dim TotalAmount As Decimal = 0.0
            Dim DifferenceAmount As Decimal = 0.0
            Dim RemarkCodeFor30Line As String = String.Empty
            Dim ThirtyLineAmount As String = String.Empty
            Dim FlagCARemarkCode As Boolean = False
            For dd As Integer = LBound(ArrayPaidAmount) To UBound(ArrayPaidAmount)
                Dim Amount As String = ArrayPaidAmount(dd)
                If Amount = "" Then
                    Amount = "0.00"
                End If
                TotalAmount = CDec(Amount) + TotalAmount
            Next
            If isROVand75or77Present(ICNList) Then
                WriteLog("ROV Amount=" + ROVAmount.ToString)
                TotalAmount = TotalAmount - ROVAmount
                WriteLog("Total Paid Amount in History= " + TotalAmount.ToString)
            End If
            If strResearchComments = "ROV found in HxICN, please ignore ROV" Then
                Return strResearchComments
            End If
            WriteLog("Current Icn PAYMENT=" + currentIcnPaymentAmount.ToString())
            If currentIcnPaymentAmount <> 0.00 Then
                DifferenceAmount = currentIcnPaymentAmount - TotalAmount
            Else
                DifferenceAmount = TotalAmount
            End If
            WriteLog("DifferenceAmount=" + DifferenceAmount.ToString())
            If DifferenceAmount > 5000 Then IsHDClaim = True 'Added 30th Jan,2023 by Sourav to mark the High $ claims

            If DifferenceAmount > 0 Then
                RemarkCodeFor30Line = "E5"
            ElseIf DifferenceAmount = 0 Then
                'RemarkCodeFor30Line = "CA"
                RemarkCodeFor30Line = "E5"          'Changed 9/13/25 -- All 30line RC must be E5
                flagZeroPayment = True
            End If
            WriteLog("RemarkCodeFor30Line= " + RemarkCodeFor30Line)
            'If isROVand75or77Present() And TotalAmount > 0 Then
            '    ThirtyLineAmount = (TotalAmount - ROVAmount).ToString
            'Else
            '    ThirtyLineAmount = TotalAmount.ToString()
            'End If
            ThirtyLineAmount = TotalAmount.ToString()
            WriteLog("ThirtyLineAmount= " + ThirtyLineAmount)
            Dim FlagPartialProcessed As Boolean = False
            Dim FlagError As Boolean = False
            performAdjustment(CurrentClaimPaymentScreen, Nothing, Nothing, FlagError, FlagPartialProcessed, strDraFT, CurrentClaimDetails, ThirtyLineAmount, flagZeroPayment, FlagCARemarkCode, "", SuffixCountOfCurrentICN)
            If FlagError = True Then
                Return strResearchComments
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name + " exception: " + ex.Message.ToString
        End Try

        Return strResearchComments
    End Function

    Sub create30LineInMPCScreen(ByRef FlagError As Boolean, ByVal ThirtyLineAmount As String, ByVal RemarkCode30Line As String)
        Try
            Dim mpcLine As Integer, Valid30Line As Integer
            Dim Temp30Line1 As String, Temp30Line2 As String, arrTemp30Line2() As String, arrTempStr() As String, arrTempPOS() As String
            Dim arrPOS2() As String, arrTempLen() As String, arrTempLen2() As String, TotalPaidInSuffix As Double = 0
            If InStr(ThirtyLineAmount, ".") = 0 Then ThirtyLineAmount += ".00"
            If CDbl(ThirtyLineAmount) = 0 Then Exit Sub
            FlagError = False
            For mpcLine = 10 To 22 Step 2
                If GetText(mpcLine, 4, 6, True) = "------" Then
                    Exit For
                ElseIf InStr("ACPD,TDPD,EMER,POD,PRE,OI,OIM,OIMEDI,NCOPAY,COPAY,CXINT,NYS,S9999", GetText(mpcLine, 4, 6, True)) = 0 Then
                    Valid30Line = mpcLine
                ElseIf GetText(mpcLine, 4, 6, True) = "NCOPAY" Then
                    SendKeyes("001", mpcLine, 25)
                    SendEnter()
                Else
                    TotalPaidInSuffix += CDbl(GetText(mpcLine, 61, 5, True) + "." + GetText(mpcLine, 67, 2))
                End If
            Next
            If mpcLine = 24 Then
                'FlagError = True
                'strResearchStatus = "SKIPPED"
                'strResearchComments = "Unable to create 30 line in payment as there are 7 lines present."
                ThirtyLineAmount = CStr(CDbl(ThirtyLineAmount) + TotalPaidInSuffix)
                If InStr(ThirtyLineAmount, ".") = 0 Then ThirtyLineAmount += ".00"
                Exit Sub
            End If

            Temp30Line1 = GetText(Valid30Line, 1, 80, True)
            Temp30Line2 = GetText(Valid30Line + 1, 1, 80, True)

            Do
                Temp30Line2 = Replace(Temp30Line2, "  ", " ")
            Loop Until InStr(Temp30Line2, "  ") = 0

            arrTemp30Line2 = Split(Temp30Line2, " ")

            Do
                Temp30Line1 = Replace(Temp30Line1, "  ", " ")
            Loop Until InStr(Temp30Line1, "  ") = 0
            arrTempStr = Split(Temp30Line1, " ")
            arrTempPOS = Split("1,4,11,18,25,29,32,34,36,39,41,47,50,56,59,65,68,74,77", ",")
            arrPOS2 = Split("1,3,7,9,15,18,20,25,30,36,39,41,47,50,53,55,59,61,67,70,72,78", ",")
            arrTempLen = Split("2,5,6,6,3,2,1,1,2,1,5,2,5,2,5,2,5,2,2", ",")
            arrTempLen2 = Split("1,3,1,5,2,1,5,2,1,5,2,1,5,2,2,1,3,1,5,2,1,5,2", ",")
            If CDbl(ThirtyLineAmount) > 0 Then
                For L = LBound(arrTempStr) To UBound(arrTempStr)
                    If IsNumeric(arrTempStr(L)) Then
                        If L = 1 Then
                            SendKeyes(arrTempStr(L), mpcLine, CInt(arrTempPOS(L)))
                        Else
                            SendKeyes(ConvertToFormat(arrTempStr(L), arrTempLen(L), True, "0"), mpcLine, CInt(arrTempPOS(L)))
                        End If

                    Else
                        If L = 1 Then
                            SendKeyes(arrTempStr(L), mpcLine, CInt(arrTempPOS(L)))
                        Else
                            SendKeyes(ConvertToFormat(arrTempStr(L), arrTempLen(L)), mpcLine, CInt(arrTempPOS(L)))
                        End If

                    End If

                Next
                SendKeyes("00000", mpcLine, 50)
                SendKeyes("00", mpcLine, 56)

                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(0), 5, True, "0"), mpcLine, 41)
                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(1), 2, False, "0"), mpcLine, 47)
                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(0), 5, True, "0"), mpcLine + 1, 30)
                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(1), 2, False, "0"), mpcLine + 1, 36)
                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(0), 5, True, "0"), mpcLine + 1, 61)
                SendKeyes(ConvertToFormat(Split(ThirtyLineAmount, ".")(1), 2, False, "0"), mpcLine + 1, 67)
                SendKeyes("30", mpcLine, 29)
                SendKeyes(RemarkCode30Line, mpcLine, 36)
                SendKeyes("100", mpcLine + 1, 55)
            End If
            Dim TotCharge As String = GetText(6, 4, 6) + "." + GetText(6, 11, 2)

            SendEnter()
            If GetColorScheme(6, 4) = 3 Then
                SendKeyes(TotCharge.Split(".")(0), 6, 4)
                SendKeyes(TotCharge.Split(".")(1), 6, 11)
                SendEnter()
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub

    Public Sub HistoryClaimDXCodeFromHardCopy(ByRef UserCtrlFrm As frmMain, ByRef D() As STDXCODE)
        Dim getDoc360 As String
        Dim LineCount As Integer = 0
        Dim strMainLine As String = String.Empty
        Dim arrMainLine() As String
        Dim strHardCopyDXCOde As String = String.Empty
        Dim arrDXCode() As String
        ReDim D(0)
        Dim I As Integer = 0
        UserCtrlFrm.rtfEdss.Text = getDoc360
        For Each strLine As String In UserCtrlFrm.rtfEdss.Lines
            If InStr(strLine, "--- ICD VERSION") > 0 Then
                LineCount = LineCount + 1
                strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))

                While InStr(strMainLine, "----------------") = 0

                    strMainLine = Trim(UserCtrlFrm.rtfEdss.Lines(LineCount))
                    While strMainLine.IndexOf("  ") <> -1
                        strMainLine = strMainLine.Replace("  ", " ")
                    End While

                    If InStr(strMainLine, "CODE:") > 0 Then

                        Erase arrMainLine

                        arrMainLine = strMainLine.Trim.Replace("|", String.Empty).Split(" ")

                        If InStr(Trim(strHardCopyDXCOde), Trim(arrMainLine(3))) = 0 Then
                            If Trim(arrMainLine(3)) <> Trim(strHardCopyDXCOde) Then
                                strHardCopyDXCOde = strHardCopyDXCOde & "," & Trim(arrMainLine(3))
                            End If
                        End If
                    End If
                    LineCount = LineCount + 1
                End While
            End If
        Next
        Erase arrDXCode

        arrDXCode = strHardCopyDXCOde.Split(",")
        For ff As Integer = LBound(arrDXCode) To UBound(arrDXCode) - 1
            If I = 0 Then
                ReDim D(0).DXCode(ff)
            Else
                ReDim Preserve D(0).DXCode(ff)
            End If
            D(0).DXCode(ff) = arrDXCode(ff)
            I += 1
        Next

    End Sub

    Public Sub DxCodeFunctionForHistory(ByRef D() As STDXCODE, ByVal ICN As String, ByRef FlagError As Boolean)
        Try
            WriteLog("Inside " + GetCurrentMethod.Name)
            ReDim D(0)
            SendClear()
            SendClear()
            SendKeyes("mei,i" + ICN + ",1", 2, 2)
            SendEnter()
            apiChk()
            If InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Then

                SendClear()
                SendClear()
                SendKeyes("ret,i" + ICN + ",m", 2, 2)
                SendEnter()
                apiChk()
                SendKeyes("RHI", 2, 2)
                SendEnter()
                apiChk()
                apiChk()
                Dim draftComm As String = ""
                GetHRHI(ICN, HistoryRHI, draftComm)
                If GetText(2, 2, 3) = "RHN" Then
                    SendKeyesFor01("I", 2, 4)
                    SendEnter()
                End If
                If InStr(GetText(24, 1, 80), "NO CLAIMS ON FILE") > 0 Then
                    FlagError = True
                    strResearchStatus = "ERROR in pulling RHI"
                Else
                    WriteLog("RHI FOUND")
                    WriteLog(GetTextRect)
                    Dim I As Integer = 0
                    For ff As Integer = 5 To 79 Step 12
                        Dim checkvalue As String = Trim(GetText(5, ff - 2, 4))
                        If checkvalue = "" Then Exit For
                        If I = 0 Then
                            ReDim D(0).DXCode(I)
                        Else
                            ReDim Preserve D(0).DXCode(I)
                        End If
                        D(0).DXCode(I) = Trim(GetText(5, ff, 6))
                        I += 1
                    Next
                End If
            Else
                SendKeyes("dx", 1, 76)
                SendEnter()
                apiChk()
                WriteLog("DX SCREEN FOUND")
                WriteLog(GetTextRect)
                ReDim D(0)
                Dim LineCount As Integer = 0
                Dim x As Integer = 0
                For dd As Integer = 7 To 11 Step 1
                    For hh As Integer = 6 To 64 Step 13
                        Dim checkLine As String = String.Empty
                        checkLine = Trim(GetText(dd, hh, 7))
                        If checkLine = "-------" Then Exit For
                        If x = 0 Then
                            ReDim D(LineCount).DXCode(x)
                        Else
                            ReDim Preserve D(LineCount).DXCode(x)
                        End If
                        D(LineCount).DXCode(x) = Trim(GetText(dd, hh, 7))
                        x += 1
                    Next
                Next
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
        Finally
            WriteLog("EXIT " + GetCurrentMethod.Name)
        End Try

    End Sub
    Sub performAdjustment(ByRef MyClaim() As ClaimFromPaymentScreen, ByRef myMap() As MapCurrentProcDetailsWithHistory, ByRef MyISI() As ISI, ByRef FlagError As Boolean,
                          ByRef FlagPartialProcessed As Boolean, ByVal Draft As String, ByRef MyClaimDetails As ClaimDetails, ByVal AmountToBeAdjusted As String,
                          ByVal FlagZeroPayment As Boolean, ByVal FlagCARemarkCode As Boolean, ByVal SubroGationRemarkCode As String, SuffixCount As Integer) 'ByRef MyDrafts() As DraftDetails
        Try
            WriteLog($"===================Inside {GetCurrentMethod.Name}====================")
            Dim Editwarning As String = "", Count2 As Integer = 0
            Dim RCLineLevel As String = "", TempFlag As Boolean = False
            Dim CounterRC As Integer, ClaimPaidTotal As Double = 0, ClaimAllowableTotal As Double = 0.0
            Dim MyDraft() As String
            Dim flagOV As Boolean = False
            Dim TotalPaidinCurrent As Double = 0.0, TotalAllowableinCurrent As Double = 0.0
            If Draft IsNot Nothing Then
                MyDraft = Draft.Split(",")
            Else
                MyDraft = New String() {}
            End If
            Dim suffix As Integer = 0
            For I As Integer = LBound(MyClaim) To UBound(MyClaim)
                If MyClaim(I).POS.Length <= 7 Then
                    ClaimPaidTotal = Math.Round((CDbl(MyClaim(I).TotalPaid) + ClaimPaidTotal), 2)
                End If
                For J As Integer = LBound(MyClaim(I).Allowable) To UBound(MyClaim(I).Allowable)
                    If MyClaim(I).POS.Length <= 7 Then
                        ClaimAllowableTotal = Math.Round((CDbl(MyClaim(I).Allowable(J)) + ClaimAllowableTotal), 2)
                    End If
                Next
            Next

            WriteLog("Current claim total paid=" + ClaimPaidTotal.ToString)
            WriteLog("Current claim total allowable=" + ClaimAllowableTotal.ToString)
            '-----------------------------------------------------
            'This part has been added on 29th Oct,2026 by Sourav to compare values of payment screen in two structures. If not same then, skip the ICN to avoid incorrect processing
            If ICESDenialFlag = True Then
                For l As Integer = LBound(CurrentIcnMPCScreenIfICES) To UBound(CurrentIcnMPCScreenIfICES)
                    For m As Integer = LBound(CurrentIcnMPCScreenIfICES(l).Payment) To UBound(CurrentIcnMPCScreenIfICES(l).Payment)
                        TotalPaidinCurrent = CDbl(CurrentIcnMPCScreenIfICES(l).Payment(m)) + TotalPaidinCurrent
                    Next
                Next
            Else
                For l As Integer = LBound(CurreentICNPaymentScreen) To UBound(CurreentICNPaymentScreen)
                    For m As Integer = LBound(CurreentICNPaymentScreen(l).Payment) To UBound(CurreentICNPaymentScreen(l).Payment)
                        TotalPaidinCurrent = CDbl(CurreentICNPaymentScreen(l).Payment(m)) + TotalPaidinCurrent
                    Next
                Next
            End If
            If currentIcnPaymentAmount <> ClaimPaidTotal Or allowableFromOCI <> ClaimAllowableTotal Then
                strResearchComments = "Data mismatch in allowable and total paid in Current Claim"
                FlagError = True
                Exit Sub
            End If
            '--------------------------------------------------


            'If currentIcnPaymentAmount < CDbl(AmountToBeAdjusted) Then      'Added this and commented the below code for overpayment check
            '    FlagError = True
            '    strResearchStatus = "SKIPPED"
            '    strResearchComments = "Overpayment adjustment to be performed."
            '    Exit Sub
            'End If
            If ClaimPaidTotal < CDbl(AmountToBeAdjusted) And CDbl(allowableFromOCI) = ClaimAllowableTotal Then               'Commented by Sourav on 10/28/25 because claims are skipping as overpayment when there is no data in Myclaim()
                If SuffixCountOfCurrentICN > 1 Then
                    strResearchComments = "Overpayment in multiple suffix"
                    FlagError = True
                    Exit Sub
                End If
                WriteLog("Executing overpayment flow")
                Dim ubOV As UBOverpayment = New UBOverpayment(Office, FinalHistoryICNList, CurrICN)
                flagOV = ubOV.UBOverpaymentProcessFlow(True)

                If strResearchStatus = "Overpayment adjustment performed" And flagOV = True Then
                    Exit Sub
                Else
                    FlagError = True
                    strResearchStatus = "SKIPPED"
                    strResearchComments = "Overpayment adjustment to be performed."
                    Exit Sub
                End If

            End If

            For I As Integer = LBound(MyClaim) To UBound(MyClaim)
                If CDbl(MyClaim(I).Draft) = 0 Then
                    pullEDSScreen(MyClaimDetails.ICN, "9")
                    SendKeyes("A" & CStr(I + 1), 1, 76)
                    SendEnter()
                    WriteLog(GetTextRect)
                    IncreementMPCCount()
                    If InStr(GetText(24, 1, 80, True), "WE1191NO LINES PENDED") > 0 Then
                        For J As Integer = LBound(MyDraft) To UBound(MyDraft)
                            If suffix = 1 Then
                                SendClear()
                                SendClear()
                                SendClear()
                                Dim AdjConLine As String = "ADJ," + CurrentClaim(0).Policy + "," + CurrentClaim(0).SSN + "," +
                                                           CurrentClaim(0).EmployeeName + "," + CurrentClaim(0).Relationship + "," +
                                                           MyDraft(J) + ",I" + CurrentClaim(0).ICN
                                SendKeyes(AdjConLine, 2, 2)
                                SendEnter()
                                Exit For
                            End If
                        Next
                    End If
                    Editwarning = initialOperationsToBeDoneWhenPaymentScreenIsPulled(I + 1)
                    'FollowICESInstructionInPaymentScreen(MyISI, I + 1)
                    If InStr(Editwarning, "E2570") > 0 Then
                        SendOverrideInScreen("18")
                    End If
                    PasteHxICNAndDraft(AmountToBeAdjusted, MyClaim(I).TotalPaid)
                    PasteHistoryICN(CurrICN, FlagError, strResearchStatus, strResearchComments)
                    If flagDenial = True Then
                        implicitDenial(SuffixCount)
                    End If
                    If nonServiceCodePresent = True Then
                        implicitDenialInEDS(SuffixCount)
                    End If
                    'ICESCheck(CurrICN)
                    If ICESDenialFlag Then
                        placingiCESDenialInfo(CurrICN, SuffixCountOfCurrentICN)
                    End If
                    If icesSuccess = False And ICESDenialFlag = True Then
                        strResearchComments = "ICES denial failed"
                        FlagError = True
                        Exit Sub
                    End If
                    If flag0762 Then
                        Dim revClass2 As rev0762 = New rev0762(CurrICN)
                        revClass2.g0378Flow(dt_762Denial, CurrentICNEDS1)
                    End If

                    pullEDSScreen(CurrICN, 1)
                    SendKeyes("A" & I + 1, 1, 76)
                    apiChk()
                    SendEnter()

                    If GetText(2, 2, 2) = "MP" Then
                        SendKeyes("EDS", 2, 2)
                        SendEnter()
                        apiChk()
                        If GetText(4, 15, 10) = "9999999999" Or GetText(4, 15, 10) = "AA12345678" Or GetText(4, 15, 10) = "----------" Then
                            SendKeyesFor01(OriginalDraft, 4, 32)
                            SendKeyesFor01(OriginalICN, 4, 15)
                            SendEnter()
                            apiChk()
                            SendKeyes("A", 1, 75)
                            'SendEnter()
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                            If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                                SendKeyes("N", 4, 79)
                                SendEnter()
                                apiChk()
                                SendKeyes("A", 1, 75)
                                autECLPSObj.SendKeys("[ENTER]")
                                'SendEnter()
                                apiChk()
                            End If
                        ElseIf GetText(3, 2, 5) = "EDS 6" Then
                            SendKeyes("A", 1, 75)
                            'SendEnter()
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                            If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                                SendKeyes("N", 4, 79)
                                SendEnter()
                                apiChk()
                                SendKeyes("A", 1, 75)
                                autECLPSObj.SendKeys("[ENTER]")
                                'SendEnter()
                                apiChk()
                            End If
                        End If
                    End If


                    If (GetText(2, 2, 3) <> "MPC") Then
                        pullEDSScreen(CurrICN, 1)
                        SendKeyes("A", 1, 76)
                        apiChk()
                        'autECLPSObj.SendKeys("[ENTER]")
                        'apiChk()
                        'apiChk()
                        'apiChk()
                        SendEnter()
                        If GetText(2, 2, 2) = "MP" Then
                            'RemovePZ1()
                            IncreementMPCCount()
                        End If
                    End If

                    CounterRC = 0
                    'WriteLog("For Suffix " & CStr(I + 1) & " total paid = " & MyClaim(I).TotalPaid & " And 30 Line Amount = " & AmountToBeAdjusted)
                    If CDbl(MyClaim(I).TotalPaid) > 0 Then 'AndAlso MyClaim(I).POS.Length < 7      Changed 3rd March,2022
                        If CDbl(AmountToBeAdjusted) > 0 Then
                            'need to create 30 line
                            If CDbl(AmountToBeAdjusted) >= CDbl(MyClaim(I).TotalPaid) Then

                                If FlagZeroPayment Then 'checkFor30LineRemarkCAApplicable(myMap, FlagZeroPayment, FlagCARemarkCode)
                                    'ProcessClaim(FlagError, FlagPartialProcessed, MyClaim(I).TotalPaid, "CA", I + 1, SubroGationRemarkCode)
                                    ProcessClaim(FlagError, FlagPartialProcessed, MyClaim(I).TotalPaid, "E5", I + 1, SubroGationRemarkCode)    'Changed 9/13/25 -- All 30line RC must be E5
                                Else
                                    ProcessClaim(FlagError, FlagPartialProcessed, MyClaim(I).TotalPaid, "E5", I + 1, SubroGationRemarkCode)
                                End If

                                If FlagError Then Exit Sub

                                FlagPartialProcessed = True

                                AmountToBeAdjusted = CStr(CDbl(AmountToBeAdjusted) - CDbl(MyClaim(I).TotalPaid))
                            Else
                                If FlagZeroPayment Then
                                    'checkFor30LineRemarkCAApplicable(myMap, FlagZeroPayment, FlagCARemarkCode)
                                    'ProcessClaim(FlagError, FlagPartialProcessed, AmountToBeAdjusted, "CA", I + 1, SubroGationRemarkCode)
                                    ProcessClaim(FlagError, FlagPartialProcessed, AmountToBeAdjusted, "E5", I + 1, SubroGationRemarkCode)      'Changed 9/13/25 -- All 30line RC must be E5
                                Else
                                    ProcessClaim(FlagError, FlagPartialProcessed, AmountToBeAdjusted, "E5", I + 1, SubroGationRemarkCode)
                                End If
                                If FlagError Then Exit Sub
                                FlagPartialProcessed = True
                                AmountToBeAdjusted = "0.00"
                            End If
                            If strResearchComments = "NO SPACE FOR 30 LINE" Then
                                Exit Sub
                            End If
                            If strResearchComments = "Cannot remove NotCov with RC D2" Then
                                Exit Sub
                            End If
                        Else
                            'no need to create 30 line
                            ProcessClaim(FlagError, FlagPartialProcessed, "0.00", "", I + 1, SubroGationRemarkCode, "E5", False)
                            If FlagError Then Exit Sub
                            FlagPartialProcessed = True
                        End If
                    Else
                        'no need to create 30 line
                        ProcessClaim(FlagError, FlagPartialProcessed, "0.00", "", I + 1, SubroGationRemarkCode, "E5", False)
                        If FlagError Then Exit Sub
                        FlagPartialProcessed = True
                    End If
                End If
            Next
            FlagPartialProcessed = False
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
        Finally
            WriteLog($"===================Exit {GetCurrentMethod.Name}====================")
        End Try
    End Sub
    Sub ProcessClaim(ByRef FlagError As Boolean, ByRef FlagPartialProcessed As Boolean, ByVal AmountToBeAdjusted As String, ByVal RemarkCode30Line As String,
                     ByVal CurrentSuffix As Integer, ByVal SubRogationRemarkCode As String,
                     Optional ByVal FloatingReamrkCode As String = "", Optional ByVal is30LineRequired As Boolean = True)
        Try
            '--------ADDED TO REMOVE ANYTHING FROM IRC FIELD IN EDS 6.5  => BY SOURAV ON 14/07/2023
            pullEDSScreen(CurrICN, 1)
            SendKeyes("A" & CurrentSuffix, 1, 76)
            SendEnter()
            apiChk()
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
                Dim SafetyPageCount As Integer = 0
                For I = 6 To 18 Step 3
                    If GetText(I, 2, 3) <> "---" Then Exit For
                    If GetText(I + 1, 50, 2) <> "--" Then
                        SendKeyes("--", I + 1, 50)
                        SendEnter()
                        apiChk()
                    End If

                    If I = 18 Then
                        WriteLog(GetTextRect())
                        Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                        SendF8()
                        If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                            Exit For
                        Else
                            SafetyPageCount = SafetyPageCount + 1

                            ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                            If SafetyPageCount > 50 Then
                                WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                Exit For
                            End If
                        End If
                        I = 3
                    End If
                Next
                WriteLog(GetTextRect)
                SendKeyes("A", 1, 75)
                'SendEnter()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
                WriteLog(GetTextRect)
                If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    'SendEnter()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                End If
            End If
            '-------------
            Dim RCLineLevel As String = "", IRCRemarkCode As String = "", IRCRemCode As String = ""
            'Dim TempFlag As Boolean
            Dim CounterRC As Integer = 0
            Dim Count2 As Integer
            Dim EditWarning As String = ""
            Dim PageCount As Integer, LineCount As Integer
            Try
                For I = 10 To 22 Step 2                 'Added on 14/06/2023 by Sourav- to skip a claim if Not covered can't be removed in any line with D2 remark code
                    If GetText(I, 4, 6) = "------" Then Exit For
                    If GetText(I, 36, 2) = "D2" Then
                        If Trim(GetText(I, 50, 5)) <> "" Then
                            strResearchComments = "Cannot remove NotCov with RC D2"
                            Exit Sub
                        End If
                    End If
                Next
                CheckBypassHistoryField()
                Dim isAllowableMatch As Boolean = MatchAllowable()
                'Added on 3rd March,2022
                If is30LineRequired And is30LineToBeused Then
                    If Trim(GetText(22, 1, 80).Replace("-", "")) = "" Then
                        WriteLog("--------------------After creating 30Line--------------------")
                        create30LineInMPCScreen(FlagError, AmountToBeAdjusted, RemarkCode30Line)
                        If Check30IfLinePasted() = False Then
                            strResearchComments = "30 line not pasted"
                            FlagError = True
                            Exit Sub
                        End If
                        WriteLog(GetTextRect)
                    Else
                        WriteLog("NO SPACE FOR 30 LINE")
                        WriteLog(GetTextRect)
                        strResearchComments = "NO SPACE FOR 30 LINE"
                        FlagError = True
                        Exit Sub
                    End If
                Else
                    WriteLog("==============================================================")
                    WriteLog(GetTextRect)
                End If
                'till here changed on 3rd March,2022

                Do
                    RCLineLevel = ""
                    'WriteLog("Before CheckHICFInnProcessing")
                    'CheckHICFInnProcessing(TempFlag, RCLineLevel, IRCRemarkCode, ConvertToFormat(CStr(CurrentSuffix), 2, True, "0"), RemarkCode30Line)
                    'WriteLog("After CheckHICFInnProcessing")
                    'If TempFlag = True Then
                    '    'Reduction logic applies, exit from here
                    '    strResearchStatus = "SKIPPED"
                    '    strResearchComments = "Reduction logic applies, Please check manually"
                    '    Exit Sub
                    'End If
                    remove7YforCOB(RCLineLevel, IRCRemarkCode)
                    If CheckRemarkCodePlacementOnPaymentScreen(RCLineLevel) = False Then
                        If CounterRC > 2 Then
                            strResearchStatus = "SKIPPED"
                            strResearchComments = "Unable to hold proper remarkcode in payment screen"
                            FlagError = True
                            Exit Sub
                        End If
                        CounterRC += 1
                    Else
                        Exit Do
                    End If
                Loop
                If FloatingReamrkCode <> "" Then
                    If isDeniedEntireSuffixFromMPCScreen() = False Then
                        If FloatingReamrkCode = "E5" And is30LineToBeused Then
                            If SubRogationRemarkCode.Trim <> "" Then
                                SendKeyes("EDS", 2, 2)
                                SendEnter()
                                PageCount = 1
                                LineCount = 1
                                For I = 6 To 18 Step 3

                                    If GetText(I, 1, 80, True) = "" Then Exit For
                                    'SendKeyes("---", I + 1, 54)
                                    'SendKeyes("---", I + 1, 50)
                                    If GetText(I, 38, 2) <> "30" Then SendKeyes("13", I, 38)
                                    For Each str As String In IRCRemarkCode.Split(",")
                                        If str.Trim <> "" Then
                                            If str.Split("-")(0) = LineCount Then
                                                If InStr(IRCRemCode, Split(str, "-")(1)) > 0 Then
                                                    SendKeyes("E5", I + 1, 54)
                                                Else
                                                    SendKeyes(Split(str, "-")(1), I + 1, 54)
                                                    IRCRemCode += Split(str, "-")(1) + ","
                                                    Exit For
                                                End If

                                            End If
                                        End If
                                    Next
                                    For Each str As String In RCLineLevel.Split(",")
                                        If str.Trim <> "" Then
                                            If str.Split("-")(0) = LineCount Then
                                                If GetText(I, 38, 2, True) <> "30" Then
                                                    SendKeyes("13", I, 38)
                                                End If
                                                SendKeyes(Split(str, "-")(1), I, 45)
                                                SendKeyes(Split(str, "-")(1), I + 1, 50)
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If GetText(I + 1, 54, 2) = "--" Then SendKeyes("E5", I + 1, 54)
                                    If I = 18 Then
                                        If PageCount = 2 Then Exit For
                                        PageCount += 1
                                        SendEnter()
                                        SendF8()
                                        I = 3
                                    End If
                                    LineCount += 1
                                    'If LineCount > PaymentScreenLineCount Then Exit For
                                Next
                                SendEnter()
                                'End If
                                SendKeyes("A", 1, 75)
                                'SendEnter()
                                autECLPSObj.SendKeys("[ENTER]")
                                'apiChk()
                                'apiChk()
                                If GetText(2, 2, 2) = "MP" Then
                                    'RemovePZ1()
                                    IncreementMPCCount()
                                End If
                                If GetText(2, 2, 2, True) = "AR" Then
                                    SendKeyes("S", 3, 5)
                                    SendEnter()
                                End If
                                If GetText(2, 2, 2) <> "MP" Then
                                    strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                                    ExitAllAndReturnWithException()
                                End If
                                'SendKeyes("G" & ClaimLevelRemarkCode, 2, 2)
                                'SendEnter()
                                PutCopayNBR()
                                SendOverrideInScreen("01")

                                If SubRogationRemarkCode.Trim <> "" Then
                                    If SubRogationRemarkCode.Trim = "HI" Then
                                        SendKeyes("G" & "E5", 2, 2)
                                        SendEnter()
                                        Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                        SendKeyes("HI", 2, Count2 + 2)
                                        SendEnter()
                                    Else
                                        SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                        SendEnter()
                                    End If

                                End If
                            Else
                                SendKeyes("G" & FloatingReamrkCode, 2, 2)
                                SendEnter()
                            End If
                        Else
                            If SubRogationRemarkCode.Trim <> "" Then
                                If SubRogationRemarkCode.Trim = "HI" Then
                                    SendKeyes("G" & "E5", 2, 2)
                                    SendEnter()
                                    Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                    SendKeyes("HI", 2, Count2 + 2)
                                    SendEnter()
                                Else
                                    SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                    SendEnter()
                                End If

                            End If
                        End If
                    Else
                        If InStr(GetText(2, 1, 80), ",,") > 0 Then
                            SendEof(2, InStr(GetText(2, 1, 80), ",,"))
                            SendEnter()
                            If SubRogationRemarkCode.Trim <> "" Then
                                If SubRogationRemarkCode.Trim = "HI" Then
                                    SendKeyes("G" & "E5", 2, 2)
                                    SendEnter()
                                    Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                    SendKeyes("HI", 2, Count2 + 2)
                                    SendEnter()
                                Else
                                    SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                    SendEnter()
                                End If

                            End If
                        End If
                    End If
                Else
                    If is30LineRequired Then
                        For J As Integer = 10 To 22 Step 2
                            If GetText(J, 1, 80).Replace("-", "").Trim = "" Then Exit For
                            If GetText(J, 36, 2) = RemarkCode30Line Then
                                FloatingReamrkCode = ""
                            Else
                                FloatingReamrkCode = "E5"
                            End If
                        Next
                        If FloatingReamrkCode <> "" Then
                            If isDeniedEntireSuffixFromMPCScreen() = False Then
                                If SubRogationRemarkCode.Trim <> "" Then
                                    SendKeyes("EDS", 2, 2)
                                    SendEnter()
                                    PageCount = 1
                                    LineCount = 1
                                    For I = 6 To 18 Step 3

                                        If GetText(I, 1, 80, True) = "" Then Exit For
                                        'SendKeyes("---", I + 1, 54)
                                        'SendKeyes("---", I + 1, 50)
                                        If GetText(I, 38, 2) <> "30" Then SendKeyes("13", I, 38)
                                        For Each str As String In IRCRemarkCode.Split(",")
                                            If str.Trim <> "" Then
                                                If str.Split("-")(0) = LineCount Then
                                                    If InStr(IRCRemCode, Split(str, "-")(1)) > 0 Then
                                                        SendKeyes("E5", I + 1, 54)
                                                    Else
                                                        SendKeyes(Split(str, "-")(1), I + 1, 54)
                                                        IRCRemCode += Split(str, "-")(1) + ","
                                                        Exit For
                                                    End If

                                                End If
                                            End If
                                        Next
                                        For Each str As String In RCLineLevel.Split(",")
                                            If str.Trim <> "" Then
                                                If str.Split("-")(0) = LineCount Then
                                                    If GetText(I, 38, 2, True) <> "30" Then
                                                        SendKeyes("13", I, 38)
                                                    End If
                                                    SendKeyes(Split(str, "-")(1), I, 45)
                                                    SendKeyes(Split(str, "-")(1), I + 1, 50)
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                        If GetText(I + 1, 54, 2) = "--" Then SendKeyes("E5", I + 1, 54)
                                        If I = 18 Then
                                            If PageCount = 2 Then Exit For
                                            PageCount += 1
                                            SendEnter()
                                            SendF8()
                                            I = 3
                                        End If
                                        LineCount += 1
                                        'If LineCount > PaymentScreenLineCount Then Exit For
                                    Next
                                    SendEnter()
                                    'End If
                                    SendKeyes("A", 1, 75)
                                    'SendEnter()
                                    autECLPSObj.SendKeys("[ENTER]")
                                    'apiChk()
                                    'apiChk()
                                    If GetText(2, 2, 2) = "MP" Then
                                        'RemovePZ1()
                                        IncreementMPCCount()
                                    End If
                                    If GetText(2, 2, 2, True) = "AR" Then
                                        SendKeyes("S", 3, 5)
                                        SendEnter()
                                    End If
                                    If GetText(2, 2, 2) <> "MP" Then
                                        strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                                        ExitAllAndReturnWithException()
                                    End If
                                    'SendKeyes("G" & ClaimLevelRemarkCode, 2, 2)
                                    'SendEnter()
                                    PutCopayNBR()
                                    SendOverrideInScreen("01")

                                    If SubRogationRemarkCode.Trim <> "" Then
                                        If SubRogationRemarkCode.Trim = "HI" Then
                                            SendKeyes("G" & "E5", 2, 2)
                                            SendEnter()
                                            Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                            SendKeyes("HI", 2, Count2 + 2)
                                            SendEnter()
                                        Else
                                            SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                            SendEnter()
                                        End If

                                    End If
                                Else
                                    SendKeyes("G" & FloatingReamrkCode, 2, 2)
                                    SendEnter()
                                End If
                            Else
                                If InStr(GetText(2, 1, 80), ",,") > 0 Then
                                    SendEof(2, InStr(GetText(2, 1, 80), ",,"))
                                    SendEnter()
                                    If SubRogationRemarkCode.Trim <> "" Then
                                        If SubRogationRemarkCode.Trim = "HI" Then
                                            SendKeyes("G" & "E5", 2, 2)
                                            SendEnter()
                                            Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                            SendKeyes("HI", 2, Count2 + 2)
                                            SendEnter()
                                        Else
                                            SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                            SendEnter()
                                        End If

                                    End If
                                End If
                            End If
                        Else
                            If SubRogationRemarkCode.Trim <> "" Then
                                If SubRogationRemarkCode.Trim = "HI" Then
                                    SendKeyes("G" & "E5", 2, 2)
                                    SendEnter()
                                    Count2 = InStr(GetText(2, 2, 79, True), "GE5")
                                    SendKeyes("HI", 2, Count2 + 2)
                                    SendEnter()
                                Else
                                    SendKeyes("G" & SubRogationRemarkCode, 2, 2)
                                    SendEnter()
                                End If

                            End If
                        End If
                    End If
                End If
                Count2 = 0
                Dim flagOKToPay As Boolean = False
                remove7YforCOB(RCLineLevel, IRCRemarkCode)
                EditWarning = ""
                For KK As Integer = 10 To 22 Step 2
                    If GetText(KK, 32, 1).ToUpper = "Z" Then SendKeyes("2", KK, 32)
                Next
                Do
                    SendEnter()
                    If Count2 > 50 Then Exit Do
                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                    EditWarning = EditWarning & GetText(24, 1, 80, True)
                    Count2 = Count2 + 1
                Loop

                If InStr(EditWarning, "WE2560") > 0 AndAlso (InStr(EditWarning, "E2497") = 0 And InStr(EditWarning, "W512") = 0) Then
                    SendOverrideInScreen("09")
                    EditWarning = ""
                    Count2 = 0
                    Do
                        SendEnter()
                        If Count2 > 50 Then Exit Do
                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        EditWarning = EditWarning & GetText(24, 1, 80, True)
                        Count2 = Count2 + 1
                    Loop
                End If
                'Added below loop to fill empty RC fields  
                For I As Integer = 10 To 22 Step 2
                    If GetText(I, 1, 80).Replace("-", "").Trim <> "" Then
                        If GetText(I, 36, 2) = "--" Then
                            If InStr(EditWarning, "W1616") > 0 Then
                                SendKeyes("D2", I, 36)
                                apiChk()
                                autECLPSObj.SendKeys("[ENTER]")
                                apiChk()
                                apiChk()
                            Else
                                strResearchStatus = "SKIPPED"
                                strResearchComments = "Skipped: RC field empty and W1616 not present"
                                FlagError = True
                                Exit Sub
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
                'Added below logic to correct the incorrect RC field in 30Line
                Dim pdAmt As Decimal
                If InStr(EditWarning, "W169") > 0 And is30LineRequired Then
                    'Dim pdAmtArray1() As String = EditWarning.Split("W169P1-")
                    'Dim pdAmtArray2() As String = (Trim(pdAmtArray1(1))).Split(",")
                    'Dim pdAmt As Integer = Convert.ToInt32(pdAmtArray2(0))
                    Dim word As String() = New String() {"W169P1-"}
                    Dim wordArr As String() = EditWarning.Split(word, StringSplitOptions.RemoveEmptyEntries)
                    Dim pdAmtArr As String()
                    If wordArr.Length > 1 Then
                        pdAmtArr = wordArr(1).Trim.Split(",")
                    Else
                        pdAmtArr = wordArr(0).Trim.Split(",")
                    End If
                    If CheckIfAlphaCharacterIsPresent(pdAmtArr(0)) Then
                        strResearchComments = "Cannot extract paid amount from EDIT"
                        Exit Sub
                    Else
                        pdAmt = pdAmtArr(0).Trim
                    End If
                    WriteLog($"Paid amount from Editwarning= ${pdAmt}")
                    If pdAmt > 0 Then
                        For J As Integer = 10 To 22 Step 2
                            If GetText(J, 1, 80).Replace("-", "").Trim = "" And GetText(J - 2, 29, 2) = "30" Then
                                SendKeyes("E5", J - 2, 36)
                                apiChk()
                                Exit For
                            ElseIf J = 22 And GetText(J, 29, 2) = "30" Then
                                SendKeyes("E5", 22, 36)
                                apiChk()
                                Exit For
                            End If
                        Next
                    ElseIf pdAmt = 0 Then
                        For J As Integer = 10 To 22 Step 2
                            If GetText(J, 1, 80).Replace("-", "").Trim = "" And GetText(J - 2, 29, 2) = "30" Then
                                'SendKeyes("CA", J - 2, 36)
                                SendKeyes("E5", J - 2, 36)    'Changed 9/13/25 -- All 30line RC must be E5
                                apiChk()
                                Exit For
                            ElseIf J = 22 And GetText(J, 29, 2) = "30" Then
                                'SendKeyes("CA", J - 2, 36)
                                SendKeyes("E5", J - 2, 36)    'Changed 9/13/25 -- All 30line RC must be E5
                                apiChk()
                                Exit For
                            End If
                        Next
                    End If
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    apiChk()
                End If
                If is30LineRequired Then       'Added on 27/2/2026 for double check
                    For J As Integer = 10 To 22 Step 2
                        If GetText(J, 1, 80).Replace("-", "").Trim = "" Then Exit For

                        If GetText(J, 29, 2) <> "30" And GetText(J, 36, 2) = "E5" And GetText(J, 50, 5).Trim = "" Then
                            SendKeyes("D2", J, 36)
                            'SendEnter()
                        End If

                        'If GetText(J, 36, 2) = RemarkCode30Line Then
                        '    FloatingReamrkCode = ""
                        'Else
                        '    FloatingReamrkCode = "E5"
                        'End If
                    Next
                End If
                Dim Count3 As Integer = 0
                Do
                    SendEnter()
                    apiChk()
                    If Count3 > 50 Then Exit Do
                    If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                    EditWarning = EditWarning & GetText(24, 1, 80, True)
                    Count3 = Count3 + 1
                Loop
                If InStr(EditWarning, "W118 OK TO PAY") > 0 OrElse
               (InStr(EditWarning, "WE") = 0 AndAlso InStr(EditWarning, ",E") = 0 AndAlso InStr(EditWarning, ", E") = 0) Then
                    flagOKToPay = True
                End If
                'PasteHistoryICN(CurrICN, FlagError, strResearchStatus, strResearchComments)
                'If is30LineRequired And is30LineToBeused Then create30LineInMPCScreen(FlagError, AmountToBeAdjusted, RemarkCode30Line)
                If FlagError = True Then
                    Exit Sub
                    'Return strResearchComments
                End If
                If flagOKToPay = False And GetColorScheme(7, 5) = 3 Then
                    Dim isremoved As Boolean = removeC85RevFrom30Line()
                    If isremoved Then flagOKToPay = checkOkToPayinEdit()
                End If
                If flagOKToPay Then
                    If isNCopayNBR = True Then
                        PlaceNCopayNBRinMPC()
                    End If
                    WriteLog($"-------------------------Suffix-{CurrentSuffix} BEFORE MPP----------------------------")
                    WriteLog(GetTextRect)
                    If ICESDenialFlag = False And flagDenial = False Then
                        If isAllowableMatch = False Then
                            'strResearchComments = "Allowable mismatch before MPP"
                            'FlagError = True
                            'Exit Sub
                        End If
                    End If
                    remove7YforCOB(RCLineLevel, IRCRemarkCode)
                    'SendF12()
                    'If CheckClaimProcessedBeforeMPP() = True Then
                    '    strResearchComments = "Current claim is already processed"
                    '    FlagError = True
                    '    Exit Sub
                    'End If
                    'SendF12()
                    SendMPP()
                    If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                        If CurrentSuffix = SuffixCountOfCurrentICN Then
                            If pdAmt > 5000 Then
                                IsHDClaim = True
                                strResearchComments = "Claim process by tool; Paid amt>5000"
                            Else
                                strResearchComments = "Claim process by tool; All suffixes processed"
                            End If

                        Else
                            FlagPartialProcessed = True
                            strResearchComments = $"Suffix {CurrentSuffix} of Claim process by tool"
                            WriteLog($"-------------------------SUFFIX-{CurrentSuffix} AFTER MPP----------------------------")
                            WriteLog(GetTextRect)
                        End If
                    End If

                    If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                        strResearchStatus = "SKIPPED"
                        strResearchComments = $"Suffix {CurrentSuffix}:- " + GetText(24, 1, 80, True)
                        FlagError = True
                        Exit Sub
                    End If
                    CaptureScreenLog()
                Else
                    WriteLog($"============OK TO PAY not found in MPC screen-----------------")
                    WriteLog(GetTextRect)
                    If InStr(EditWarning, "WE1434INVESTIGATE COB") > 0 Then
                        strResearchStatus = "SKIPPED"
                        strResearchComments = "WE1434INVESTIGATE COB"
                        FlagError = True
                        Exit Sub
                    Else
                        strResearchStatus = "SKIPPED"
                        strResearchComments = EditWarning + $"of Suffix {CurrentSuffix}"
                        FlagError = True
                        WriteLog(GetTextRect)
                        Exit Sub
                    End If
                End If
            Catch ex As Exception
                FlagError = True
                WriteLog("Exception: " + ex.Message.ToString)
                strResearchComments = GetCurrentMethod.Name() + " method: " + ex.Message.ToString
                Exit Sub
            End Try

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub

    Public Function checkOkToPayinEdit() As Boolean
        Dim EditWarning As String = String.Empty
        Dim Count3 As Integer = 0
        Dim flagOKToPay = False
        Try
            Do
                SendEnter()
                apiChk()
                If Count3 > 50 Then Exit Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                EditWarning = EditWarning & GetText(24, 1, 80, True)
                Count3 = Count3 + 1
            Loop
            If InStr(EditWarning, "W118 OK TO PAY") > 0 OrElse
           (InStr(EditWarning, "WE") = 0 AndAlso InStr(EditWarning, ",E") = 0 AndAlso InStr(EditWarning, ", E") = 0) Then
                flagOKToPay = True
            End If
        Catch ex As Exception

        End Try
        Return flagOKToPay
    End Function

    Public Sub remove7YforCOB(ByVal RC1 As String, ByVal RC2 As String)
        Try
            Dim line As Integer = 1
            SendEnter()

            If IsCOB AndAlso is7YinScreen() Then
                For I As Integer = 10 To 22 Step 2
                    If GetText(I, 36, 2) = "7Y" And GetColorScheme(I, 36) = 2 Then
                        SendKeyes("01", I, 29)
                    End If
                Next
                SendEnter()
                For i As Integer = 10 To 22 Step 2
                    For Each Str As String In RC1.Split(",")
                        If Str.Trim <> "" Then
                            If CInt(Str.Split("-")(0)) = line AndAlso GetText(i, 36, 2) <> Str.Split("-")(1) Then
                                SendKeyes("  ", i, 36)
                                Exit For
                            End If
                        End If
                    Next
                    line += 1
                Next
                SendEnter()
            Else
                Dim FlagNcopay As Boolean = False
                For I As Integer = 10 To 22 Step 2
                    If GetText(I, 4, 6, True) = "NCOPAY" Then
                        FlagNcopay = True
                        Exit For
                    End If
                Next
                If FlagNcopay Then
                    FlagNcopay = False
                    For i As Integer = 10 To 22 Step 2
                        For Each Str As String In RC1.Split(",")
                            If Str.Trim <> "" Then
                                If CInt(Str.Split("-")(0)) = line Then
                                    If GetText(i, 36, 2) <> Str.Split("-")(1) Then
                                        FlagNcopay = True
                                        SendKeyes(Str.Split("-")(1), i, 36)
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        line += 1
                    Next
                    If FlagNcopay Then SendOverrideInScreen("02")
                Else
                    For i As Integer = 10 To 22 Step 2
                        For Each Str As String In RC1.Split(",")
                            If Str.Trim <> "" AndAlso GetText(i, 4, 6, True) <> "NCOPAY" Then
                                If CInt(Str.Split("-")(0)) = line Then
                                    If GetText(i, 36, 2) <> Str.Split("-")(1) Then
                                        SendKeyes(Str.Split("-")(1), i, 36)
                                        SendKeyes("02", i, 29)
                                        SendEnter()
                                    Else
                                        If GetText(i, 29, 2) = "02" Then
                                            SendKeyes("01", i, 29)
                                            SendEnter()
                                            If GetColorScheme(i, 29) = 3 Then
                                                For J As Integer = 10 To 22 Step 2
                                                    If i <> J Then
                                                        If GetColorScheme(J, 29) = 3 AndAlso InStr(GetText(24, 1, 80), "WE1013INCONSISTENT DATA") > 0 Then 'GetText(J, 29, 2) <> "30" Then
                                                            SendKeyes("01", J, 29)
                                                            SendEnter()
                                                            'If GetColorScheme(i, 29) <> 3 Then Exit For
                                                        End If
                                                    End If
                                                Next
                                            End If
                                            If GetText(i, 36, 2) <> Str.Split("-")(1) Then
                                                SendKeyes(Str.Split("-")(1), i, 36)
                                                SendKeyes("02", i, 29)
                                                SendEnter()
                                            End If
                                            If GetColorScheme(i, 29) = 3 Then
                                                For J As Integer = 10 To 22 Step 2
                                                    If i <> J Then
                                                        If GetColorScheme(J, 29) = 3 AndAlso InStr(GetText(24, 1, 80), "WE1013INCONSISTENT DATA") > 0 Then 'GetText(J, 29, 2) <> "30" Then
                                                            SendKeyes("02", J, 29)
                                                            SendEnter()
                                                            'If GetColorScheme(i, 29) <> 3 Then Exit For
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next
                        line += 1
                    Next
                End If

                SendEnter()
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function is7YinScreen() As Boolean
        Try
            For I As Integer = 10 To 22 Step 2
                If GetText(I, 36, 2) = "7Y" And GetColorScheme(I, 36) = 2 Then Return True
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return False
    End Function
    Public Function removeC85RevFrom30Line() As Boolean
        Dim isRemoved As Boolean = False
        Dim mpcLine As Integer, InValid30Line As Integer
        Try
            For mpcLine = 10 To 22 Step 2
                If GetText(mpcLine, 4, 6, True) = "------" Then
                    Exit For
                ElseIf InStr("C85REV", GetText(mpcLine, 4, 6, True)) > 1 And GetText(mpcLine, 36, 2) = "30" And GetColorScheme(mpcLine, 4) = 3 Then
                    InValid30Line = mpcLine
                    SendKeyes("C09REV", mpcLine, 4)
                    SendEnter()
                    Exit For
                End If
            Next
            If GetColorScheme(InValid30Line, 4) = 0 Then isRemoved = True
        Catch ex As Exception

            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return isRemoved

    End Function

    Public Function isDeniedEntireSuffixFromMPCScreen() As Boolean
        Try
            Dim BillCharge As Decimal = 0, DeniedAmount As Decimal = 0

            For I As Integer = 10 To 22 Step 2
                If Replace(GetText(I, 1, 80, True), "-", "").Trim = "" Then Exit For
                If InStr(InvalidProcCode, GetText(I, 4, 6, True)) = 0 And GetText(I, 29, 2) <> "30" Then
                    BillCharge += CDec(GetText(I, 41, 5, True) + "." + GetText(I, 47, 2))
                    DeniedAmount += CDec(GetText(I, 50, 5, True) + "." + GetText(I, 56, 2))
                End If
            Next
            If BillCharge = DeniedAmount Then
                Return True
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return False
    End Function
    Function CheckRemarkCodePlacementOnPaymentScreen(ByVal strRC As String) As Boolean
        Try
            Dim LineCount As Integer
            For i As Integer = 10 To 22 Step 2
                LineCount = ((i - 10) / 2) + 1
                For Each str As String In strRC.Split(",")
                    If str.Trim <> "" Then
                        If str.Split("-")(0) = LineCount Then
                            If GetText(i, 36, 2) <> str.Split("-")(1) Then
                                If CDbl(GetText(i, 41, 5) & "." & GetText(i, 47, 2)) <> CDbl(GetText(i, 50, 5) & "." & GetText(i, 56, 2)) Then
                                    If GetText(i, 36, 2) <> "B9" Then Return False
                                Else
                                    Return False
                                End If

                            End If
                            'Exit For
                        End If
                    End If
                Next

            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return True
    End Function
    Public Sub ReadClaimPaymentScreen(ByRef Myclaim() As ClaimFromPaymentScreen, ByRef MyDrafts() As DraftDetails, ByRef CurrentClaim() As Claim, ByRef MyISI() As ISI, ByVal CureentClaimCount As Integer, ByRef FlagError As Boolean, ByVal ICN As String, ByVal MaxSuffix As Integer)
        Dim Suffixes As String = "", drafts As String = ""
        Dim CountOfLine As Integer = -1

        For I As Integer = 1 To MaxSuffix
            For J As Integer = LBound(MyDrafts) To UBound(MyDrafts)
                If IsNumeric(MyDrafts(J).Suffix) AndAlso CInt(MyDrafts(J).Suffix) = I Then
                    If MyDrafts(J).Reversal = False Then
                        Suffixes += MyDrafts(J).Suffix & ","
                        drafts += MyDrafts(J).Draft & ","
                    End If
                    Exit For
                End If
            Next
        Next



        'For I As Integer = LBound(MyDrafts) To UBound(MyDrafts)
        '    If IsNumeric(MyDrafts(I).Draft) AndAlso (MyDrafts(I).ReversalType = "Not Voided" OrElse InStr("74,71", MyDrafts(I).ReversalType) = 0) Then
        '        If InStr(Suffixes, MyDrafts(I).Suffix) = 0 Then
        '            Suffixes += MyDrafts(I).Suffix & ","
        '            drafts += MyDrafts(I).Draft & ","
        '        End If
        '    End If
        'Next
        ReDimension(Myclaim, MaxSuffix - 1)
        For I As Integer = 1 To MaxSuffix
            CountOfLine = -1
            If InStr(Suffixes, ConvertToFormat(CStr(I), 2, True, "0")) > 0 Then
                'copy from mhi structure
                For J = 0 To CureentClaimCount - 1
                    If CInt(CurrentClaim(J).Suffix) = I And InStr(drafts, CurrentClaim(J).Draft) > 0 And CurrentClaim(J).Voided = False Then
                        ReDim Myclaim(I - 1).POS(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).ProcCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).FDOS(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).LDOS(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).NBR(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Override(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).PayeeCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).RemarkCode(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Charge(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).NotCovered(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Allowable(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Deductible(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Coins(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Paid(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Sanction(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).HxCindex(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).HxLIndex(UBound(CurrentClaim(J).ServiceLine) - 1)
                        ReDim Myclaim(I - 1).Modifier(UBound(CurrentClaim(J).ServiceLine) - 1)
                        Myclaim(I - 1).CauseCode = CurrentClaim(J).CauseCode
                        Myclaim(I - 1).Suffix = CInt(CurrentClaim(J).Suffix)
                        Myclaim(I - 1).Draft = CurrentClaim(J).Draft
                        Myclaim(I - 1).DateOfProcessing = CurrentClaim(J).DateOfProcessing
                        Myclaim(I - 1).TotalPaid = CurrentClaim(J).PaidAmount
                        For K = 0 To UBound(CurrentClaim(J).ServiceLine) - 1
                            Myclaim(I - 1).POS(K) = CurrentClaim(J).POS(K)
                            Myclaim(I - 1).ProcCode(K) = CurrentClaim(J).ProcCode(K)
                            Myclaim(I - 1).FDOS(K) = CurrentClaim(J).FDOS(K)
                            Myclaim(I - 1).LDOS(K) = CurrentClaim(J).LDOS(K)
                            Myclaim(I - 1).NBR(K) = CurrentClaim(J).NBR(K)
                            Myclaim(I - 1).Override(K) = CurrentClaim(J).OV(K)
                            Myclaim(I - 1).PayeeCode(K) = "21"
                            Myclaim(I - 1).RemarkCode(K) = CurrentClaim(J).RemarkCode(K)
                            Myclaim(I - 1).Charge(K) = CurrentClaim(J).Charge(K)
                            Myclaim(I - 1).NotCovered(K) = CurrentClaim(J).NotCovered(K)
                            Myclaim(I - 1).Allowable(K) = CurrentClaim(J).Allowable(K)
                            Myclaim(I - 1).Deductible(K) = CurrentClaim(J).Deductible(K)
                            Myclaim(I - 1).Coins(K) = CurrentClaim(J).CoIns(K)
                            Myclaim(I - 1).Sanction(K) = Mid(CurrentClaim(J).ServiceLine(K), 39, 1)
                            Myclaim(I - 1).Paid(K) = CurrentClaim(J).Paid(K)
                            Myclaim(I - 1).HxCindex(K) = J
                            Myclaim(I - 1).HxLIndex(K) = K
                        Next
                        Exit For
                    End If
                Next
            Else
                'Copy frm payment screen
                pullEDSScreen(ICN, "9")
                SendKeyes("A" & I, 1, 76)
                SendEnter()
                IncreementMPCCount()

                If GetText(2, 2, 2) = "AR" Then
                    FlagError = True
                    strResearchStatus = "SKIPPED"
                    strResearchComments = "Notification Related Claim"
                    Exit Sub
                End If
                If GetText(2, 2, 2) = "MR" Then
                    FlagError = True
                    strResearchStatus = "SKIPPED"
                    strResearchComments = "Patient information mismatch"
                    Exit Sub
                End If
                'in case no lines pended
                If InStr(GetText(24, 1, 80, True), "WE1191NO LINES PENDED") > 0 Or InStr(GetText(24, 1, 80, True), "WE1196DRAFT NBR REQUIRED") > 0 Or InStr(GetText(24, 1, 80, True), "E944INCOMPLETE DATA") > 0 Then
                    For J = LBound(MyDrafts) To UBound(MyDrafts)
                        If CInt(MyDrafts(J).Suffix) = I Then
                            SendClear()
                            SendClear()
                            SendClear()
                            Dim AdjConLine As String = "ADJ," + CurrentClaim(0).Policy + "," + CurrentClaim(0).SSN + "," +
                                                   CurrentClaim(0).EmployeeName + "," + CurrentClaim(0).Relationship + "," +
                                                   MyDrafts(J).Draft + ",I" + CurrentClaim(0).ICN
                            SendKeyes(AdjConLine, 2, 2)
                            SendEnter()
                            Exit For
                        End If
                    Next
                End If
                If InStr(GetText(2, 1, 80), "PZ1") > 0 Then
                    SendKeyes("[eraseeof]", 2, InStr(GetText(2, 1, 80), "PZ1"))
                    'SendOverrideInScreen("01")
                    CaptureScreenLog()
                End If
                RemoveBlankAndPISLLine()
                'ICESCheck(CurrICN)
                SendOverrideInScreen("01")
                If CheckWhetherAllEDS1LinesPulledInPaymentScreen(CurrentClaimModifier, I) = False Then
                    strResearchStatus = "SKIPPED"
                    strResearchComments = "All Eds1 Lines are not pulled in payment screen"
                    FlagError = True
                    Exit Sub
                End If
                If CheckIfMultipleCopayLineWithNoNBR(I) Then
                    FlagError = True
                    strResearchStatus = "skipped"
                    strResearchComments = "more than 2 copay line present with 0 units. manual intervention required."
                    Exit Sub
                End If
                Dim EditWarning As String = initialOperationsToBeDoneWhenPaymentScreenIsPulled(I)
                If InStr(EditWarning, "W512") > 0 Then 'Chk Man Spi
                    FlagError = True
                    strResearchStatus = "skipped"
                    strResearchComments = "Check manual spi exists."
                    ExitAllAndReturnWithException()
                    Exit Sub
                End If
                'FollowICESInstructionInPaymentScreen(MyISI, I)
                Dim RCLine As String = "", RcLine2 As String = ""
                ' CheckHICFInnProcessing(FlagError, RCLine, RcLine2, I)
                CountOfLine = -1
                Dim EditWarning1 = ""
                Dim Count2 = 0
                Do
                    SendEnter()
                    If Count2 > 50 Then Exit Do
                    If InStr(EditWarning1, GetText(24, 1, 80, True)) > 0 Then Exit Do
                    EditWarning1 = EditWarning1 & GetText(24, 1, 80, True)
                    Count2 = Count2 + 1
                Loop
                If InStr(EditWarning1, "WE2560") > 0 AndAlso (InStr(EditWarning1, "E2497") = 0 And InStr(EditWarning1, "W512") = 0) Then
                    SendOverrideInScreen("09")
                    EditWarning1 = ""
                    Count2 = 0
                    Do
                        SendEnter()
                        If Count2 > 50 Then Exit Do
                        If InStr(EditWarning1, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        EditWarning1 = EditWarning1 & GetText(24, 1, 80, True)
                        Count2 = Count2 + 1
                    Loop
                End If
                'count lines in payment screen
                For J As Integer = 10 To 22 Step 2
                    If GetText(J, 1, 23, True).Replace("-", "").Trim = "" Then Exit For
                    Dim POS As String = GetText(J, 2, 2)
                    Dim ProcCode As String = GetText(J, 4, 6)
                    Dim Fdos As String = GetText(J, 11, 6)
                    Dim Ldos As String = GetText(J, 18, 6)
                    Dim NBR As String = GetText(J, 25, 3)
                    Dim OV As String = GetText(J, 29, 2)
                    Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                    Dim Remark As String = GetText(J, 36, 2)
                    Dim charge As String = GetText(J, 41, 8, True).Replace(" ", ".")
                    Dim NotCove As String = GetText(J, 50, 8, True).Replace(" ", ".")
                    Dim Allwable As String = GetText(J + 1, 30, 8, True).Replace(" ", ".")
                    Dim Deductible As String = GetText(J + 1, 41, 8, True).Replace(" ", ".")
                    Dim Coins As String = GetText(J + 1, 55, 3, True)
                    Dim Paid As String = GetText(J + 1, 61, 8, True).Replace(" ", ".")
                    If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                        CountOfLine += 1
                    End If
                Next


                ReDim Myclaim(I - 1).POS(CountOfLine)
                ReDim Myclaim(I - 1).ProcCode(CountOfLine)
                ReDim Myclaim(I - 1).FDOS(CountOfLine)
                ReDim Myclaim(I - 1).LDOS(CountOfLine)
                ReDim Myclaim(I - 1).NBR(CountOfLine)
                ReDim Myclaim(I - 1).Override(CountOfLine)
                ReDim Myclaim(I - 1).PayeeCode(CountOfLine)
                ReDim Myclaim(I - 1).RemarkCode(CountOfLine)
                ReDim Myclaim(I - 1).Charge(CountOfLine)
                ReDim Myclaim(I - 1).NotCovered(CountOfLine)
                ReDim Myclaim(I - 1).Allowable(CountOfLine)
                ReDim Myclaim(I - 1).Deductible(CountOfLine)
                ReDim Myclaim(I - 1).Coins(CountOfLine)
                ReDim Myclaim(I - 1).Paid(CountOfLine)
                ReDim Myclaim(I - 1).Sanction(CountOfLine)
                ReDim Myclaim(I - 1).HxCindex(CountOfLine)
                ReDim Myclaim(I - 1).HxLIndex(CountOfLine)
                ReDim Myclaim(I - 1).Modifier(CountOfLine)
                Myclaim(I - 1).EditWarning = EditWarning
                Myclaim(I - 1).CauseCode = GetText(7, 4, 1)
                Myclaim(I - 1).Suffix = I
                Myclaim(I - 1).Draft = "0000000000"
                Myclaim(I - 1).DateOfProcessing = "99/99/99"
                Dim TotalPaid As String = "0"
                CountOfLine = 0
                For J As Integer = 10 To 22 Step 2
                    If GetText(J, 1, 80, True).Replace("-", "").Trim = "" Then Exit For
                    Dim POS As String = GetText(J, 2, 2)
                    Dim ProcCode As String = GetText(J, 4, 6)
                    Dim Fdos As String = GetText(J, 11, 6)
                    Dim Ldos As String = GetText(J, 18, 6)
                    Dim NBR As String = GetText(J, 25, 3).Replace(" ", "0")
                    Dim OV As String = GetText(J, 29, 2)
                    Dim Payee As String = GetText(J, 32, 3).Replace(" ", "")
                    Dim Remark As String = GetText(J, 36, 2)
                    Dim charge As String = GetText(J, 41, 5, True).Replace("-", "0") & "." & GetText(J, 47, 2, True) 'GetText(J, 41, 8, True).Replace(" ", ".")
                    Dim NotCove As String = GetText(J, 50, 5, True).Replace("-", "0") & "." & GetText(J, 56, 2, True) 'GetText(J, 50, 8, True).Replace(" ", ".")
                    Dim Allwable As String = GetText(J + 1, 30, 5, True).Replace("-", "0") & "." & GetText(J + 1, 36, 2, True) 'GetText(J + 1, 30, 8, True).Replace(" ", ".")
                    Dim Deductible As String = GetText(J + 1, 41, 5, True).Replace("-", "0") & "." & GetText(J + 1, 47, 2, True) 'GetText(J + 1, 41, 8, True).Replace(" ", ".")
                    Dim Coins As String = GetText(J + 1, 55, 3, True)
                    Dim Paid As String = GetText(J + 1, 61, 5, True).Replace("-", "0") & "." & GetText(J + 1, 67, 2, True) 'GetText(J + 1, 61, 8, True).Replace(" ", ".")
                    Dim Sanction As String = GetText(J, 39, 1)
                    If ProcCode = "NCOPAY" OrElse InStr(InvalidProcCode, ProcCode) = 0 Then
                        Myclaim(I - 1).POS(CountOfLine) = POS
                        Myclaim(I - 1).ProcCode(CountOfLine) = ProcCode.Trim
                        Myclaim(I - 1).FDOS(CountOfLine) = Fdos
                        Myclaim(I - 1).LDOS(CountOfLine) = Ldos
                        Myclaim(I - 1).NBR(CountOfLine) = NBR.Replace("-", "0")
                        Myclaim(I - 1).Override(CountOfLine) = OV
                        Myclaim(I - 1).PayeeCode(CountOfLine) = Payee
                        Myclaim(I - 1).RemarkCode(CountOfLine) = Remark
                        Myclaim(I - 1).Charge(CountOfLine) = charge.Replace("-", "0")
                        Myclaim(I - 1).NotCovered(CountOfLine) = NotCove.Replace("-", "0")
                        Myclaim(I - 1).Allowable(CountOfLine) = Allwable.Replace("-", "0")
                        Myclaim(I - 1).Deductible(CountOfLine) = Deductible.Replace("-", "0")
                        Myclaim(I - 1).Coins(CountOfLine) = Coins.Replace("-", "0")
                        Myclaim(I - 1).Paid(CountOfLine) = Paid.Replace("-", "0")
                        TotalPaid = CStr(CDbl(TotalPaid) + CDbl(Myclaim(I - 1).Paid(CountOfLine)))
                        Myclaim(I - 1).Sanction(CountOfLine) = Sanction
                        Myclaim(I - 1).HxCindex(CountOfLine) = I - 1
                        Myclaim(I - 1).HxLIndex(CountOfLine) = CountOfLine
                        CountOfLine += 1
                    End If
                    Myclaim(I - 1).TotalPaid = TotalPaid
                Next
                'Goto EDS Screen and remove 13 override
                SendKeyes("EDS", 2, 2)
                SendEnter()
                SendOverrideInScreen("01")
                SendKeyes("A", 1, 75)
                autECLPSObj.SendKeys("[ENTER]")
                'SendEnter()
                apiChk()
                apiChk()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                If GetText(2, 2, 2, True) = "AR" Then
                    SendKeyes("S", 3, 5)
                    SendEnter()
                End If
                If GetText(2, 2, 2) <> "MP" Then
                    strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                    ExitAllAndReturnWithException()
                End If
            End If
        Next
    End Sub
    Public Sub GetModifierFromEDS1(ByRef ML() As ModifierList, ByVal ICN As String)
        Dim FDOS As String = "", LDOS As String = "", Modifier As String = "", Proc As String = "", Suffx As String = "", I As Integer, J As Integer, Chrg As String = ""
        Dim NBR As String = "", POS As String = ""
        SendClear()
        SendClear()
        autECLPSObj.SendKeys("MEI,I" & ICN & ",1", 2, 2)
        apiChk()
        SendEnter()
        pullEDSScreen(ICN, "1")

        ' WriteLog("========================= EDS1 ====================================")
        ' WriteLog(GetTextRect())
        autECLPSObj.auteclfieldlist.Refresh
        Dim ReferringPhysician As String = GetText(2, 16, 30, True)
        Dim ServiceingPhysician As String = GetText(22, 2, 50, True)
        For I = 8 To 20 Step 2
            If Trim(autECLPSObj.GetText(I, 1, 80)) = "" Then Exit For
            If autECLPSObj.GetText(I, 2, 6).trim <> "" Then
                FDOS = FDOS & autECLPSObj.GetText(I, 2, 6) & "/"
                LDOS = LDOS & autECLPSObj.GetText(I, 9, 6) & "/"
                Proc = Proc & autECLPSObj.GetText(I, 20, 5) & "/"
                Suffx = Suffx & Trim(autECLPSObj.gettext(I, 59, 2)) & "/"
                Chrg = Chrg & Trim(autECLPSObj.gettext(I, 50, 8)) & "/"
                NBR = NBR & Trim(GetText(I, 37, 3)) & "/"
                POS += GetText(I, 16, 2, True) + "/"
                For J = 26 To 32 Step 2
                    If autECLPSObj.GetText(I, J, 2) <> "--" Then Modifier = Modifier & autECLPSObj.GetText(I, J, 2) & ","
                Next
                Modifier = Modifier & "/"
            End If
            If I = 20 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "NO ADDITIONAL LINES TO DISPLAY") > 0 Then Exit For
                If InStr(autECLPSObj.GetText(24, 1, 80), "PF8 FOR ADDITIONAL LINES") > 0 Then
                    SendF8()
                    ' WriteLog(vbCrLf)
                    ' WriteLog(GetTextRect())
                Else
                    Exit For
                End If

                autECLPSObj.auteclfieldlist.Refresh

                I = 6
            End If

        Next
        If Microsoft.VisualBasic.Right(FDOS, 1) = "/" Then
            FDOS = Microsoft.VisualBasic.Left(FDOS, Len(FDOS) - 1)
        End If
        If Microsoft.VisualBasic.Right(Chrg, 1) = "/" Then
            Chrg = Microsoft.VisualBasic.Left(Chrg, Len(Chrg) - 1)
        End If

        If Microsoft.VisualBasic.Right(LDOS, 1) = "/" Then
            LDOS = Microsoft.VisualBasic.Left(LDOS, Len(LDOS) - 1)
        End If

        If Microsoft.VisualBasic.Right(Proc, 1) = "/" Then
            Proc = Microsoft.VisualBasic.Left(Proc, Len(Proc) - 1)
        End If

        If Microsoft.VisualBasic.Right(Modifier, 1) = "/" Then
            Modifier = Microsoft.VisualBasic.Left(Modifier, Len(Modifier) - 1)
        End If
        If Microsoft.VisualBasic.Right(FDOS, 1) = "/" Then
            NBR = Microsoft.VisualBasic.Left(NBR, Len(NBR) - 1)
        End If

        If Microsoft.VisualBasic.Right(POS, 1) = "/" Then
            POS = Microsoft.VisualBasic.Left(POS, Len(POS) - 1)
        End If
        'MsgBox(UBound(Split(FDOS.Trim, "/")))
        ReDimension(ML, (UBound(Split(FDOS, "/"))))
        For I = LBound(Split(FDOS, "/")) To UBound(Split(FDOS, "/"))
            ML(I).ReferringPhysician = ReferringPhysician
            ML(I).ServicingPhysician = ServiceingPhysician
            If FDOS <> "" Then ML(I).FDOS = Split(FDOS, "/")(I)
            If LDOS <> "" Then ML(I).LDOS = Split(LDOS, "/")(I)
            If Proc <> "" Then ML(I).ProcedureCode = Split(Proc, "/")(I)
            If Modifier <> "" Then
                If Split(Modifier, "/")(I).Trim.Length > 0 Then
                    If Right(Split(Modifier, "/")(I), 1) = "," Then
                        ML(I).Modifier = Left(Split(Modifier, "/")(I), Len(Split(Modifier, "/")(I)) - 1)
                    Else
                        ML(I).Modifier = Split(Modifier, "/")(I)
                    End If
                Else
                    ML(I).Modifier = ""
                End If
                'ML(I).Modifier = IIf(Split(Modifier, "/")(I).Trim.Length > 0, IIf(Right(Split(Modifier, "/")(I), 1) = ",", Left(Split(Modifier, "/")(I), Len(Split(Modifier, "/")(I)) - 1), Split(Modifier, "/")(I)), "")
            End If
            If Suffx <> "" Then ML(I).Suffix = Split(Suffx, "/")(I)
            If Chrg <> "" Then ML(I).Charge = Split(Chrg, "/")(I)
            If NBR <> "" Then ML(I).NBR = Split(NBR, "/")(I)
            If POS <> "" Then ML(I).POS = Split(POS, "/")(I)
        Next
        Dim LineCount As Integer = 1
        Dim sx As String = ""
        For I = LBound(ML) To UBound(ML)
            If sx = "" Then
                sx = ML(I).Suffix
                LineCount = 1
                ML(I).LineNumber = LineCount
                LineCount += 1
            Else
                If sx = ML(I).Suffix Then
                    ML(I).LineNumber = LineCount
                    LineCount += 1
                Else
                    sx = ML(I).Suffix
                    LineCount = 1
                    ML(I).LineNumber = LineCount
                    LineCount += 1
                End If
            End If
            'ML(I).LineNumber = LineCount
            'LineCount += 1
        Next

    End Sub
    Public Sub CopyPossibleIcn(ByRef myList As PossibleICN, ByVal ICN As String)
        Try
            myList.ICN = ICN
            PullMHI(ICN)
            ReDimension(myList.TempClaim, ICN)
            PullMHI(ICN)
            Dim TempClaimCount = ReadMHI(myList.TempClaim)
            GetDraftDetailsInMemory(myList.TempDraftDetails, myList.TempClaim, TempClaimCount)
            myList.MaxSuffix = GetMaxSuffix(myList.TempClaim, TempClaimCount)
            ' WriteLog("History ICN " & myList.ICN & " Payment screen read starts")
            ReadClaimPaymentScreen(myList.TempClaimFromPaymentScreen, myList.TempDraftDetails, myList.TempClaim, TempClaimCount, False, ICN, myList.MaxSuffix)
            'WriteLog(GetPaymentScreen(myList.TempClaimFromPaymentScreen))
            ' WriteLog("History ICN " & myList.ICN & " Payment screen read ends")
            GetClaimDetails(myList.TempClaimDetails, myList.ICN, Office)
            'ReadISI(myList.TempClaimDetails, myList.TempISI)
            'FetchModifierDetails(myList.TempDraftDetails, myList.TempClaim, myList.TempClaimFromPaymentScreen)''need to change
            ' MapHistoryISI(myList.TempClaimFromPaymentScreen, myList.TempISI, False, GetMaxSuffix(myList.TempClaimFromPaymentScreen))
            'ReDimension(myList.Map, 0)
            ReDimension(myList.Map, 0)
            ReDimension(Map, 0)
            ' WriteLog("History ICN " & myList.ICN & " Map starts")
            'MapHistoryClaimServiceLineWithCurrentClaimServiceLineUpdated(CurrentClaimPaymentScreen, myList.TempClaimFromPaymentScreen, myList.Map, myList.ICN)
            MapHistoryClaimServiceLineWithCurrentClaimServiceLineUpdated(CurrentClaimPaymentScreen, myList.TempClaimFromPaymentScreen, myList.Map, myList.ICN)
            MapHistoryClaimServiceLineWithCurrentClaimServiceLineUpdated(CurrentClaimPaymentScreen, myList.TempClaimFromPaymentScreen, Map, myList.ICN)

            'WriteLog(GetMapInLog(myList.Map))
            'WriteLog("History ICN " & myList.ICN & " Map ends")
            GetReadyHistory(myList, myList.TempClaim, TempClaimCount)
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Sub initiateCaStruct(ByRef mystrct() As CAStruct)
        Try
            Dim I As Integer
            For I = LBound(mystrct) To UBound(mystrct)
                mystrct(I).ClaimIndex = -1
                mystrct(I).LineIndex = -1
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Sub MapHistoryClaimServiceLineWithCurrentClaimServiceLineUpdated(ByRef Current() As ClaimFromPaymentScreen, ByRef History() As ClaimFromPaymentScreen, ByRef Map() As MapCurrentProcDetailsWithHistory, ByVal HistoryICN As String)
        Dim TempMap(100) As MapCurrentProcDetailsWithHistory
        ReDimension(TempMap, 99)
        Dim GRStruct(100) As CAStruct
        Dim CountMap As Integer = 0
        Dim Count As Integer = 0
        Dim UboundCount As Integer = -1
        Try
            initiateCaStruct(GRStruct)
            For I As Integer = LBound(History) To UBound(History)
                For J As Integer = LBound(History(I).POS) To UBound(History(I).POS)
                    If History(I).Override(J).Trim = "" OrElse InStr("30,20,X0,Y0,R", History(I).Override(J).Trim) = 0 Then
                        Dim Proc As String = Left(History(I).ProcCode(J).Trim, 5)
                        Dim Nbr As String = IIf(IsNumeric(History(I).NBR(J).Trim), History(I).NBR(J).Trim, "0")
                        Dim Fdos As String = History(I).FDOS(J).Trim
                        Dim Ldos As String = History(I).LDOS(J).Trim
                        Dim Charge As String = IIf(IsNumeric(History(I).Charge(J).Trim), History(I).Charge(J).Trim, "0")
                        Count = GetUboundCaStruct(GRStruct)
                        For M = LBound(Current) To UBound(Current) 'search for exact match
                            For N = LBound(Current(M).POS) To UBound(Current(M).POS)
                                If Left(Current(M).ProcCode(N).Trim, 5) = Proc AndAlso
                                   CInt(Current(M).NBR(N).Trim) = CInt(Nbr) AndAlso
                                   Current(M).FDOS(N).Trim = Fdos AndAlso
                                   Current(M).LDOS(N).Trim = Ldos AndAlso
                                   CDbl(Current(M).Charge(N).Trim) = CDbl(Charge) Then
                                    If IsPresentInCaStruct(GRStruct, Count, M + 1, N + 1) = False Then
                                        GRStruct(Count).ClaimIndex = M + 1
                                        GRStruct(Count).LineIndex = N + 1
                                        TempMap(Count).CurrentClaimProcCode = Left(Current(M).ProcCode(N).Trim, 5)
                                        TempMap(Count).CurrentClaimFdos = Current(M).FDOS(N).Trim
                                        TempMap(Count).CurrentClaimLdos = Current(M).LDOS(N).Trim
                                        TempMap(Count).CurrentClaimUnit = Current(M).NBR(N).Replace(" ", "0").Trim
                                        TempMap(Count).CurrentClaimCharge = CDbl(Current(M).Charge(N).Trim)
                                        TempMap(Count).CurrentClaimNotCovered = CDbl(Current(M).NotCovered(N).Trim)
                                        TempMap(Count).HistoryClaimProcCode = Left(History(I).ProcCode(J).Trim, 5)
                                        TempMap(Count).HistoryClaimFdos = History(I).FDOS(J).Trim
                                        TempMap(Count).HistoryClaimLdos = History(I).LDOS(J).Trim
                                        TempMap(Count).HistoryClaimUnit = History(I).NBR(J).Trim
                                        TempMap(Count).HistoryClaimCharge = CDbl(History(I).Charge(J).Trim)
                                        TempMap(Count).HistoryClaimNotCovered = CDbl(History(I).NotCovered(J).Trim)
                                        TempMap(Count).IsCodeChange = False

                                        TempMap(Count).CurrentClaimSuffix = M + 1
                                        TempMap(Count).CurrentClaimLineNumber = N + 1

                                        TempMap(Count).HistoryClaimCIndex = I
                                        TempMap(Count).HistoryClaimLIndex = J
                                        TempMap(Count).CurrentClaimCoins = Current(M).Coins(N).Trim
                                        TempMap(Count).CurrentClaimAllowable = Current(M).Allowable(N).Trim
                                        TempMap(Count).CurrentClaimDeductible = Current(M).Deductible(N).Trim
                                        TempMap(Count).HistoryClaimAllowable = History(I).Allowable(J).Trim
                                        TempMap(Count).HistoryClaimCoins = History(I).Coins(J).Trim
                                        TempMap(Count).HistoryClaimDeductible = History(I).Deductible(J).Trim
                                        UboundCount += 1
                                        GoTo ExitFor
                                    End If

                                End If
                            Next
                        Next
ExitFor:
                        If Count = GetUboundCaStruct(GRStruct) Then 'exact code match not found, search for different code with same fdos,ldos,unit and bill charge
                            For M = LBound(Current) To UBound(Current)
                                For N = LBound(Current(M).POS) To UBound(Current(M).POS)
                                    If CInt(Current(M).NBR(N).Trim) = CInt(Nbr) AndAlso
                                       Current(M).FDOS(N).Trim = Fdos AndAlso
                                       Current(M).LDOS(N).Trim = Ldos AndAlso
                                       (IsSameConsultation(Proc, Left(Current(M).ProcCode(N), 5)) OrElse CDbl(Current(M).Charge(N).Trim) = CDbl(Charge)) Then
                                        'check in history whether this combination is exactly present or not
                                        For X As Integer = LBound(History) To UBound(History)
                                            For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                If Current(M).ProcCode(N).Trim = History(X).ProcCode(Y).Trim AndAlso
                                                   Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                   Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim AndAlso
                                                   Current(M).NBR(N).Trim = History(X).NBR(Y).Trim AndAlso
                                                   Current(M).Charge(N).Trim = History(X).Charge(Y).Trim Then
                                                    GoTo nextLineCheck
                                                End If
                                            Next
                                        Next
                                        If IsPresentInCaStruct(GRStruct, Count, M + 1, N + 1) = False Then
                                            GRStruct(Count).ClaimIndex = M + 1
                                            GRStruct(Count).LineIndex = N + 1
                                            TempMap(Count).CurrentClaimProcCode = Left(Current(M).ProcCode(N).Trim, 5)
                                            TempMap(Count).CurrentClaimFdos = Current(M).FDOS(N).Trim
                                            TempMap(Count).CurrentClaimLdos = Current(M).LDOS(N).Trim
                                            TempMap(Count).CurrentClaimUnit = Current(M).NBR(N).Replace(" ", "0").Trim
                                            TempMap(Count).CurrentClaimCharge = CDbl(Current(M).Charge(N).Trim)
                                            TempMap(Count).CurrentClaimNotCovered = CDbl(Current(M).NotCovered(N).Trim)
                                            TempMap(Count).HistoryClaimProcCode = Left(History(I).ProcCode(J).Trim, 5)
                                            TempMap(Count).HistoryClaimFdos = History(I).FDOS(J).Trim
                                            TempMap(Count).HistoryClaimLdos = History(I).LDOS(J).Trim
                                            TempMap(Count).HistoryClaimUnit = History(I).NBR(J).Trim
                                            TempMap(Count).HistoryClaimCharge = CDbl(History(I).Charge(J).Trim)
                                            TempMap(Count).HistoryClaimNotCovered = CDbl(History(I).NotCovered(J).Trim)
                                            TempMap(Count).IsCodeChange = True

                                            TempMap(Count).CurrentClaimSuffix = M + 1
                                            TempMap(Count).CurrentClaimLineNumber = N + 1

                                            TempMap(Count).HistoryClaimCIndex = I
                                            TempMap(Count).HistoryClaimLIndex = J
                                            TempMap(Count).CurrentClaimCoins = Current(M).Coins(N).Trim
                                            TempMap(Count).CurrentClaimAllowable = Current(M).Allowable(N).Trim
                                            TempMap(Count).CurrentClaimDeductible = Current(M).Deductible(N).Trim
                                            TempMap(Count).HistoryClaimAllowable = History(I).Allowable(J).Trim
                                            TempMap(Count).HistoryClaimCoins = History(I).Coins(J).Trim
                                            TempMap(Count).HistoryClaimDeductible = History(I).Deductible(J).Trim
                                            UboundCount += 1
                                            GoTo ExitFor2
                                        End If
                                    End If
nextLineCheck:
                                Next
                            Next
ExitFor2:
                            If Count = GetUboundCaStruct(GRStruct) Then 'check for bill charge change only condition
                                For M = LBound(Current) To UBound(Current)
                                    For N = LBound(Current(M).POS) To UBound(Current(M).POS)
                                        If Current(M).FDOS(N).Trim = Fdos AndAlso
                                           Current(M).LDOS(N).Trim = Ldos AndAlso
                                           (Left(Current(M).ProcCode(N), 5) = Proc OrElse IsSameConsultation(Proc, Left(Current(M).ProcCode(N), 5))) Then 'And
                                            'Current(M).ProcCode(N).Trim = Proc Then
                                            'check in history whether this combination is exactly present or not
                                            For X As Integer = LBound(History) To UBound(History)
                                                For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                    If Left(Current(M).ProcCode(N).Trim, 5) = Left(History(X).ProcCode(Y).Trim, 5) AndAlso
                                                       Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                       Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim AndAlso
                                                       Current(M).NBR(N).Trim = History(X).NBR(Y).Trim AndAlso
                                                       Current(M).Charge(N).Trim = History(X).Charge(Y).Trim Then
                                                        GoTo nextLineCheck2
                                                    End If
                                                Next
                                            Next
                                            'check in histroy whether this combination is present for proc code change condition
                                            For X As Integer = LBound(History) To UBound(History)
                                                For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                    If Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                       Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim AndAlso
                                                       Current(M).NBR(N).Trim = History(X).NBR(Y).Trim AndAlso
                                                       Current(M).Charge(N).Trim = History(X).Charge(Y).Trim Then
                                                        GoTo nextLineCheck2
                                                    End If
                                                Next
                                            Next

                                            'For X As Integer = LBound(History) To UBound(History)
                                            '    For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                            '        If Current(M).ProcCode(N).Trim = History(X).ProcCode(Y).Trim Then
                                            '            GoTo nextLineCheck2
                                            '        End If
                                            '    Next
                                            'Next
                                            If IsPresentInCaStruct(GRStruct, Count, M + 1, N + 1) = False Then
                                                GRStruct(Count).ClaimIndex = M + 1
                                                GRStruct(Count).LineIndex = N + 1
                                                TempMap(Count).CurrentClaimProcCode = Left(Current(M).ProcCode(N).Trim, 5)
                                                TempMap(Count).CurrentClaimFdos = Current(M).FDOS(N).Trim
                                                TempMap(Count).CurrentClaimLdos = Current(M).LDOS(N).Trim
                                                TempMap(Count).CurrentClaimUnit = Current(M).NBR(N).Replace(" ", "0").Trim
                                                TempMap(Count).CurrentClaimCharge = CDbl(Current(M).Charge(N).Trim)
                                                TempMap(Count).CurrentClaimNotCovered = CDbl(Current(M).NotCovered(N).Trim)
                                                TempMap(Count).HistoryClaimProcCode = Left(History(I).ProcCode(J).Trim, 5)
                                                TempMap(Count).HistoryClaimFdos = History(I).FDOS(J).Trim
                                                TempMap(Count).HistoryClaimLdos = History(I).LDOS(J).Trim
                                                TempMap(Count).HistoryClaimUnit = History(I).NBR(J).Trim
                                                TempMap(Count).HistoryClaimCharge = CDbl(History(I).Charge(J).Trim)
                                                TempMap(Count).HistoryClaimNotCovered = CDbl(History(I).NotCovered(J).Trim)
                                                TempMap(Count).IsCodeChange = True

                                                TempMap(Count).CurrentClaimSuffix = M + 1
                                                TempMap(Count).CurrentClaimLineNumber = N + 1

                                                TempMap(Count).HistoryClaimCIndex = I
                                                TempMap(Count).HistoryClaimLIndex = J
                                                TempMap(Count).CurrentClaimCoins = Current(M).Coins(N).Trim
                                                TempMap(Count).CurrentClaimAllowable = Current(M).Allowable(N).Trim
                                                TempMap(Count).CurrentClaimDeductible = Current(M).Deductible(N).Trim
                                                TempMap(Count).HistoryClaimAllowable = History(I).Allowable(J).Trim
                                                TempMap(Count).HistoryClaimCoins = History(I).Coins(J).Trim
                                                TempMap(Count).HistoryClaimDeductible = History(I).Deductible(J).Trim
                                                UboundCount += 1
                                                GoTo ExitFor3
                                            End If
                                        End If
nextLineCheck2:
                                    Next
                                Next
ExitFor3:
                                If Count = GetUboundCaStruct(GRStruct) Then 'Assuming the proc code is not changed but any combination of Fdos & ldos or Unit andalso Bill charge is changed
                                    For M = LBound(Current) To UBound(Current)
                                        For N = LBound(Current(M).POS) To UBound(Current(M).POS)
                                            If Current(M).ProcCode(N) = Proc Then
                                                'check in history whether this combination is exactly present or not
                                                For X As Integer = LBound(History) To UBound(History)
                                                    For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                        If Left(Current(M).ProcCode(N).Trim, 5) = Left(History(X).ProcCode(Y).Trim, 5) AndAlso
                                                           Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                           Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim AndAlso
                                                           Current(M).NBR(N).Trim = History(X).NBR(Y).Trim AndAlso
                                                           Current(M).Charge(N).Trim = History(X).Charge(Y).Trim Then
                                                            GoTo nextLineCheck3
                                                        End If
                                                    Next
                                                Next
                                                For X As Integer = LBound(History) To UBound(History)
                                                    For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                        If Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                           Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim AndAlso
                                                           Current(M).NBR(N).Trim = History(X).NBR(Y).Trim AndAlso
                                                           Current(M).Charge(N).Trim = History(X).Charge(Y).Trim Then
                                                            GoTo nextLineCheck3
                                                        End If
                                                    Next
                                                Next
                                                For X As Integer = LBound(History) To UBound(History)
                                                    For Y As Integer = LBound(History(X).POS) To UBound(History(X).POS)
                                                        If Current(M).FDOS(N).Trim = History(X).FDOS(Y).Trim AndAlso
                                                           Current(M).LDOS(N).Trim = History(X).LDOS(Y).Trim Then
                                                            GoTo nextLineCheck3
                                                        End If
                                                    Next
                                                Next
                                                If IsPresentInCaStruct(GRStruct, Count, M + 1, N + 1) = False Then
                                                    GRStruct(Count).ClaimIndex = M + 1
                                                    GRStruct(Count).LineIndex = N + 1
                                                    TempMap(Count).CurrentClaimProcCode = Left(Current(M).ProcCode(N).Trim, 5)
                                                    TempMap(Count).CurrentClaimFdos = Current(M).FDOS(N).Trim
                                                    TempMap(Count).CurrentClaimLdos = Current(M).LDOS(N).Trim
                                                    TempMap(Count).CurrentClaimUnit = Current(M).NBR(N).Replace(" ", "0").Trim
                                                    TempMap(Count).CurrentClaimCharge = CDbl(Current(M).Charge(N).Trim)
                                                    TempMap(Count).CurrentClaimNotCovered = CDbl(Current(M).NotCovered(N).Trim)
                                                    TempMap(Count).HistoryClaimProcCode = Left(History(I).ProcCode(J).Trim, 5)
                                                    TempMap(Count).HistoryClaimFdos = History(I).FDOS(J).Trim
                                                    TempMap(Count).HistoryClaimLdos = History(I).LDOS(J).Trim
                                                    TempMap(Count).HistoryClaimUnit = History(I).NBR(J).Trim
                                                    TempMap(Count).HistoryClaimCharge = CDbl(History(I).Charge(J).Trim)
                                                    TempMap(Count).HistoryClaimNotCovered = CDbl(History(I).NotCovered(J).Trim)
                                                    TempMap(Count).IsCodeChange = False

                                                    TempMap(Count).CurrentClaimSuffix = M + 1
                                                    TempMap(Count).CurrentClaimLineNumber = N + 1

                                                    TempMap(Count).HistoryClaimCIndex = I
                                                    TempMap(Count).HistoryClaimLIndex = J
                                                    TempMap(Count).CurrentClaimCoins = Current(M).Coins(N).Trim
                                                    TempMap(Count).CurrentClaimAllowable = Current(M).Allowable(N).Trim
                                                    TempMap(Count).CurrentClaimDeductible = Current(M).Deductible(N).Trim
                                                    TempMap(Count).HistoryClaimAllowable = History(I).Allowable(J).Trim
                                                    TempMap(Count).HistoryClaimCoins = History(I).Coins(J).Trim
                                                    TempMap(Count).HistoryClaimDeductible = History(I).Deductible(J).Trim
                                                    UboundCount += 1
                                                    GoTo ExitFor4
                                                End If
                                            End If
nextLineCheck3:
                                        Next
                                    Next
ExitFor4:
                                End If
                            End If
                        End If

                        If Count = GetUboundCaStruct(GRStruct) Then

                            GRStruct(Count).ClaimIndex = 500
                            GRStruct(Count).LineIndex = 500
                            TempMap(Count).CurrentClaimProcCode = ""
                            TempMap(Count).CurrentClaimFdos = ""
                            TempMap(Count).CurrentClaimLdos = ""
                            TempMap(Count).CurrentClaimUnit = "0"
                            TempMap(Count).CurrentClaimCharge = 0
                            TempMap(Count).CurrentClaimNotCovered = 0
                            TempMap(Count).HistoryClaimProcCode = Left(History(I).ProcCode(J).Trim, 5)
                            TempMap(Count).HistoryClaimFdos = History(I).FDOS(J).Trim
                            TempMap(Count).HistoryClaimLdos = History(I).LDOS(J).Trim
                            TempMap(Count).HistoryClaimUnit = History(I).NBR(J).Trim
                            TempMap(Count).HistoryClaimCharge = IIf(IsNumeric(History(I).Charge(J).Trim), CDbl(History(I).Charge(J).Trim), 0)
                            TempMap(Count).HistoryClaimNotCovered = IIf(IsNumeric(History(I).NotCovered(J).Trim), CDbl(History(I).NotCovered(J).Trim), 0)
                            TempMap(Count).IsCodeChange = True

                            TempMap(Count).CurrentClaimSuffix = -1
                            TempMap(Count).CurrentClaimLineNumber = -1

                            TempMap(Count).HistoryClaimCIndex = I
                            TempMap(Count).HistoryClaimLIndex = J

                            TempMap(Count).HistoryClaimAllowable = History(I).Allowable(J).Trim
                            TempMap(Count).HistoryClaimCoins = History(I).Coins(J).Trim
                            TempMap(Count).HistoryClaimDeductible = History(I).Deductible(J).Trim

                            UboundCount += 1
                        End If
                    End If
                Next
            Next
            Count = GetUboundCaStruct(GRStruct)
            For M = LBound(Current) To UBound(Current)
                For N = LBound(Current(M).POS) To UBound(Current(M).POS)
                    CountMap = 0
                    Dim Flag = False
                    While CountMap < Count
                        If TempMap(CountMap).CurrentClaimSuffix = M + 1 AndAlso TempMap(CountMap).CurrentClaimLineNumber = N + 1 Then
                            Flag = True
                            Exit While
                        End If
                        CountMap += 1
                    End While
                    If Flag = False Then
                        TempMap(Count).CurrentClaimProcCode = Left(Current(M).ProcCode(N).Trim, 5)
                        TempMap(Count).CurrentClaimFdos = Current(M).FDOS(N).Trim
                        TempMap(Count).CurrentClaimLdos = Current(M).LDOS(N).Trim
                        TempMap(Count).CurrentClaimUnit = Current(M).NBR(N).Replace(" ", "0").Trim
                        TempMap(Count).CurrentClaimCharge = CDbl(Current(M).Charge(N).Trim)
                        TempMap(Count).CurrentClaimNotCovered = CDbl(Current(M).NotCovered(N).Trim)
                        TempMap(Count).HistoryClaimProcCode = ""
                        TempMap(Count).HistoryClaimFdos = ""
                        TempMap(Count).HistoryClaimLdos = ""
                        TempMap(Count).HistoryClaimUnit = "0"
                        TempMap(Count).HistoryClaimCharge = 0
                        TempMap(Count).HistoryClaimNotCovered = 0
                        TempMap(Count).IsCodeChange = True

                        TempMap(Count).CurrentClaimSuffix = M + 1
                        TempMap(Count).CurrentClaimLineNumber = N + 1

                        TempMap(Count).HistoryClaimCIndex = -1
                        TempMap(Count).HistoryClaimLIndex = -1
                        TempMap(Count).CurrentClaimCoins = Current(M).Coins(N).Trim
                        TempMap(Count).CurrentClaimAllowable = Current(M).Allowable(N).Trim
                        TempMap(Count).CurrentClaimDeductible = Current(M).Deductible(N).Trim
                        UboundCount += 1
                        Count += 1
                    End If
                Next
            Next
            ReDimension(Map, UboundCount)
            CountMap = 0
            While CountMap <= UboundCount
                Map(CountMap).CurrentClaimProcCode = IIf(TempMap(CountMap).CurrentClaimProcCode = "NCOPA", "NCOPAY", TempMap(CountMap).CurrentClaimProcCode)
                Map(CountMap).CurrentClaimFdos = TempMap(CountMap).CurrentClaimFdos
                Map(CountMap).CurrentClaimLdos = TempMap(CountMap).CurrentClaimLdos
                Map(CountMap).CurrentClaimUnit = TempMap(CountMap).CurrentClaimUnit
                Map(CountMap).CurrentClaimCharge = TempMap(CountMap).CurrentClaimCharge
                Map(CountMap).CurrentClaimNotCovered = TempMap(CountMap).CurrentClaimNotCovered


                Map(CountMap).HistoryClaimProcCode = IIf(TempMap(CountMap).HistoryClaimProcCode = "NCOPA", "NCOPAY", TempMap(CountMap).HistoryClaimProcCode)
                Map(CountMap).HistoryClaimFdos = TempMap(CountMap).HistoryClaimFdos
                Map(CountMap).HistoryClaimLdos = TempMap(CountMap).HistoryClaimLdos
                Map(CountMap).HistoryClaimUnit = TempMap(CountMap).HistoryClaimUnit
                Map(CountMap).HistoryClaimCharge = TempMap(CountMap).HistoryClaimCharge
                Map(CountMap).HistoryClaimNotCovered = TempMap(CountMap).HistoryClaimNotCovered


                Map(CountMap).IsCodeChange = Not (TempMap(CountMap).CurrentClaimProcCode = TempMap(CountMap).HistoryClaimProcCode)

                Map(CountMap).CurrentClaimSuffix = IIf(TempMap(CountMap).CurrentClaimSuffix = 500, -1, TempMap(CountMap).CurrentClaimSuffix)
                Map(CountMap).CurrentClaimLineNumber = IIf(TempMap(CountMap).CurrentClaimLineNumber = 500, -1, TempMap(CountMap).CurrentClaimLineNumber)

                Map(CountMap).HistoryClaimCIndex = TempMap(CountMap).HistoryClaimCIndex
                Map(CountMap).HistoryClaimLIndex = TempMap(CountMap).HistoryClaimLIndex
                Map(CountMap).CurrentClaimCoins = TempMap(CountMap).CurrentClaimCoins
                Map(CountMap).HistoryClaimCoins = TempMap(CountMap).HistoryClaimCoins
                Map(CountMap).CurrentClaimAllowable = TempMap(CountMap).CurrentClaimAllowable
                Map(CountMap).HistoryClaimAllowable = TempMap(CountMap).HistoryClaimAllowable
                Map(CountMap).CurrentClaimDeductible = TempMap(CountMap).CurrentClaimDeductible
                Map(CountMap).HistoryClaimDeductible = TempMap(CountMap).HistoryClaimDeductible
                Map(CountMap).HXICN = HistoryICN
                CountMap += 1
            End While
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

            'WriteLog(GetCurrentMethod().Name)
            'WriteLog(ex.ToString)
        End Try
    End Sub
    Public Function IsSameConsultation(ByVal Proc As String, ByVal Proc2 As String) As Boolean
        If Proc Is Nothing OrElse Proc2 Is Nothing Then Return False
        'DOMICILLIARY -> 99324,99325,99326,99327,99328,99334,99335,99336,99337
        If InStr("99324,99325,99326,99327,99328,99334,99335,99336,99337", Proc) > 0 AndAlso InStr("99324,99325,99326,99327,99328,99334,99335,99336,99337", Proc2) > 0 Then Return True
        'EMERGENCY -> 99281,99282,99283,99284,99285
        If InStr("99281,99282,99283,99284,99285", Proc) > 0 AndAlso InStr("99281,99282,99283,99284,99285", Proc2) > 0 Then Return True
        'HOME VISIT	-> 99341,99342,99343,99344,99345,99347,99348,99349,99350
        If InStr("99341,99342,99343,99344,99345,99347,99348,99349,99350", Proc) > 0 AndAlso InStr("99341,99342,99343,99344,99345,99347,99348,99349,99350", Proc2) > 0 Then Return True
        'HOSPITAL CARE -> 99231,99232,99233
        If InStr("99231,99232,99233", Proc) > 0 AndAlso InStr("99231,99232,99233", Proc2) > 0 Then Return True
        'HOSPITAL DISCHARGE -> 99238,99239
        If InStr("99238,99239", Proc) > 0 AndAlso InStr("99238,99239", Proc2) > 0 Then Return True
        'IND PHY SUPERVISION -> 99339,99340
        If InStr("99339,99340", Proc) > 0 AndAlso InStr("99339,99340", Proc2) > 0 Then Return True
        'INPATIENT CONSULTATION -> 99251,99252,99253,99254,99255
        If InStr("99251,99252,99253,99254,99255", Proc) > 0 AndAlso InStr("99251,99252,99253,99254,99255", Proc2) > 0 Then Return True
        'NURSING FACILITY -> 99304,99305,99306,99307,99308,99309,99310
        If InStr("99304,99305,99306,99307,99308,99309,99310", Proc) > 0 AndAlso InStr("99304,99305,99306,99307,99308,99309,99310", Proc2) > 0 Then Return True
        'NURSING FACILITY DISCHARGE -> 99315,99316
        If InStr("99315,99316", Proc) > 0 AndAlso InStr("99315,99316", Proc2) > 0 Then Return True
        'OBSERVATION -> 99217,99218,99219,99220,99221,99222,99223,99224,99225,99226
        If InStr("99217,99218,99219,99220,99221,99222,99223,99224,99225,99226", Proc) > 0 AndAlso InStr("99217,99218,99219,99220,99221,99222,99223,99224,99225,99226", Proc2) > 0 Then Return True
        'OBSERVATION AND INPATIENT HOSPITAL CARE -> 99234,99235,99236
        If InStr("99234,99235,99236", Proc) > 0 AndAlso InStr("99234,99235,99236", Proc2) > 0 Then Return True
        'OFFICE CONSULTATION -> 99241,99242,99243,99244,99245
        If InStr("99241,99242,99243,99244,99245", Proc) > 0 AndAlso InStr("99241,99242,99243,99244,99245", Proc2) > 0 Then Return True
        'OFFICE OR OUTPATIENT VISIT -> 99201,99202,99203,99204,99205,99211,99212,99213,99214,99215
        If InStr("99201,99202,99203,99204,99205,99211,99212,99213,99214,99215", Proc) > 0 AndAlso InStr("99201,99202,99203,99204,99205,99211,99212,99213,99214,99215", Proc2) > 0 Then Return True
        'Merge the above two codes for sample ICN 7498764993
        If InStr("99201,99202,99203,99204,99205,99211,99212,99213,99214,99215,99241,99242,99243,99244,99245", Proc) > 0 AndAlso InStr("99201,99202,99203,99204,99205,99211,99212,99213,99214,99215,99241,99242,99243,99244,99245", Proc2) > 0 Then Return True
        Return False

    End Function
    Public Function IsPresentInCaStruct(ByRef myStrct() As CAStruct, ByVal count As Integer, ByVal cIndex As Integer, ByVal lIndex As Integer) As Boolean
        Dim I As Integer
        Try
            For I = 0 To count - 1
                If myStrct(I).ClaimIndex = cIndex And myStrct(I).LineIndex = lIndex Then Return True
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return False
    End Function
    Public Function GetUboundCaStruct(ByRef mystrct() As CAStruct) As Integer
        Dim I As Integer
        Try
            For I = LBound(mystrct) To UBound(mystrct)
                If mystrct(I).ClaimIndex = -1 And mystrct(I).LineIndex = -1 Then Return I
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return I
    End Function
    Public Function GetMaxSuffix(ByRef c() As Claim, ByVal Count As Integer) As Integer
        Dim I, J As Integer
        J = 0
        Try
            For I = 0 To Count - 1
                If CInt(c(I).Suffix) > J Then J = CInt(c(I).Suffix)
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return J
    End Function
    Public Function GetDraftDetailsInMemory(ByRef Draft() As DraftDetails, ByRef myClaim() As Claim, ByVal Count As String) As Integer
        Dim DraftList As String = ""
        Dim DraftList2 As String = ""
        Dim J As Integer
        Dim FlagROV As Boolean = False
        Try
            For I As Integer = 0 To Count - 1
                If InStr(DraftList, myClaim(I).Draft) = 0 And CDbl(myClaim(I).Draft) <> 0 Then
                    FlagROV = False
                    For J = LBound(myClaim(I).ServiceLine) To UBound(myClaim(I).ServiceLine) - 1
                        If myClaim(I).OV(J) = "R" And myClaim(I).RemarkCode(J) = "OV" Then
                            FlagROV = True
                            Exit For
                        ElseIf myClaim(I).OV(J) = "20" Then
                            FlagROV = True
                            Exit For
                        End If
                    Next
                    If FlagROV = False Then DraftList += myClaim(I).Draft + ","
                End If
            Next

            J = UBound(DraftList.Split(",")) - 1
            ReDimension(Draft, J)
            DraftList2 = ""
            J = 0
            For I As Integer = 0 To Count - 1
                If (DraftList2 = "" OrElse InStr(DraftList2, myClaim(I).Draft) = 0) And CDbl(myClaim(I).Draft) <> 0 And Split(DraftList, ",")(J) = myClaim(I).Draft Then
                    DraftList2 += myClaim(I).Draft + ","
                    Draft(J).Draft = myClaim(I).Draft
                    Draft(J).Suffix = myClaim(I).Suffix
                    If myClaim(I).Closed(0) = "1" Then
                        Draft(J).isClosed = True
                        Draft(J).ClosureCode = myClaim(I).RemarkCode(0)
                    Else
                        Draft(J).isClosed = False

                        Draft(J).ClosureCode = ""
                    End If
                    Draft(J).TotalAllowable = CDbl(myClaim(I).TotalAllowable)
                    Draft(J).TotalBillCharge = CDbl(myClaim(I).BillCharge)
                    Draft(J).TotalNotCovered = CDbl(myClaim(I).TotalNotCovered)
                    Draft(J).TotalPaid = CDbl(myClaim(I).PaidAmount)
                    Draft(J).DateOfProcessing = myClaim(I).DateOfProcessing
                    Draft(J).DateOfReversal = ""
                    Draft(J).Reversal = False
                    Dim Flag As Boolean = False
                    For K = 0 To Count - 1
                        If myClaim(I).Draft = myClaim(K).Draft And myClaim(K).Voided Then
                            Draft(J).Reversal = True
                            Draft(J).DateOfReversal = myClaim(K).DateOfProcessing
                            If myClaim(K).RemarkCode(0) = "69" Then
                                Draft(J).ReversalType = "69-70" : Flag = True
                            End If
                            If myClaim(K).RemarkCode(0) = "74" Then
                                Draft(J).ReversalType = "74" : Flag = True
                            End If
                            If myClaim(K).RemarkCode(0) = "71" And myClaim(K).OV(0).Trim = "R" Then
                                Draft(J).ReversalType = "R-71" : Flag = False : Draft(J).Reversal = False
                            ElseIf myClaim(K).RemarkCode(0) = "71" And myClaim(K).OV(0).Trim = "P" Then
                                Draft(J).ReversalType = "71" : Flag = False : Draft(J).Reversal = False
                            ElseIf myClaim(K).RemarkCode(0) = "71" Then
                                Draft(J).ReversalType = "71 Void" : Flag = True
                            End If

                            Exit For
                        End If
                    Next
                    If Flag = False Then
                        Draft(J).Reversal = False
                        Draft(J).ReversalType = "Not Voided"
                    End If
                    J += 1
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return J
    End Function
    Public Sub FetchModifierDetails(ByRef MyDetails() As DraftDetails, ByRef Myclaim() As Claim, ByRef MyPaymentScreen() As ClaimFromPaymentScreen)
        '//Get from MHI 'Special GHI in hx icn 7502920917 for current icn 7580874713 295E
        For I As Integer = 1 To 1 'GetMaxSuffix(MyPaymentScreen)
            For J As Integer = LBound(MyDetails) To UBound(MyDetails)
                If CInt(MyDetails(J).Suffix) = I And MyDetails(J).isClosed = False Then
                    Dim Indx As Integer = GetIndexInClaimByDraftNumber(Myclaim, MyDetails(J).Draft)
                    If Indx >= 0 Then
                        PullGHI(Myclaim, MyDetails(J).Draft)
                        'ReDim Myclaim(Indx).Modifier(UBound(Myclaim(Indx).ServiceLine))
                        'ReDim MyPaymentScreen(CInt(Myclaim(Indx).Suffix) - 1).Modifier(UBound(MyPaymentScreen(CInt(Myclaim(Indx).Suffix) - 1).POS))
                        Dim Counter As Integer = 0
                        While 1
                            autECLPSObj.auteclfieldlist.Refresh
                            Dim FindObject = autECLPSObj.auteclfieldlist.FindFieldByRowCol(13, 1)
                            If FindObject.Display And InStr("69,70,71,74,77,75,42,87,44,80,81,82,OV,AC", GetText(14, 50, 2, True)) = 0 Then
                                Myclaim(Indx).Modifier(Counter) = GetText(14, 11, 11, True).Replace(" ", ",")
                                MyPaymentScreen(CInt(Myclaim(Indx).Suffix) - 1).Modifier(Counter) = GetText(14, 11, 11, True).Replace(" ", ",")
                                Counter += 1
                            End If
                            If InStr(GetText(24, 1, 80, True), "E065NO MORE CLMS ON FILE") OrElse InStr(GetText(24, 1, 80, True), "W062NO CLAIMS ON FILE") Then 'W062NO CLAIMS ON FILE
                                Exit While
                            Else

                                Dim MyScreen As String = GetTextRect(13, 1, 19, 79)
                                Dim Count As Integer = 0
                                Do
                                    Count += 1
                                    SendEnter()
                                    autECLPSObj.auteclfieldlist.Refresh
                                    'Dim FindObject = frmCorrectedMain.autECLPSObj.auteclfieldlist.FindFieldByRowCol(13, 1)
                                    'MsgBox(frmCorrectedMain.GetText(24, 1, 80, True))
                                    If Count > 100 Then
                                        strStatus = "skipped"
                                        strResearchComments = "Problem in ghi"
                                        ExitAllAndReturnWithException()
                                    End If
                                Loop While InStr(GetText(24, 1, 80, True), "E065NO MORE CLMS ON FILE") = 0 AndAlso InStr(GetText(24, 1, 80, True), "W062NO CLAIMS ON FILE") = 0 AndAlso GetTextRect(13, 1, 19, 79) = MyScreen 'W062NO CLAIMS ON FILE


                            End If
                        End While
                    End If
                    Exit For
                End If
            Next
        Next
    End Sub
    Sub PullGHI(ByRef MyClaim() As Claim, ByVal Draft As String)
        Try
            SendClear()
            SendClear()
            SendClear()

            SendKeyes("GHI," + MyClaim(0).Policy + "," + MyClaim(0).SSN + "," + ConvertToFormat(MyClaim(0).EmployeeName, 10,, " ") + "," + MyClaim(0).Relationship + "," + Draft + ",I" + MyClaim(0).ICN, 2, 2)
            SendEnter()

            If GetText(2, 2, 2) <> "GH" Then
                strStatus = "SKIPPED"
                strResearchComments = "unable to pull ghi"
                ExitAllAndReturnWithException()
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Public Function GetIndexInClaimByDraftNumber(ByRef MyClaim() As Claim, ByVal Draft As String) As Integer
        For I As Integer = LBound(MyClaim) To UBound(MyClaim) - 1
            If MyClaim(I).Draft = Draft And MyClaim(I).Voided = False And (MyClaim(I).RemarkCode(0).Trim = "" OrElse InStr("77,75,OV,42,87,AC", MyClaim(I).RemarkCode(0)) = 0) Then Return I
        Next
        Return -999
    End Function
    Public Sub GetReadyHistory(ByRef ListOfPossibleICN As PossibleICN, ByRef MyClaim() As Claim, ByVal J As Integer)
        'ListOfPossibleICN(IndexCount).ProcCodeMatchCount = ProcCodeMatchCount
        'ListOfPossibleICN(IndexCount).ICN = AHIICN
        Try
            Dim sxList = ""
            Dim drftList = ""

            ValidDraftAndSuffix(sxList, drftList, MyClaim, J)
            Dim KK As Integer = 0
            ListOfPossibleICN.IsClosed = True
            ListOfPossibleICN.E9or12Closed = False
            ListOfPossibleICN.isPaid = False
            ListOfPossibleICN.TotalBillCharge = "0"
            ListOfPossibleICN.TotalPaid = "0"
            While KK < J
                If InStr(drftList, MyClaim(KK).Draft) > 0 And InStr(sxList, MyClaim(KK).Suffix) Then
                    If MyClaim(KK).Voided = False Or (MyClaim(KK).RemarkCode(0) = "74" And MyClaim(KK).Voided = True) Then

                        ListOfPossibleICN.TotalBillCharge = CStr(CDbl(ListOfPossibleICN.TotalBillCharge) + CDbl(MyClaim(KK).BillCharge))
                        If MyClaim(KK).OV(0).Trim = "C" Then
                            ListOfPossibleICN.IsClosed = True And ListOfPossibleICN.IsClosed
                            If InStr("E9,12", MyClaim(KK).RemarkCode(0)) > 0 Then
                                ListOfPossibleICN.E9or12Closed = True Or ListOfPossibleICN.E9or12Closed
                            Else
                                ListOfPossibleICN.E9or12Closed = False Or ListOfPossibleICN.E9or12Closed
                            End If
                        Else
                            ListOfPossibleICN.IsClosed = False And ListOfPossibleICN.IsClosed
                            If CDbl(MyClaim(KK).TotalNotCovered) = CDbl(MyClaim(KK).BillCharge) Then
                                ListOfPossibleICN.isPaid = False Or ListOfPossibleICN.isPaid
                            Else
                                ListOfPossibleICN.TotalPaid = CStr(CDbl(ListOfPossibleICN.TotalPaid) + CDbl(MyClaim(KK).PaidAmount))
                                ListOfPossibleICN.isPaid = True Or ListOfPossibleICN.isPaid
                            End If
                        End If
                    End If
                End If
                KK += 1
            End While
            KK = 0
            ListOfPossibleICN.isVoided = False
            ListOfPossibleICN.isPartiallyVoided = True
            While KK < J
                If InStr(drftList, MyClaim(KK).Draft) > 0 And InStr(sxList, MyClaim(KK).Suffix) Then
                    If MyClaim(KK).Voided = True Then
                        ListOfPossibleICN.isVoided = True Or ListOfPossibleICN.isVoided
                        ListOfPossibleICN.isPartiallyVoided = True And ListOfPossibleICN.isPartiallyVoided
                    Else
                        ListOfPossibleICN.isVoided = False Or ListOfPossibleICN.isVoided
                        ListOfPossibleICN.isPartiallyVoided = False And ListOfPossibleICN.isPartiallyVoided
                    End If
                End If
                KK += 1
            End While
            If ListOfPossibleICN.isVoided = True Then ListOfPossibleICN.isPartiallyVoided = False
            ListOfPossibleICN.FLN = MyClaim(0).FLN
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Sub ValidDraftAndSuffix(ByRef SuffixList As String, ByRef ValidDraftList As String, ByRef c() As Claim, ByVal count As Integer)
        Try
            Dim TempDraft1 As String
            For i = 0 To count - 1
                If InStr(SuffixList, c(i).Suffix) = 0 And Trim(c(i).Suffix) <> "" Then
                    SuffixList = SuffixList & c(i).Suffix & ","
                End If
            Next
            If Len(SuffixList) > 0 And Microsoft.VisualBasic.Right(SuffixList, 1) = "," Then SuffixList = Microsoft.VisualBasic.Left(SuffixList, Len(SuffixList) - 1)
            ValidDraftList = ""
            For Each Suffx As String In Split(SuffixList, ",")
                TempDraft1 = ""
                For i = count - 1 To 0 Step -1
                    If c(i).Suffix = Suffx Then
                        If TempDraft1 <> c(i).Draft And c(i).Closed(0) <> "1" Then
                            TempDraft1 = c(i).Draft
                        End If
                    End If
                Next
                ValidDraftList = ValidDraftList & TempDraft1 & ","
            Next Suffx
            While Microsoft.VisualBasic.Right(ValidDraftList, 1) = ","
                ValidDraftList = Microsoft.VisualBasic.Left(ValidDraftList, Len(ValidDraftList) - 1)
            End While
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Sub PasteHxICNAndDraft(ByVal Amount As String, ByVal Paid As String)
        'If is30LineToBeused = False Then Exit Sub
        Try
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If

            For J As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                If ListOfPossibleICNActual(J).isPaid Then 'AndAlso Mid(ListOfPossibleICNActual(J).TempClaim(0).Tin, 3, 9) = Mid(CurrentClaim(0).Tin, 3, 9) 
                    For K As Integer = LBound(ListOfPossibleICNActual(J).TempDraftDetails) To UBound(ListOfPossibleICNActual(J).TempDraftDetails)
                        apiChk()
                        If ListOfPossibleICNActual(J).TempDraftDetails(K).Draft <> "" AndAlso CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft) > 0 AndAlso ListOfPossibleICNActual(J).TempDraftDetails(K).isClosed = False And InStr("R-71,69-70", ListOfPossibleICNActual(J).TempDraftDetails(K).ReversalType) > 0 And CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).TotalAllowable) <> CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).TotalNotCovered) Then
                            SendKeyes(ListOfPossibleICNActual(J).TempClaim(0).ICN, 4, 15)
                            SendKeyes(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft, 4, 32)
                            GoTo OutOfEdsHalfScreen
                        End If
                    Next
                End If
            Next

            For J As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                For K As Integer = LBound(ListOfPossibleICNActual(J).TempDraftDetails) To UBound(ListOfPossibleICNActual(J).TempDraftDetails)
                    apiChk()
                    If ListOfPossibleICNActual(J).TempDraftDetails(K).Draft <> "" AndAlso CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft) > 0 AndAlso ListOfPossibleICNActual(J).TempDraftDetails(K).isClosed = False AndAlso InStr("R-71,69-70", ListOfPossibleICNActual(J).TempDraftDetails(K).ReversalType) > 0 AndAlso Mid(ListOfPossibleICNActual(J).TempClaim(0).Tin, 3, 9) = Mid(CurrentClaim(0).Tin, 3, 9) Then
                        SendKeyes(ListOfPossibleICNActual(J).TempClaim(0).ICN, 4, 15)
                        SendKeyes(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft, 4, 32)
                        GoTo OutOfEdsHalfScreen
                    End If
                Next
            Next
            SendKeyes("----------", 4, 15)
            SendKeyes("----------", 4, 32)


            'Else
            '    For I As Integer = 0 To HistoryClaimCount
            '        If CDbl(HistoryClaim(I).Draft) <> 0 AndAlso HistoryClaim(I).Voided = False Then
            '            SendKeyes(HistoryClaim(0).ICN, 4, 15)
            '            SendKeyes(HistoryClaim(I).Draft, 4, 32)
            '            Exit For
            '        End If
            '    Next
            'End If
OutOfEdsHalfScreen:
            SendEnter()
            Dim text As String = String.Empty
            text = GetSixPointFiveScreen()
            WriteLog(text)
            WriteLog(GetTextRect)
            SendKeyes("A", 1, 75)
            'SendEnter()
            autECLPSObj.SendKeys("[ENTER]")
            apiChk()
            WriteLog(GetTextRect)
            If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                SendKeyes("N", 4, 79)
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
            End If
            'SendKeyes("N", 4, 79)
            'SendEnter()
            'SendKeyes("A", 1, 75)
            'SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                ExitAllAndReturnWithException()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub
    Function DetermineIfCAApplicable(ByRef Map() As MapCurrentProcDetailsWithHistory, ByRef myClaim() As ClaimFromPaymentScreen, ByRef iSRCCA As Integer, ByVal HxCauseCode As String, ByVal HxICN As String) As Boolean
        Dim currentClaimCopayAmount As Double = 0
        Dim historyClaimCopayAmount As Double = 0
        Dim TempClaim(100) As Claim, TempClaimCount As Integer = 0
        Dim StoreLines(100) As CAStruct, count As Integer = 0
        Try

            'need to check for ncopay lines in map, assuming that ncopay lines are mapped correctly in map structure

            'Corrected claim has NO difference or change in Code, Allowable, Coins, Deductible and Paid compared to history
            'For I As Integer = LBound(Map) To UBound(Map)
            '    If Map(I).IsCodeChange Then Return False
            '    If CDbl(Map(I).CurrentClaimAllowable) <> CDbl(Map(I).HistoryClaimAllowable) OrElse
            '       CDbl(Map(I).CurrentClaimDeductible) <> CDbl(Map(I).HistoryClaimDeductible) OrElse
            '       CDbl(Map(I).CurrentClaimCoins) <> CDbl(Map(I).HistoryClaimCoins) OrElse
            '       CDbl(Map(I).CurrentClaimCharge) <> CDbl(Map(I).HistoryClaimCharge) Then
            '        Return False
            '    End If
            'Next
            'Need to check for Paid or ZeroPaid claim in history for becoming eligible for CA denial
            Dim CAEligibilityFlag As Boolean = False
            For I As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                For J As Integer = LBound(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen) To UBound(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen)
                    If CDbl(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).TotalPaid) > 0 Then
                        CAEligibilityFlag = True
                        Exit For
                    Else
                        For K As Integer = LBound(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).POS) To UBound(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).POS)
                            If InStr(InvalidProcCode, Left(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).ProcCode(K), 5)) = 0 Then
                                If CDbl(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).Charge(K)) <> CDbl(ListOfPossibleICNActual(I).TempClaimFromPaymentScreen(J).Charge(K)) Then
                                    CAEligibilityFlag = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                If CAEligibilityFlag Then Exit For
            Next
            If CAEligibilityFlag = False Then Return False
            Dim Flag As Boolean
            For I As Integer = LBound(myClaim) To UBound(myClaim)
                Flag = False
                If myClaim(I).CauseCode = "4" Then 'Also if the suffix is not a denied suffix then
                    Dim BC As Double = 0, NC As Double = 0
                    For J As Integer = LBound(myClaim(I).POS) To UBound(myClaim(I).POS)
                        BC = System.Math.Round(CDbl(myClaim(I).Charge(J)) + BC, 2)
                        NC = System.Math.Round(CDbl(myClaim(I).NotCovered(J)) + NC, 2)
                    Next
                    If BC <> NC Then
                        For J As Integer = LBound(myClaim(I).POS) To UBound(myClaim(I).POS)
                            For K As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                                Dim CurrentSuffix As Integer = FindIndexInMapByClaimIndexAndLineIndex(ListOfPossibleICNActual(K).Map, I + 1, J + 1, True)
                                If CurrentSuffix >= 0 Then
                                    If ListOfPossibleICNActual(K).Map(CurrentSuffix).HistoryClaimCIndex >= 0 AndAlso ListOfPossibleICNActual(K).TempClaim(ListOfPossibleICNActual(K).Map(CurrentSuffix).HistoryClaimCIndex).CauseCode = "4" Then
                                        If CDbl(ListOfPossibleICNActual(K).TempClaim(ListOfPossibleICNActual(K).Map(CurrentSuffix).HistoryClaimCIndex).BillCharge) <> CDbl(ListOfPossibleICNActual(K).TempClaim(ListOfPossibleICNActual(K).Map(CurrentSuffix).HistoryClaimCIndex).TotalNotCovered) Then
                                            'Ex: 7619045507	M	117
                                            Flag = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                            If Flag Then Exit For
                        Next
                        If Not Flag Then
                            'WriteLog("Current claim cause code 4 and history claim cause code 0, hence ca not applicable.")
                            Return False
                        End If
                    End If
                End If
            Next


            Dim IndextoCheck As String = ""

            For J As Integer = LBound(Map) To UBound(Map)
                If InStr(Map(J).HistoryClaimProcCode, "NCOPAY") = 0 Then
                    Flag = False
                    Dim Lindex As Integer = Map(J).CurrentClaimLineNumber
                    Dim Cindex As Integer = Map(J).CurrentClaimSuffix
                    If Cindex >= 0 And Lindex >= 0 Then
                        For K As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                            For L As Integer = LBound(ListOfPossibleICNActual(K).Map) To UBound(ListOfPossibleICNActual(K).Map)
                                'WriteLog("PASS " & J + 1 & ":" & "Matching History ICN " & ListOfPossibleICNActual(K).ICN & "for Map PASS " & L + 1)
                                'WriteLog("ListOfPossibleICNActual(" & K & ").Map(" & L & ").CurrentClaimLineNumber: " & ListOfPossibleICNActual(K).Map(L).CurrentClaimLineNumber & " = " & "Lindex: " & Lindex & " and " & "ListOfPossibleICNActual(" & K & ").Map(" & L & ").CurrentClaimSuffix: " & ListOfPossibleICNActual(K).Map(L).CurrentClaimSuffix & " = " & "Cindex: " & Cindex & " Evaluates to " & (ListOfPossibleICNActual(K).Map(L).CurrentClaimLineNumber = Lindex And ListOfPossibleICNActual(K).Map(L).CurrentClaimSuffix = Cindex))
                                If ListOfPossibleICNActual(K).Map(L).CurrentClaimLineNumber = Lindex And ListOfPossibleICNActual(K).Map(L).CurrentClaimSuffix = Cindex Then

                                    If ListOfPossibleICNActual(K).Map(L).HistoryClaimCharge = Map(J).CurrentClaimCharge OrElse (CDbl(ListOfPossibleICNActual(K).Map(L).HistoryClaimCharge) = 0 And CDbl(Map(J).CurrentClaimCharge) = 0.01) OrElse (CDbl(ListOfPossibleICNActual(K).Map(L).HistoryClaimCharge) = 0.01 And CDbl(Map(J).CurrentClaimCharge) = 0) Then
                                        Dim x As Integer = ListOfPossibleICNActual(K).TempClaimFromPaymentScreen(ListOfPossibleICNActual(K).Map(L).HistoryClaimCIndex).HxCindex(ListOfPossibleICNActual(K).Map(L).HistoryClaimLIndex)
                                        Dim y As Integer = ListOfPossibleICNActual(K).TempClaimFromPaymentScreen(ListOfPossibleICNActual(K).Map(L).HistoryClaimCIndex).HxLIndex(ListOfPossibleICNActual(K).Map(L).HistoryClaimLIndex)
                                        If ListOfPossibleICNActual(K).TempClaim(x).OV(y) <> "C" AndAlso InStr("B5,B6,HW,GZ,05,TQ", ListOfPossibleICNActual(K).TempClaim(x).RemarkCode(y)) = 0 AndAlso (CDbl(ListOfPossibleICNActual(K).Map(L).HistoryClaimAllowable) = CDbl(Map(J).CurrentClaimAllowable) OrElse
                                            System.Math.Round(CDbl(ListOfPossibleICNActual(K).TempClaim(x).Copay(y) + CDbl(ListOfPossibleICNActual(K).TempClaim(x).Allowable(y))), 2) = CDbl(Map(J).CurrentClaimAllowable)) Then
                                            If CDbl(Map(J).CurrentClaimCharge) = CDbl(Map(J).CurrentClaimNotCovered) OrElse CDbl(Map(J).CurrentClaimAllowable) = CDbl(Map(J).CurrentClaimDeductible) OrElse
                                                CDbl(ListOfPossibleICNActual(K).Map(L).HistoryClaimCoins) = CDbl(Map(J).CurrentClaimCoins) OrElse CInt(Map(J).CurrentClaimCoins) = 100 Then
                                                Flag = True
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next
                            If Flag Then
                                'WriteLog("Mismatch")
                                Exit For
                            End If
                        Next
                        If Flag = False Then Return False
                    End If
                End If
            Next
            'check for Ncopay combined amount
            'For I As Integer = LBound(Map) To UBound(Map)
            '    If Map(I).CurrentClaimProcCode = "NCOPAY" Then
            '        currentClaimCopayAmount += Map(I).CurrentClaimNotCovered
            '    End If
            '    If Map(I).CurrentClaimProcCode = "NCOPAY" Then
            '        historyClaimCopayAmount += Map(I).HistoryClaimNotCovered
            '    End If
            'Next
            'If currentClaimCopayAmount <> historyClaimCopayAmount Then Return False
            'For I As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
            '    PullMHI(ListOfPossibleICNActual(I).ICN)
            '    TempClaimCount = ReadMHI(TempClaim)
            '    GetDraftDetailsInMemory(TempDraftDetails, TempClaim, TempClaimCount)
            '    ReadClaimPaymentScreen(TempClaimFromPaymentScreen, TempDraftDetails, TempClaim, TempClaimCount, False, False, GetMaxSuffix(TempClaim, TempClaimCount))
            '    For J As Integer = LBound(TempClaimFromPaymentScreen) To UBound(TempClaimFromPaymentScreen)
            '        For K As Integer = LBound(TempClaimFromPaymentScreen(J).POS) To UBound(TempClaimFromPaymentScreen(J).POS)

            '        Next
            '    Next
            'Next
            FlagCARemarkCode = True
            If CADenialApplicabilityAsPerProcCodeAndBillCharge() = False Then
                'WriteLog("Current claim code count mismatch with history claim codes, hence ca not applicable.")
                Return False
            End If
            iSRCCA = 1
            'WriteLog("CA denial Applies, now checking history void condition")
            ' If any ICN in Hx with Paid or Zero Paid suffix is previously voided(entire claim) then CA denial is not applicable, otherwise CA denial is not applicable

            For K As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                'For M As Integer = 0 To GetClaimCountFromStruct(ListOfPossibleICNActual(K).TempClaim) - 1
                '    If ListOfPossibleICNActual(K).TempClaim(M).Voided Then
                '        If ListOfPossibleICNActual(K).TempClaim(M).ProcCode(UBound(ListOfPossibleICNActual(K).TempClaim(M).ProcCode) - 1).Trim <> "RX" AndAlso CDbl(ListOfPossibleICNActual(K).TempClaim(M).PaidAmount) > 0 Then
                '            Return False
                '        End If
                '    End If
                'Next
                If CAApplicabilityForHxClaimByVoidedStatus(ListOfPossibleICNActual(K)) = False Then Return False
            Next
            Return True
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name + " exception: " + ex.Message.ToString
            Return False
        End Try
    End Function
    Public Function GetMaxSuffix(ByRef MyClaim() As ClaimFromPaymentScreen) As Integer
        Dim I, J As Integer
        J = 0
        Try
            For I = LBound(MyClaim) To UBound(MyClaim)
                If MyClaim(I).Suffix > J Then J = MyClaim(I).Suffix
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return J
    End Function
    Public Function CAApplicabilityForHxClaimByVoidedStatus(ByRef MyIcn As PossibleICN) As Boolean
        Dim MaxSuffix As Integer = GetMaxSuffix(MyIcn.TempClaimFromPaymentScreen)
        Dim FlagProcessed As Boolean = False
        Dim LastProcessedDrafts As String = ""

        For I As Integer = LBound(MyIcn.TempClaim) To UBound(MyIcn.TempClaim) - 1
            For J As Integer = LBound(MyIcn.TempClaim(I).ServiceLine) To UBound(MyIcn.TempClaim(I).ServiceLine) - 1
                If (MyIcn.TempClaim(I).RemarkCode(J) = "OV" AndAlso MyIcn.TempClaim(I).OV(J).Trim = "R") OrElse
                    (MyIcn.TempClaim(I).RemarkCode(J) = "77" AndAlso MyIcn.TempClaim(I).OV(J) = "20") Then
                    Return False
                End If
            Next
        Next
        For I As Integer = 1 To MaxSuffix
            For J = LBound(MyIcn.TempDraftDetails) To UBound(MyIcn.TempDraftDetails)
                If CInt(MyIcn.TempDraftDetails(J).Suffix) = I Then
                    LastProcessedDrafts += MyIcn.TempDraftDetails(J).Draft + ","
                    Exit For
                End If
            Next
        Next
        For Each draft As String In LastProcessedDrafts.Split(",")
            If draft.Trim <> "" Then
                For I As Integer = 0 To GetClaimCountFromStruct(MyIcn.TempClaim) - 1
                    If MyIcn.TempClaim(I).Draft = draft AndAlso MyIcn.TempClaim(I).Voided = False Then
                        Dim FlagValidSuffix As Boolean = True
                        For J As Integer = LBound(MyIcn.TempClaim(I).POS) To UBound(MyIcn.TempClaim(I).POS) - 1
                            If InStr("75,77", IIf(MyIcn.TempClaim(I).RemarkCode(0).Trim = "", "**", (MyIcn.TempClaim(I).RemarkCode(0).Trim))) > 0 Then
                                FlagValidSuffix = False
                                FlagProcessed = True
                                Exit For
                            ElseIf InStr("69,70,71", IIf(MyIcn.TempClaim(I).RemarkCode(0).Trim = "", "**", (MyIcn.TempClaim(I).RemarkCode(0).Trim))) > 0 Then
                                FlagValidSuffix = False
                                Exit For
                            ElseIf InStr("AC", IIf(MyIcn.TempClaim(I).RemarkCode(0).Trim = "", "**", (MyIcn.TempClaim(I).RemarkCode(0).Trim))) > 0 AndAlso MyIcn.TempClaim(I).OV(J).Trim = "R" Then
                                FlagValidSuffix = False
                                Exit For
                            End If
                        Next
                        If FlagValidSuffix Then
                            For J As Integer = LBound(MyIcn.TempClaim(I).POS) To UBound(MyIcn.TempClaim(I).POS) - 1
                                If MyIcn.TempClaim(I).OV(J).Trim = "C" OrElse CDbl(MyIcn.TempClaim(I).Charge(J)) <> CDbl(MyIcn.TempClaim(I).NotCovered(J)) Then
                                    FlagProcessed = True
                                    Exit For
                                End If
                            Next
                        End If
                        If FlagProcessed Then Exit For
                    End If
                Next
            End If
            If FlagProcessed Then Exit For
        Next
        If FlagProcessed Then
            For Each draft As String In LastProcessedDrafts.Split(",")
                If draft.Trim <> "" Then
                    For I As Integer = 0 To GetClaimCountFromStruct(MyIcn.TempClaim) - 1
                        If MyIcn.TempClaim(I).Draft = draft AndAlso MyIcn.TempClaim(I).Voided = True Then
                            Return False
                        End If
                    Next
                End If
            Next
        End If
        Return True


    End Function
    Function FindIndexInMapByClaimIndexAndLineIndex(ByRef Map() As MapCurrentProcDetailsWithHistory, ByVal Cindex As Integer, ByVal Lindex As Integer, Optional ByVal IsCurrent As Boolean = False) As Integer
        Try
            For I As Integer = LBound(Map) To UBound(Map)
                If IsCurrent Then
                    If Map(I).CurrentClaimSuffix = Cindex And Map(I).CurrentClaimLineNumber = Lindex Then Return I
                Else
                    If Map(I).HistoryClaimCIndex = Cindex And Map(I).HistoryClaimLIndex = Lindex Then Return I
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return -1
    End Function
    Public Function CADenialApplicabilityAsPerProcCodeAndBillCharge() As Boolean
        'Check for Current claim code not present in history claim
        Dim Suffix As Integer = 0
        Dim Line As Integer = 0
        Dim index As Integer
        Dim FlagFound As Boolean
        For I As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
            For J As Integer = LBound(ListOfPossibleICNActual(I).Map) To UBound(ListOfPossibleICNActual(I).Map)
                Suffix = ListOfPossibleICNActual(I).Map(J).CurrentClaimSuffix
                Line = ListOfPossibleICNActual(I).Map(J).CurrentClaimLineNumber
                If ListOfPossibleICNActual(I).Map(J).CurrentClaimProcCode <> "" AndAlso ListOfPossibleICNActual(I).Map(J).CurrentClaimProcCode <> "NCOPAY" Then
                    If (ListOfPossibleICNActual(I).Map(J).CurrentClaimProcCode <> ListOfPossibleICNActual(I).Map(J).HistoryClaimProcCode AndAlso ListOfPossibleICNActual(I).Map(J).IsCodeChange = False) OrElse
                        (CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) <> CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) AndAlso CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) + CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) > 0.01) Then
                        FlagFound = False
                        For K As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                            If K <> I Then
                                index = GetIndexInMapByCurrentSuffixAndLineNumber(ListOfPossibleICNActual(K).Map, Suffix, Line)
                                If index >= 0 Then
                                    If (ListOfPossibleICNActual(I).Map(J).CurrentClaimProcCode = ListOfPossibleICNActual(I).Map(J).HistoryClaimProcCode OrElse ListOfPossibleICNActual(I).Map(J).IsCodeChange) AndAlso
                                        (CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) = CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) OrElse CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) + CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) <= 0.01) Then
                                        FlagFound = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        If Not FlagFound Then Return False
                    End If
                End If
            Next
        Next
        For I As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
            For J As Integer = LBound(ListOfPossibleICNActual(I).Map) To UBound(ListOfPossibleICNActual(I).Map)
                If ListOfPossibleICNActual(I).Map(J).HistoryClaimProcCode <> "" AndAlso ListOfPossibleICNActual(I).Map(J).HistoryClaimProcCode <> "NCOPAY" Then
                    If (ListOfPossibleICNActual(I).Map(J).CurrentClaimProcCode <> ListOfPossibleICNActual(I).Map(J).HistoryClaimProcCode AndAlso ListOfPossibleICNActual(I).Map(J).IsCodeChange = False) OrElse
                        (CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) <> CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) AndAlso CDbl(ListOfPossibleICNActual(I).Map(J).CurrentClaimCharge) + CDbl(ListOfPossibleICNActual(I).Map(J).HistoryClaimCharge) > 0.01) Then
                        Return False
                    End If
                End If
            Next
        Next
        Return True
    End Function
    Public Function GetIndexInMapByCurrentSuffixAndLineNumber(ByRef MyMap() As MapCurrentProcDetailsWithHistory, ByVal Suffix As Integer, ByVal LineNumber As Integer) As Integer
        Dim indx As Integer = -1
        Try
            For I As Integer = LBound(MyMap) To UBound(MyMap)

                If MyMap(I).CurrentClaimSuffix = Suffix AndAlso MyMap(I).CurrentClaimLineNumber = LineNumber Then Return I
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return indx
    End Function

    Public Function GetSixPointFiveScreen() As String
        Dim Text As String = String.Empty
        Text = GetText(4, 2, 79)


        Return Text
    End Function
    Public Function GetTimeZone() As String
        Dim TempTime As String
        Dim Temp() As String
        Dim now As DateTime = DateTime.Now
        Dim ISTNow = System.TimeZoneInfo.ConvertTime(now, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")) '"India Standard Time")) https://stackoverflow.com/questions/7908343/list-of-timezone-ids-for-use-with-findtimezonebyid-in-c
        TempTime = ISTNow.ToLongDateString()
        Temp = Split(TempTime, ",")
        TempTime = Temp(1) & ", " & Temp(2) & " " & ISTNow.ToLongTimeString
        'TempTime = Replace(TempTime, ",", vbTab)
        GetTimeZone = Replace(TempTime, "  ", " ") & vbTab & Temp(1) & ", " & Temp(2)
    End Function
    Dim CurNTID As String = ""
    Public Function GetNTID() As String
        'If Not boolOffshore Then
        '    If CurNTID = "" Then
        '        Dim num = CInt(Math.Ceiling(Rnd() * 3)) + 1
        '        Select Case num
        '            Case 1
        '                CurNTID = "YDOMINIC"
        '            Case 2
        '                CurNTID = "CPAVON1"
        '            Case 3
        '                CurNTID = "MIBARRA7"
        '            Case Else
        '                CurNTID = "MIBARRA7"
        '        End Select
        '    End If
        '    Return CurNTID
        'Else
        'End If
        Dim wshShell As Object
        wshShell = CreateObject("WScript.Shell")
        Return wshShell.ExpandEnvironmentStrings("%USERNAME%").ToString.ToUpper '.Replace("_", "")

    End Function
    Public Function GetTimeZone(ByVal Timenow As DateTime) As String
        Dim TempTime As String
        Dim Temp() As String
        'Dim now As DateTime = DateTime.Now
        Dim ISTNow = System.TimeZoneInfo.ConvertTime(Timenow, TimeZoneInfo.FindSystemTimeZoneById("India Standard Time"))
        TempTime = ISTNow.ToLongDateString()
        TempTime = ISTNow.ToLongTimeString()
        Return Replace(TempTime, ":", "-")
        Temp = Split(TempTime, " ")
        Temp = Split(TempTime, ",")
        TempTime = Temp(1) & ", " & Temp(2) & " " & ISTNow.ToLongTimeString
        'TempTime = Replace(TempTime, ",", vbTab)
        Return Replace(TempTime, "  ", " ") & vbTab & Temp(1) & ", " & Temp(2)
    End Function
    Public Sub DumpLog(r As StreamReader)
        Dim line As String
        line = r.ReadLine()
        While Not (line Is Nothing)
            Console.WriteLine(line)
            line = r.ReadLine()
        End While
    End Sub

    Public Function CountDigits(txt As String) As Integer
        Dim num As Integer = 0
        Dim count As Integer = 0
        Try
            'num = Integer.Parse(txt)

            'While (num > 0)
            '    count = count + 1
            '    num = num / 10
            'End While
            count = txt.Length()
        Catch e As System.FormatException
            strResearchStatus = "Failed to convert to integer"
            strResearchComments = "Cannot extract Type from EDIT"
            WriteLog(GetCurrentMethod.Name + " exception")
            'Exit Function
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return count
    End Function

    Public Sub PasteHistoryICN(ByVal ICN As String, Optional ByRef FlagError As Boolean = False, Optional ByRef strResearchStatus As String = "", Optional ByRef strResearchComments As String = "")
        Dim fln As String
        Dim dcc As String
        Try
            If GetText(2, 2, 2) <> "MP" Then
                SendClear()
                SendClear()
                pullEDSScreen(ICN, 2)
                SendKeyesFor01("A", 1, 76)
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
            End If
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
                apiChk()
                If (GetText(4, 15, 1) = "-" AndAlso GetText(4, 32, 1) = "-") Or (GetText(4, 15, 10) = "9999999999" AndAlso GetText(4, 32, 10) = "9999999999") Or (GetText(4, 15, 10) = "AA12345678" AndAlso GetText(4, 32, 10) = "0012345678") Then
                    pullEDSScreen(ICN, 9)
                    fln = GetText(2, 23, 10)
                    dcc = GetText(2, 37, 3)
                    pullEDSScreen(ICN, 8)
                    Dim edsArr(1000) As String
                    'Dim lineTxtArr(100) As String
                    Dim origAdjICN As String
                    Dim draftNum As String = ""
                    If GetText(5, 56, 2) = "UB" Then
                        GetEdssText(edsArr, fln, dcc)
                        If edsArr.Length > 53 Then
                            If edsArr(53) IsNot Nothing Then
                                Dim linetext As String
                                linetext = edsArr(53).Trim()
                                Dim arr As String()
                                arr = Regex.Split(linetext.Trim(), "ORIG CLM REF NUMBER:")
                                If arr.Count > 1 Then
                                    ' arr = arr.Where(X >= !String.IsNullOrEmpty(X)).ToArray()
                                    Dim spText = arr(1).Replace("|", "").Trim
                                    If spText <> Nothing Then
                                        If arr(1).Trim().Length >= 10 Then
                                            arr(1) = arr(1).Replace(" ", "")
                                            origAdjICN = arr(1).Trim().Substring(0, 10)
                                            If origAdjICN <> "" Then
                                                historyICN = origAdjICN
                                                'lineTxtArr = edsArr(53).Replace("|", " ").Trim.Split(" ")
                                                'origAdjICN = lineTxtArr(4).Trim
                                                draftNum = GetDraftFromICNForEDS5(origAdjICN)
                                                If draftNum <> "" Then
                                                    histDraft = draftNum
                                                    SendClear()
                                                    SendClear()
                                                    pullEDSScreen(ICN, 2)
                                                    SendKeyes("A", 1, 76)
                                                    SendEnter()
                                                    'apiChk()
                                                    'apiChk()
                                                    If GetText(2, 2, 2) = "MP" Then
                                                        'RemovePZ1()
                                                        IncreementMPCCount()
                                                    End If
                                                    SendKeyes("EDS", 2, 2)
                                                    SendEnter()
                                                    apiChk()
                                                    SendKeyes(origAdjICN, 4, 15)
                                                    'SendEnter()
                                                    SendKeyes(draftNum, 4, 32)
                                                    SendEnter()
                                                    apiChk()
                                                    If GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") Then
                                                        autECLPSObj.SendKeys("N", 4, 79)
                                                        apiChk()
                                                        autECLPSObj.SendKeys("[ENTER]")
                                                        apiChk()
                                                    End If
                                                    SendKeyes("A", 1, 75)
                                                    SendEnter()
                                                    'apiChk()
                                                    'apiChk()
                                                    If GetText(2, 2, 2) = "MP" Then
                                                        'RemovePZ1()
                                                        IncreementMPCCount()
                                                    End If
                                                    If GetText(2, 2, 3) <> "MPC" Then
                                                        FlagError = True
                                                        WriteLog(GetTextRect())
                                                        strResearchStatus = "SKIPPED"
                                                        strResearchComments = "unable to pull payment screen"
                                                        Exit Sub
                                                    End If
                                                    'CreateSecondaryLog({"CopyHistoryICN"}, {"YES"}, CurrICN)
                                                Else
                                                    FindHxICNandDraftWhenHardCopyICNisInvalid(ICN)
                                                    FlagError = True
                                                    WriteLog(GetTextRect())
                                                    Exit Sub
                                                End If
                                            Else
                                                FindHxICNandDraftWhenHardCopyICNisInvalid(ICN)
                                                WriteLog(GetTextRect())
                                            End If
                                        End If
                                    Else
                                        FindHxICNandDraftWhenHardCopyICNisInvalid(ICN)
                                        WriteLog(GetTextRect())

                                    End If

                                End If
                            End If
                        End If
                        'SendKeyes("A", 1, 76)
                        'SendEnter()
                        'If GetText(2, 2, 3) <> "MPC" Then
                        '    FlagError = True
                        '    WriteLog(GetTextRect())
                        '    strResearchStatus = "SKIPPED"
                        '    strResearchComments = "unable to pull payment screen"
                        '    Exit Sub
                        'End If

                    End If
                Else
                    WriteLog("INSIDE " + GetCurrentMethod.Name)
                    WriteLog(GetTextRect)
                    autECLPSObj.SendKeys("N", 4, 79)
                    apiChk()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    'apiChk()
                    'apiChk()
                    'apiChk()
                    'apiChk()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
            FindHxICNandDraftWhenHardCopyICNisInvalid(ICN)
            WriteLog(GetTextRect())
        End Try
        WriteLog("EXIT " + GetCurrentMethod.Name)

    End Sub

    Public Function GetDraftFromICNForEDS5(ByVal ICN As String) As String
        Dim isLastPage As Boolean = False
        Dim draftNum As String = ""
        Try
            PullMHI(ICN)
            If InStr(strResearchComments, "NOT AUTHORIZED") > 0 Then
                Return draftNum
                Exit Function
            ElseIf InStr(strResearchComments, "Unable to pull MHI") > 0 Then
                strResearchComments = GetCurrentMethod.Name + " :Unable to pull MHI"
                Return draftNum
                Exit Function
            ElseIf InStr(strResearchComments, "W703DOC PROCESSED") > 0 Then
                Return draftNum
                Exit Function
            End If
            If GetText(2, 2, 2) = "MH" Then 'Valid ICN

                For row As Integer = 12 To 20 Step 4
                    If GetText(row, 1, 3) = "ICN" Then
                        draftNum = GetText(row - 2, 28, 10)

                    End If

                    If row = 20 Then
                        'Dim MyscreenMHI As String = GetTextRect(10, 1, 20, 80)
                        SendF8()
                        If isLastPage Then
                            Exit For
                        End If
                        If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then
                            isLastPage = True
                        End If
                        row = 8
                    End If
                Next

            Else 'Invalid ICN
                Return ""
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try

        Return draftNum
    End Function

    Public Function addZeroesToBillCharge(txt As String) As String
        Dim dec As String = Math.Round(Convert.ToDecimal(txt), 2).ToString.Split(".")(1)
        'Dim dec As String = Right(txt, 2)
        Dim wholeNum As String = Math.Truncate(Convert.ToDecimal(txt))

        Dim wholeNumDigitCount As Integer = CountDigits(wholeNum)
        If wholeNumDigitCount = 1 Then
            wholeNum = "00000" + wholeNum + "." + dec
        ElseIf wholeNumDigitCount = 2 Then
            wholeNum = "0000" + wholeNum + "." + dec
        ElseIf wholeNumDigitCount = 3 Then
            wholeNum = "000" + wholeNum + "." + dec
        ElseIf wholeNumDigitCount = 4 Then
            wholeNum = "00" + wholeNum + "." + dec
        ElseIf wholeNumDigitCount = 5 Then
            wholeNum = "0" + wholeNum + "." + dec
        ElseIf wholeNumDigitCount = 0 And wholeNum = "0" Then
            wholeNum = "000000" + "." + dec
        Else
            wholeNum = wholeNum + "." + dec
        End If
        Return wholeNum
    End Function

    Public Function isROVand75or77Present(icnList As String) As Boolean
        Dim flag As Boolean = False
        Dim totalAmt As Decimal = 0
        Dim newdata As String = ""
        Dim olddata As String = ""
        Dim SafetyPageCount As Integer = 0
        Try
            WriteLog("================ Inside " + GetCurrentMethod.Name + " ================")
            PullAHI()
            WriteLog(GetTextRect())
            For dd As Integer = 4 To 20 Step 2
                If Trim(GetText(dd, 4, 6)) = "------" Then Exit For

                If InStr(Trim(GetText(dd, 64, 5).Replace("-", "")), "OV") > 0 And InStr(icnList, Trim(GetText(dd + 1, 54, 10))) > 0 Then
                    SendKeyes("S", dd, 2)
                    SendEnter()
                    apiChk()
                    WriteLog(GetTextRect())
                    Dim isUB As Boolean = False
                    For cc As Integer = 10 To 18 Step 4         'This loop is to check whether the ICN is UB or not  --on 18/11/24
                        If Trim(GetText(cc, 1, 2)) = "" Then Exit For

                        If (InStr(Trim(GetText(cc, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(cc, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(cc, 4, 4)), "MISC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OPS") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OA") > 0) = True Then
                            isUB = True
                            Exit For
                        End If
                        If cc = 18 And isUB = False Then
                            SendKeyes("A", 3, 15)
                            apiChk()
                            SendEnter()
                            apiChk()
                            Exit For
                        End If
                    Next
                    If isUB = True Then
                        For cc As Integer = 10 To 18 Step 4
                            If Trim(GetText(cc + 1, 1, 80)) = "" And InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then
                                SendKeyes("A", 3, 15)
                                apiChk()
                                SendEnter()
                                apiChk()
                                Exit For
                            End If
                            If Trim(GetText(cc + 2, 1, 3)) = "ICN" And cc = 18 Then
                                If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then
                                    If flag = False Then
                                        strResearchComments = "ROV found in HxICN, please ignore ROV"
                                        Return flag
                                    Else
                                        SendKeyes("A", 3, 15)
                                        apiChk()
                                        SendEnter()
                                        apiChk()
                                        Exit For
                                    End If
                                End If
                                SendF8()
                                cc = 6
                                Continue For
                            ElseIf Trim(GetText(cc + 2, 1, 3)) = "ICN" Then
                                Continue For
                            End If
                            'If Trim(GetText(cc, 36, 2)) = "77" Or Trim(GetText(cc, 36, 2)) = "75" Then              THIS WAS EXPLAINED BY RAJIB BEFORE ERROR CAME       
                            '    totalAmt = totalAmt + Convert.ToDecimal(Trim(GetText(cc, 41, 8)))
                            '    flag = True
                            'End If
                            If Trim(GetText(cc, 29, 1)) = "R" Or Trim(GetText(cc, 36, 2)) = "OV" Then
                                totalAmt = totalAmt + Convert.ToDecimal(Trim(GetText(cc, 41, 8)))
                                flag = True
                                WriteLog($"ROV in History flag= {flag}")
                                WriteLog(GetTextRect())
                            End If
                            If cc = 18 Then
                                If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then
                                    If flag = False Then
                                        strResearchComments = "ROV found in HxICN, please ignore ROV"
                                        Return flag
                                    Else
                                        SendKeyes("A", 3, 15)
                                        apiChk()
                                        SendEnter()
                                        apiChk()
                                        Exit For
                                    End If
                                End If
                                olddata = GetTextRect(10, 1, 20, 80)
                                SendF8()
                                newdata = GetTextRect(10, 1, 20, 80)
                                SafetyPageCount = SafetyPageCount + 1

                                ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                If SafetyPageCount > 50 Then
                                    WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                    Exit For
                                End If
                                Dim loopCount As Integer = 0
                                While olddata = newdata
                                    SendF8()
                                    loopCount += 1
                                    If loopCount > 5 Then Exit While
                                    newdata = GetTextRect(10, 1, 20, 80)
                                    If newdata <> olddata Then Exit While
                                End While
                                cc = 6
                            End If
                        Next
                    End If


                    'For cc As Integer = 10 To 18 Step 4
                    '    If Trim(GetText(cc, 29, 1)) = "R" And Trim(GetText(cc, 36, 2)) = "OV" Then
                    '        ROVAmount = Trim(GetText(cc, 41, 8))
                    '        flag = True
                    '    End If
                    'Next
                End If
                If dd = 20 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    SendF8()
                    dd = 2
                End If
            Next
            If flag Then
                ROVAmount = totalAmt
            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try

        Return flag
    End Function
    Public Sub FindHxICNandDraftWhenHardCopyICNisInvalid(ByVal ICN As String)
        Try
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            Else
                SendClear()
                SendClear()
                pullEDSScreen(ICN, 2)
                SendKeyes("A", 1, 76)
                SendEnter()
                'apiChk()
                'apiChk()
                'apiChk()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
            End If

            For J As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                If ListOfPossibleICNActual(J).isPaid AndAlso Mid(ListOfPossibleICNActual(J).TempClaim(0).Tin, 3, 9) = Mid(CurrentClaim(0).Tin, 3, 9) Then
                    For K As Integer = LBound(ListOfPossibleICNActual(J).TempDraftDetails) To UBound(ListOfPossibleICNActual(J).TempDraftDetails)
                        apiChk()
                        If CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft) > 0 And ListOfPossibleICNActual(J).TempDraftDetails(K).isClosed = False And InStr("R-71,69-70", ListOfPossibleICNActual(J).TempDraftDetails(K).ReversalType) > 0 And CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).TotalAllowable) <> CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).TotalNotCovered) Then
                            SendKeyes(ListOfPossibleICNActual(J).TempClaim(0).ICN, 4, 15)
                            SendKeyes(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft, 4, 32)
                            GoTo OutOfEdsHalfScreen
                        End If
                    Next
                End If
            Next

            For J As Integer = LBound(ListOfPossibleICNActual) To UBound(ListOfPossibleICNActual)
                For K As Integer = LBound(ListOfPossibleICNActual(J).TempDraftDetails) To UBound(ListOfPossibleICNActual(J).TempDraftDetails)
                    apiChk()
                    If CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft) > 0 AndAlso ListOfPossibleICNActual(J).TempDraftDetails(K).isClosed = False AndAlso InStr("R-71,69-70", ListOfPossibleICNActual(J).TempDraftDetails(K).ReversalType) > 0 AndAlso Mid(ListOfPossibleICNActual(J).TempClaim(0).Tin, 3, 9) = CurrentClaimDetails.Tin Then
                        SendKeyes(ListOfPossibleICNActual(J).TempClaim(0).ICN, 4, 15)
                        SendKeyes(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft, 4, 32)
                        GoTo OutOfEdsHalfScreen
                    ElseIf CDbl(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft) > 0 AndAlso ListOfPossibleICNActual(J).TempDraftDetails(K).isClosed = False AndAlso InStr("R-71,69-70", ListOfPossibleICNActual(J).TempDraftDetails(K).ReversalType) > 0 Then
                        SendKeyes(ListOfPossibleICNActual(J).TempClaim(0).ICN, 4, 15)
                        SendKeyes(ListOfPossibleICNActual(J).TempDraftDetails(K).Draft, 4, 32)
                        GoTo OutOfEdsHalfScreen
                    End If
                Next
            Next
            SendKeyes("----------", 4, 15)
            SendKeyes("----------", 4, 32)


            'Else
            '    For I As Integer = 0 To HistoryClaimCount
            '        If CDbl(HistoryClaim(I).Draft) <> 0 AndAlso HistoryClaim(I).Voided = False Then
            '            SendKeyes(HistoryClaim(0).ICN, 4, 15)
            '            SendKeyes(HistoryClaim(I).Draft, 4, 32)
            '            Exit For
            '        End If
            '    Next
            'End If
OutOfEdsHalfScreen:
            SendEnter()
            Dim text As String = String.Empty
            text = GetSixPointFiveScreen()
            WriteLog(text)
            SendKeyes("N", 4, 79)
            SendEnter()
            SendKeyes("A", 1, 75)
            SendEnter()
            'apiChk()
            'apiChk()
            If GetText(2, 2, 2) = "MP" Then
                'RemovePZ1()
                IncreementMPCCount()
            End If
            If GetText(2, 2, 2, True) = "AR" Then
                SendKeyes("S", 3, 5)
                SendEnter()
            End If
            If GetText(2, 2, 2) <> "MP" Then
                strResearchComments = "UNABLE TO A OVER FROM EDS 6 SCREEN"
                ExitAllAndReturnWithException()
            End If
            PutCopayNBR()
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
    End Sub

    Public Function AlternatePercentageMatching(tempList As String) As String   'Added 21/01/2022 to perform %age matching on the eliminated HxICN based on only ProcCode and BillCharge
        Dim FinalList As String = String.Empty
        Dim tempListArr() As String = tempList.Split(",")
        Dim FindIcn As String
        Dim PICN As String = String.Empty

        Try
            For ff As Integer = LBound(tempListArr) To UBound(tempListArr) - 1
                Dim EDS1HCPC As String = String.Empty
                Dim EDS1Charge As String = String.Empty
                Dim HEDS1HCPC As String = String.Empty
                Dim HEDS1Charge As String = String.Empty
                FindIcn = String.Empty
                PICN = tempListArr(ff)
                SendClear()
                SendClear()
                SendKeyes("mei,i" + PICN + ",1", 2, 2)
                SendEnter()
                apiChk()
                If InStr(Trim(GetText(24, 2, 79)), "SELECTED CLAIM NOT FOUND") > 0 Then
                    Dim DraftComments As String = ""
                    GetHRHI(PICN, HistoryRHI, DraftComments)
                    If DraftComments = "Draft list is empty" Then
                        Continue For
                    End If
                    If InStr(strResearchComments, "W062NO CLAIMS ON FILE") > 0 Then
                        Continue For
                    End If
                    Dim kk As Integer = 0
                    Dim ll As Integer = 0
                    Dim count1 As Integer = 0
                    Dim TotalCount1 As Integer = 0
                    Dim Divresult1 As Decimal = 0
                    Dim PercentageResult1 As Integer = 0
                    For rr As Integer = LBound(HistoryRHI) To UBound(HistoryRHI)
                        For uu As Integer = LBound(HistoryRHI(rr).RHIServiceLineCount.RevCode) To UBound(HistoryRHI(rr).RHIServiceLineCount.RevCode)
                            TotalCount1 += 1
                            HEDS1Charge = HistoryRHI(rr).RHIServiceLineCount.Charge(uu)
                            HEDS1HCPC = HistoryRHI(rr).RHIServiceLineCount.ProcCode(uu)
                            For kk = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                For ll = LBound(CurrentICNEDS1(kk).ServiceLine.RevCode) To UBound(CurrentICNEDS1(kk).ServiceLine.RevCode)
                                    EDS1Charge = CurrentICNEDS1(kk).ServiceLine.BillCharge(ll)
                                    EDS1HCPC = CurrentICNEDS1(kk).ServiceLine.CPTCode(ll)
                                    If InStr(EDS1HCPC, "-") > 0 Then
                                        EDS1HCPC = Trim(EDS1HCPC.Replace("-", ""))
                                    End If
                                    If HEDS1HCPC = EDS1HCPC Then
                                        count1 += 1
                                    End If
                                Next
                            Next
                        Next
                        kk = 0
                        ll = 0
                    Next
                    Divresult1 = count1 / TotalCount1
                    PercentageResult1 = Divresult1 * 100
                    If PercentageResult1 >= 40 Then
                        If InStr(FinalList, PICN) = 0 Then
                            FinalList = PICN + "," + FinalList
                        End If
                    End If
                Else
                    If Trim(GetText(6, 2, 5)) = "EDS 1" Then
                        GetEDS1(HistoryICNEDS1, PICN)
                        Dim hh As Integer = 0
                        Dim nn As Integer = 0
                        Dim count As Integer = 0
                        Dim TotalCount As Integer = 0
                        Dim PercentageResult As Integer = 0
                        Dim DivResult As Decimal = 0
                        For bb As Integer = LBound(HistoryICNEDS1) To UBound(HistoryICNEDS1)
                            For ll As Integer = LBound(HistoryICNEDS1(bb).ServiceLine.RevCode) To UBound(HistoryICNEDS1(bb).ServiceLine.RevCode)
                                TotalCount += 1
                                HEDS1Charge = HistoryICNEDS1(bb).ServiceLine.BillCharge(ll)
                                HEDS1HCPC = HistoryICNEDS1(bb).ServiceLine.CPTCode(ll)
                                For hh = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                                    For nn = LBound(CurrentICNEDS1(hh).ServiceLine.RevCode) To UBound(CurrentICNEDS1(hh).ServiceLine.RevCode)
                                        EDS1Charge = CurrentICNEDS1(hh).ServiceLine.BillCharge(nn)
                                        EDS1HCPC = CurrentICNEDS1(hh).ServiceLine.CPTCode(nn)
                                        If HEDS1HCPC = EDS1HCPC Then
                                            count += 1
                                        End If
                                    Next
                                Next
                            Next
                            hh = 0
                            nn = 0
                        Next
                        DivResult = count / TotalCount
                        PercentageResult = DivResult * 100

                        If PercentageResult >= 40 Then
                            FindIcn = PICN
                            If InStr(FinalList, PICN) = 0 Then
                                FinalList = PICN + "," + FinalList
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name() + " exception: " + ex.Message.ToString)
            strResearchComments = GetCurrentMethod.Name() + " exception: " + ex.Message.ToString
        End Try

        Return FinalList
    End Function

    Public Function CheckIfValidICN(hardcopyICN As String) As Boolean    'Added on 25/01/2022 to check if hardCopyICN is valid by opening MHI of that ICN
        Dim isValid As Boolean = False
        If hardcopyICN <> "" Then
            SendClear()
            SendClear()
            SendKeyes("ret,i" + hardcopyICN + ",m", 2, 2)
            SendEnter()
            apiChk()
            apiChk()
            If InStr(Trim(GetText(24, 2, 80)), "INVALID ICN") > 0 Or hardcopyICN = CurrentClaimDetails.ICN Then
                isValid = False
            ElseIf GetText(2, 2, 2) = "MH" Then
                isValid = True
            End If
        End If

        Return isValid
    End Function

    Public Function CheckCurrentClaimforOHandOPS() 'Added 18/02/2022
        Dim flag As Boolean = False
        SendClear()
        SendClear()
        SendKeyes("mei,i" + CurrentClaimDetails.ICN + ",1", 2, 2)
        SendEnter()
        apiChk()
        apiChk()
        If Trim(GetText(6, 2, 5)) = "EDS 1" Then
            For Y = 8 To 21 Step 1
                If Trim(GetText(Y, 2, 4)) <> "" Then
                    If Trim(GetText(Y, 63, 2)) = "OH" And Trim(GetText(Y, 66, 3)) = "OPS" Then
                        flag = True
                        Exit For
                    End If
                Else
                    Exit For
                End If
                If Y = 21 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO ADDITIONAL LINES TO DISPLAY") > 0 Then Exit For

                    If InStr(autECLPSObj.GetText(24, 1, 80), "PF8 FOR ADDITIONAL LINES") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "I049 - MORE DATA ON EDSS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "PF4 TO REVIEW EDC/ICES DECISION") > 0 Then
                        SendF8()
                        'WriteLog(vbCrLf)
                        'WriteLog(GetTextRect())
                    Else
                        Exit For
                    End If
                    Y = 7
                End If
            Next
        End If
        Return flag
    End Function
    Public Function OHandOPSinHxICN(tempList As String) 'Added 18/02/2022
        Dim FinalList As String = String.Empty
        Dim HICN As String = String.Empty
        Dim tempListArr() As String = tempList.Split(",")
        Try
            For ff As Integer = LBound(tempListArr) To UBound(tempListArr) - 1
                HICN = tempListArr(ff)
                PullMHI(HICN)
                For cc As Integer = 10 To 18 Step 4
                    If Trim(GetText(cc + 2, 1, 3)) = "ICN" Then Continue For

                    If Trim(GetText(cc, 1, 2)) = "OH" And Trim(GetText(cc, 4, 3)) = "OPS" Then
                        If InStr(FinalList, HICN) = 0 Then
                            FinalList = HICN + "," + FinalList
                            Exit For
                        End If
                    End If
                    If cc = 18 Then
                        If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        SendF8()
                        cc = 6
                    End If
                Next
            Next

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + ": " + ex.Message.ToString)
        End Try
        Return FinalList
    End Function
    Public Function pullHxRHI(ByVal icn As String) As Boolean
        Dim isHxUB As Boolean = False
        Dim DraftNoList As String = String.Empty
        Dim MHILine As String = String.Empty
        DraftNoList = HistoryDraft(icn)
        If String.IsNullOrEmpty(DraftNoList) Or InStr(DraftNoList, "E028NOT AUTHORIZED") > 0 Then
            'DraftListComments = "Draft list is empty"
            Return isHxUB
        End If
        Dim DraftList() As String = DraftNoList.Split(",")
        MHILine = GetMHIControlLine()

        For dd As Integer = LBound(DraftList) To UBound(DraftList) - 1

            SendClear()
            SendClear()
            SendKeyes(MHILine + DraftList(dd) + ",I" + icn, 2, 2)
            SendKeyes("R", 2, 2)
            SendEnter()
            apiChk()
            apiChk()
            If Trim(GetText(8, 3, 3)) = "REV" Then
                isHxUB = True
                Return isHxUB
            Else
                isHxUB = False
                Return isHxUB
            End If
        Next
        Return isHxUB
    End Function

    Public Function REMOVE01IFSKIPPED(ByVal ICN As String, SuffixCount As Integer) As String                   'Added June,2022; Modified March, 2023
        WriteLog("===IN " + GetCurrentMethod.Name + "===========")
        Dim flag01Found As Boolean = False
        For sfx As Integer = 1 To SuffixCount

            SendClear()
            SendClear()
            pullEDSScreen(ICN, 2)
            SendKeyesFor01("A" & sfx, 1, 76)
            autECLPSObj.SendKeys("[ENTER]")
            apiChk()

            If InStr(GetText(24, 1, 80), "NOT AUTHORIZED") > 0 Then
                Return "E028NOT AUTHORIZED: REMOVE 01 OV"
            End If
            If GetText(2, 2, 2) = "AR" Then
                Return "Manual intervention: REMOVE 01 OV"
            End If
            If GetText(2, 2, 2) <> "MP" Then
                Return "Manual intervention: REMOVE 01 OV"
            End If
            'If InStr("NOT AUTHORIZED", GetText(24, 1, 80)) > 0 Then
            '    Return "Not Authorized: MPC"
            'End If
            If GetText(2, 2, 2) = "MP" Then
                WriteLog("=============before 01 OV is checked=========")
                WriteLog(GetTextRect)
                For I = 10 To 22 Step 2
                    If GetText(I, 29, 1) <> "P" And GetText(I, 29, 1) <> "R" And GetText(I, 29, 2) <> "11" And GetText(I, 29, 1) <> "C" And
                    GetText(I, 29, 2) <> "--" And GetText(I, 29, 2) <> "23" And GetText(I, 29, 2) <> "24" Then
                        flag01Found = True
                        Exit For
                    End If
                Next
            End If

            If flag01Found = True Then
                SendKeyesFor01("EDS", 2, 2)
                autECLPSObj.SendKeys("[ENTER]")

                apiChk()
                apiChk()
                If GetText(3, 2, 3) <> "EDS" Then Return "Manual intervention: REMOVE 01 OV"

                Dim olddata As String = ""
                Dim newdata As String = ""
Start:
                If autECLPSObj.SearchText("EDS 5", 3, 2, 5) Or autECLPSObj.SearchText("EDS 6", 3, 2, 5) Then
                    Do
                        For m = 6 To 18 Step 3
                            'If InStr("ACPD,TDPD,EMER,DISCO,OI,OIM,OIMEDI,NYS,INT", Trim(GetText(m, 11, 5))) > 1 Then
                            '    SendKeyesFor01("BLANK", m, 11)
                            '    SendKeyesFor01("--", m, 38)
                            'ElseIf InStr("ACPD,TDPD,EMER,DISCO,OI,OIM,OIMEDI,NYS,INT", Trim(GetText(m, 11, 5))) = 0 Then
                            '    If GetText(m, 2, 3) = "---" Then
                            '        If isProtected(m + 1, 6) = False Then
                            '            SendKeyesFor01("--", m, 38)
                            '        End If
                            '    End If
                            'End If
                            If Trim(GetText(m, 1, 80)) = String.Empty Then Exit For
                            If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                RemoveNotCoveredFromEDS6Point5()
                                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                    For cc As Integer = 7 To 19 Step 3
                                        If GetText(cc - 1, 8, 9) <> "         " Then
                                            autECLPSObj.SendKeys("--------", cc, 6)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 20)
                                            apiChk()
                                            autECLPSObj.SendKeys("--------", cc, 28)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 42)
                                            apiChk()
                                            autECLPSObj.SendKeys("------", cc, 50)
                                            apiChk()
                                        End If
                                    Next
                                    SendEnter()
                                    apiChk()
                                End If
                                If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                                    SendKeyesFor01(OriginalDraft, 4, 32)
                                    SendKeyesFor01(OriginalICN, 4, 15)
                                    SendEnter()
                                    apiChk()
                                    SendKeyes("A", 1, 75)
                                    SendEnter()
                                    apiChk()
                                    If GetText(2, 2, 2) = "MP" Then
                                        autECLPSObj.SendKeys("EDS", 2, 2)
                                        SendEnter()
                                        apiChk()
                                    End If
                                End If
                            End If

                            If InStr("ACPD,TDPD,EMER,DISCO,OI,OIM,OIMEDI,NYS,INT", Trim(GetText(m, 11, 5))) > 1 Then
                                'SendKeyesFor01("BLANK", m, 11)
                                SendKeyesFor01("--", m, 38)
                            ElseIf InStr("ACPD,TDPD,EMER,DISCO,OI,OIM,OIMEDI,NYS,INT", Trim(GetText(m, 11, 5))) = 0 Then
                                If GetText(m, 2, 3) = "---" Then
                                    If isProtected(m + 1, 6) = False Then
                                        SendKeyesFor01("--", m, 38)
                                    Else
                                        SendKeyesFor01("--", m, 38)
                                    End If
                                End If
                            End If
                        Next m
                        autECLPSObj.SendKeys("[ENTER]")

                        olddata = GetTextRect(1, 2, 23, 80)
                        SendF8()
                        newdata = GetTextRect(1, 2, 23, 80)
                    Loop Until StrComp(olddata, newdata) = 0
                End If

                SendKeyesFor01("N", 4, 79)
                autECLPSObj.SendKeys("[ENTER]")

                SendKeyesFor01("A", 1, 75)
                autECLPSObj.SendKeys("[ENTER]")

                apiChk()
                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                    For cc As Integer = 7 To 19 Step 3
                        If GetText(cc - 1, 8, 9) <> "         " Then
                            autECLPSObj.SendKeys("--------", cc, 6)
                            apiChk()
                            autECLPSObj.SendKeys("---", cc, 20)
                            apiChk()
                            autECLPSObj.SendKeys("--------", cc, 28)
                            apiChk()
                            autECLPSObj.SendKeys("---", cc, 42)
                            apiChk()
                            autECLPSObj.SendKeys("------", cc, 50)
                            apiChk()
                        End If
                    Next
                    SendEnter()
                    apiChk()
                    SendKeyesFor01("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")

                    SendKeyesFor01("A", 1, 75)
                    autECLPSObj.SendKeys("[ENTER]")

                    apiChk()
                End If
                If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                    RemoveNotCoveredFromEDS6Point5()
                    SendEnter()
                    apiChk()
                    SendKeyesFor01("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")

                    SendKeyesFor01("A", 1, 75)
                    autECLPSObj.SendKeys("[ENTER]")

                    apiChk()
                End If
                If InStr(Trim(GetText(24, 2, 78)), "NOT COV AMT > CHG AMT") > 0 Then
                    Do
                        For m = 6 To 18 Step 3
                            If InStr("ACPD,TDPD,EMER,DISCO,OI,OIM,OIMEDI,NYS,INT", Trim(GetText(m, 11, 5))) = 0 Then
                                If GetText(m, 2, 3) = "---" Then
                                    SendKeyesFor01("------", m + 1, 6)
                                    SendKeyesFor01("--", m + 1, 13)
                                    SendKeyesFor01("------", m + 1, 28)
                                    SendKeyesFor01("--", m + 1, 35)
                                    SendKeyesFor01("---", m + 1, 20)
                                    SendKeyesFor01("---", m + 1, 42)
                                End If
                            End If
                        Next m
                        autECLPSObj.SendKeys("[ENTER]")

                        olddata = GetTextRect(1, 2, 23, 80)
                        SendF8()
                        newdata = GetTextRect(1, 2, 23, 80)
                    Loop Until StrComp(olddata, newdata) = 0
                    GoTo Start
                End If

                If GetText(2, 2, 2) <> "MP" Then
                    SendClear()
                    If GetText(2, 2, 2) <> "MP" Then
                        For I = 10 To 22 Step 2
                            If GetText(I, 29, 1) <> "P" And GetText(I, 29, 1) <> "R" And GetText(I, 29, 2) <> "11" And GetText(I, 29, 1) <> "C" And
                            GetText(I, 29, 2) <> "--" And GetText(I, 29, 2) <> "23" And GetText(I, 29, 2) <> "24" Then 'and GetText(I, 29, 2) = "01" 
                                SendKeyesFor01("--", I, 29)
                            End If
                        Next
                    End If

                    autECLPSObj.SendKeys("[ENTER]")
                    WriteLog(GetTextRect)
                    Return "Manual intervention: REMOVE 01 OV"
                End If


                If GetText(2, 2, 2) = "MP" Then
                    For I = 10 To 22 Step 2
                        If GetText(I, 29, 1) <> "P" And GetText(I, 29, 1) <> "R" And GetText(I, 29, 2) <> "11" And GetText(I, 29, 1) <> "C" And
                        GetText(I, 29, 2) <> "--" And GetText(I, 29, 2) <> "23" And GetText(I, 29, 2) <> "24" Then Return "REMOVE 01 OV " + strResearchComments
                    Next
                End If

                If GetText(2, 2, 2) <> "MP" Then
                    Return "Manual intervention: REMOVE 01 OV"
                End If
            End If
        Next

        WriteLog(GetTextRect)
        WriteLog("============AFTER REMOVING 01 OV=============")

        'PasteHistoryICN(ICN)
        WriteLog("========EXIT " + GetCurrentMethod.Name + "========")
    End Function

    Sub IncreementMPCCount()
        If GetText(2, 2, 2) = "MP" Then
            MPCScreenCount += 1
        End If
    End Sub
    Public Function WorkOnHardCopyClaim(hardCopyArr() As String) As Boolean     'Added on 18th Oct,22 to check if there's paid amount for hardcopy claim, if yes then it gets added to final list
        Dim flag As Boolean = False
        Dim ICN = hardCopyArr(0).Trim
        Dim amount As Decimal = 0.0
        If ICN <> "" Then
            PullMHI(ICN)
            apiChk()
            For dd As Integer = 12 To 20 Step 4
                If GetText(dd, 1, 80).Trim() = "" Then Exit For
                If InStr(Trim(GetText(dd, 1, 3)), "ICN") > 0 Then
                    Dim amtline As String = GetText(dd - 1, 1, 80)
                    Dim word As String() = New String() {"PD "}
                    Dim wordArr As String() = amtline.Split(word, StringSplitOptions.RemoveEmptyEntries)
                    Dim pdAmtArr As String()
                    If wordArr.Length > 1 Then
                        pdAmtArr = wordArr(1).Trim.Split(" ")
                    Else
                        pdAmtArr = wordArr(0).Trim.Split(",")
                    End If
                    amount = pdAmtArr(0).Trim
                End If

                If dd = 20 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                    SendF8()
                    dd = 8

                End If
            Next
            If amount > 0 Then
                flag = True
            Else
                flag = False
            End If
        End If
        Return flag
    End Function

    Public Function RemoveDummyDraftandICN(ByVal ICN As String, ByVal Payloc As String) As String
        WriteLog("==================Entering " + GetCurrentMethod.Name + "===================")
        Dim OrICN As String = ICN
        Dim val As String = ""
        Dim histICN As String = ""
        Dim histDraft As String = ""
        Dim isUB As Boolean = False
        Try

            'If GetText(2, 2, 2) = "MP" Then
            '    autECLPSObj.SendKeys("EDS", 2, 2)
            '    SendEnter()
            '    apiChk()

            'End If
            For SFX As Integer = 1 To SuffixCountOfCurrentICN

                SendClear()
                SendClear()
                pullEDSScreen(ICN, 2)
                SendKeyes("A" & SFX, 1, 76)
                SendEnter()
                apiChk()

                If InStr(Trim(GetText(24, 1, 80)), "E028NOT AUTHORIZED") = 0 Then
                    SendKeyesFor01("EDS", 2, 2)
                    SendEnter()
                End If
                If GetText(3, 2, 5) = "EDS 6" Then
                    If (GetText(4, 15, 10) = "9999999999" Or GetText(4, 15, 10) = "AA12345678") And (GetText(4, 32, 10) = "9999999999" Or GetText(4, 32, 10) = "0012345678") Then

                        GetClaimDetails(CurrentClaimDetails, ICN, Payloc)
                        PullAHI()
                        If GetText(2, 2, 2) <> "AH" Or InStr(Trim(GetText(24, 1, 80)), "W062NO CLAIMS ON FILE") > 0 Then
                            Return "RemoveDummyICN"
                        End If
                        For dd As Integer = 4 To 20 Step 2
                            If Trim(GetText(dd, 4, 6)) = "------" Then Exit For
                            If Trim(GetText(dd + 1, 54, 10)) = OrICN Then
                                Continue For
                            Else
                                If GetText(dd + 1, 52, 1) = "F" Then
                                    Continue For

                                ElseIf GetText(dd + 1, 52, 1) = "I" Then

                                    SendKeyes("S", dd, 2)
                                    SendEnter()
                                    apiChk()
                                    If InStr(Trim(GetText(2, 2, 3)), "GHI") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHN") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHB") > 0 Then
                                        'isMHI = False
                                        SendKeyes("MHI", 2, 2)
                                        autECLPSObj.SendKeys("[ENTER]")
                                        apiChk()
                                    End If
                                    If Trim(GetText(2, 2, 2)) <> "MH" Then
                                        PullAHI()
                                        Continue For
                                        '''''''''''''?????????????''''''''''''''''
                                    End If
                                    For cc As Integer = 10 To 18 Step 4
                                        If Trim(GetText(cc, 1, 2)) = "" Then Exit For

                                        If (InStr(Trim(GetText(cc, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(cc, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(cc, 4, 4)), "MISC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OPS") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OA") > 0) = True Then
                                            isUB = True
                                            Exit For
                                        End If
                                    Next
                                    If isUB = True Then
                                        For cc As Integer = 10 To 18 Step 4
                                            If InStr("ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,PLPRE,POD,PRE,REF,SSO,TDPD,OIM,OI,OIMEDI", Trim(GetText(cc, 4, 6))) > 0 Then
                                                Continue For
                                            End If

                                            If Trim(GetText(cc + 2, 1, 3)) = "ICN" And Trim(GetText(cc, 28, 10)) <> "0000000000" Then
                                                histICN = GetText(cc + 2, 5, 10)
                                                histDraft = GetText(cc, 28, 10)
                                                Exit For
                                            End If

                                            If cc = 18 Then
                                                If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                                                SendF8()
                                                cc = 6
                                            End If
                                        Next
                                        SendKeyes("A", 3, 15)
                                        SendEnter()
                                        apiChk()

                                        If Trim(GetText(2, 2, 2)) <> "AH" Then
                                            PullAHI()
                                        End If
                                    Else
                                        SendKeyes("A", 3, 15)
                                        SendEnter()
                                        apiChk()
                                        If Trim(GetText(2, 2, 2)) <> "AH" Then
                                            PullAHI()
                                        End If
                                    End If

                                End If
                            End If
                            If histDraft <> "" And histICN <> "" Then Exit For
                            If dd = 20 Then
                                If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                                SendF8()
                                dd = 2
                            End If
                        Next

                        If histDraft <> "" And histICN <> "" Then
                            SendClear()
                            SendClear()
                            pullEDSScreen(ICN, 2)
                            SendKeyes("A", 1, 76)
                            SendEnter()
                            apiChk()
                            If GetText(2, 2, 2) <> "MP" Then
                                Return "RemoveDummyICN"
                            End If
                            SendKeyes("EDS", 2, 2)
                            SendEnter()
                            apiChk()
                            If GetText(3, 2, 3) <> "EDS" Then
                                Return "RemoveDummyICN"
                            End If
                            SendKeyesFor01(histDraft, 4, 32)
                            SendKeyesFor01(histICN, 4, 15)
                            SendEnter()
                            apiChk()
                            WriteLog(GetTextRect)
                            If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                RemoveNotCoveredFromEDS6Point5()
                                SendEnter()
                                If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                    RemoveNotCoveredFromEDS6Point5()
                                    SendEnter()
                                End If
                                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                                    For cc As Integer = 7 To 19 Step 3
                                        If GetText(cc - 1, 8, 9) <> "         " Then
                                            autECLPSObj.SendKeys("--------", cc, 6)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 20)
                                            apiChk()
                                            autECLPSObj.SendKeys("--------", cc, 28)
                                            apiChk()
                                            autECLPSObj.SendKeys("---", cc, 42)
                                            apiChk()
                                            autECLPSObj.SendKeys("------", cc, 50)
                                            apiChk()
                                        End If
                                    Next
                                    SendEnter()
                                    apiChk()
                                End If
                            End If
                            SendKeyes("A", 1, 75)
                            SendEnter()
                            apiChk()
                            WriteLog(GetTextRect)
                            If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                                SendKeyes("N", 4, 79)
                                autECLPSObj.SendKeys("[ENTER]")
                                apiChk()
                                SendKeyes("A", 1, 75)
                                SendEnter()
                                apiChk()
                            End If
                            If GetText(2, 2, 2) <> "MP" Then
                                Return "RemoveDummyICN"
                            End If
                        Else
                            Return "RemoveDummyICN"
                        End If

                        'Else
                        '    Return ""
                    End If
                End If
                'If flag Then
                '    ROVAmount = totalAmt
                'End If
            Next

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)

        End Try
        Return val
    End Function

    Public Function StoreOriginalDraftandICN(ByVal ICN As String, ByVal Payloc As String) As String
        WriteLog("==================Entering " + GetCurrentMethod.Name + "===================")
        Dim OrICN As String = ICN
        Dim val As String = ""
        Dim histICN As String = ""
        Dim histDraft As String = ""
        Dim isUB As Boolean = False
        Try
            GetClaimDetails(CurrentClaimDetails, ICN, Payloc)
            PullAHI()
            WriteLog(GetTextRect())
            If GetText(2, 2, 2) <> "AH" Or InStr(Trim(GetText(24, 1, 80)), "W062NO CLAIMS ON FILE") > 0 Then

            End If
            For dd As Integer = 4 To 20 Step 2
                If Trim(GetText(dd, 4, 6)) = "------" Then Exit For
                If Trim(GetText(dd + 1, 54, 10)) = OrICN Then
                    Continue For
                Else
                    If GetText(dd + 1, 52, 1) = "F" Then
                        Continue For

                    ElseIf GetText(dd + 1, 52, 1) = "I" Then

                        SendKeyes("S", dd, 2)
                        SendEnter()
                        apiChk()
                        If InStr(Trim(GetText(2, 2, 3)), "GHI") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHN") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHB") > 0 Then
                            'isMHI = False
                            SendKeyes("MHI", 2, 2)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        If Trim(GetText(2, 2, 2)) <> "MH" Then
                            PullAHI()
                            Continue For
                            '''''''''''''?????????????''''''''''''''''
                        End If

                        For cc As Integer = 10 To 18 Step 4
                            If Trim(GetText(cc, 1, 80)) = "" Then Exit For
                            If (InStr(Trim(GetText(cc, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(cc, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(cc, 4, 4)), "MISC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OPS") > 0 OrElse InStr(Trim(GetText(cc, 4, 6)), "OA") > 0) = True Then
                                isUB = True
                                LastDatefromHx = GetText(10, 18, 6)
                                Exit For
                            End If
                        Next
                        If isUB = True Then
                            For cc As Integer = 10 To 18 Step 4
                                If InStr("ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,PLPRE,POD,PRE,REF,SSO,TDPD,OIM,OI,OIMEDI", Trim(GetText(cc, 4, 6))) > 0 Then
                                    Continue For
                                End If

                                If Trim(GetText(cc + 2, 1, 3)) = "ICN" And Trim(GetText(cc, 28, 10)) <> "0000000000" Then
                                    histICN = GetText(cc + 2, 5, 10)
                                    histDraft = GetText(cc, 28, 10)
                                    Exit For
                                End If

                                If cc = 18 Then
                                    If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                                    SendF8()
                                    cc = 6
                                End If
                            Next
                            SendKeyes("A", 3, 15)
                            SendEnter()
                            apiChk()

                            If Trim(GetText(2, 2, 2)) <> "AH" Then
                                PullAHI()
                            End If

                        End If
                        If isUB = False Then
                            SendKeyes("A", 3, 15)
                            SendEnter()
                            apiChk()
                            If Trim(GetText(2, 2, 2)) <> "AH" Then
                                PullAHI()
                            End If
                        End If
                    End If
                End If
                If histDraft <> "" And histICN <> "" Then Exit For
                If dd = 20 Then
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    SendF8()
                    dd = 2
                End If
            Next

            If histDraft <> "" And histICN <> "" Then
                OriginalDraft = histDraft
                OriginalICN = histICN
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception: " + ex.Message.ToString)
            WriteLog(GetTextRect)
        End Try
        Return val
    End Function
    Public Sub RemovePISLline()                     'This method is to remove PISlline from EDS6 screen- 07/07/2023 @Sourav
        Try
            Dim SafetyPageCount As Integer = 0
            WriteLog("===============ENTERING " + GetCurrentMethod.Name + "==============")
            If GetText(2, 2, 2) = "MP" Then
                autECLPSObj.SendKeys("EDS", 2, 2)
                SendEnter()
                apiChk()
            ElseIf GetText(3, 2, 5) = "EDS 6" Then

            Else
                SendClear()
                SendClear()
                pullEDSScreen(CurrICN, 2)
                SendKeyes("A", 1, 76)
                SendEnter()
                apiChk()

                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If

            For KK As Integer = 6 To 18 Step 3

                If GetText(KK, 11, 6, True) = "------" OrElse Trim(GetText(KK, 11, 6, True)) = "" Then
                    Exit For
                End If

                If (GetText(KK, 11, 6, True) = "OI") OrElse
                   (GetText(KK, 11, 6, True) = "OIM") OrElse
                   (InStr(PISLLINE, GetText(KK, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(KK, 11, 6, True)) = 0) Then

                    SendKeyes("------", KK, 11)
                    SendKeyes("------", KK, 18)
                    SendKeyes("------", KK, 25)
                    SendKeyes("-----", KK, 32)
                    SendKeyes("--", KK, 38)
                    SendKeyes("-", KK, 41)
                    SendKeyes("-", KK, 43)
                    SendKeyes("---", KK, 45)
                    SendKeyes("------", KK, 49)
                    SendKeyes("--", KK, 56)
                    SendEnter()
                    'SendKeyes("N", 4, 79)
                    'SendEnter()
                    If (GetText(KK, 11, 6, True) = "BLANK-") Then
                        SendKeyes("------", KK, 11)
                        SendEnter()
                        KK = 3
                    End If
                End If
                'If KK = 18 Then
                '    SendEnter()
                '    SendF8()
                '    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                '        Exit For
                '    Else
                '        KK = 3
                '    End If
                'End If

                If KK = 18 Then
                    WriteLog(GetTextRect())
                    Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                    SendF8()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                        Exit For
                    ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                        Exit For
                    Else
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                    End If
                    KK = 3
                End If
            Next
            autECLPSObj.SendKeys("N", 4, 79)
            apiChk()
            autECLPSObj.SendKeys("[ENTER]")
            apiChk()

            'SendKeyes("A", 1, 75)
            autECLPSObj.SendKeys("A", 1, 75)
            SendEnter()
            If InStr(GetText(24, 1, 80).Trim, "NBAL NC") > 0 Then
                RemoveNotCoveredFromEDS6Point5()
                autECLPSObj.SendKeys("[ENTER]")

                autECLPSObj.SendKeys("N", 4, 79)
                apiChk()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()

                'SendKeyes("A", 1, 75)
                autECLPSObj.SendKeys("A", 1, 75)
                SendEnter()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub implicitDenial(SuffixCount As Integer)
        Dim SafetyPageCount As Integer = 0
        If GetText(2, 2, 2) <> "MP" Then
            pullEDSScreen(CurrICN, 1)
            SendKeyes("A", 1, 76)
            apiChk()
            autECLPSObj.SendKeys("[ENTER]")
        End If
        'Going to MPC then EDS6.5 to place not covered amount

        If GetText(2, 2, 2) = "MP" Then
            SendKeyes("EDS", 2, 2)
            apiChk()
            autECLPSObj.SendKeys("[ENTER]")
            apiChk()
        End If
        Dim LC As Integer = 0
        'Dim chrg As String
        Dim RCfromArray As String = String.Empty
        Dim ChargeFromDt As Decimal = 0.0
        If GetText(3, 2, 5) = "EDS 6" Then
            For z = 0 To dt_denialRC.Rows.Count - 1
                Dim row As DataRow = dt_denialRC.Rows(z)
                If row(0) <> "" And row(0) <> Nothing Then
                    'If DenialRC(z) <> "" And DenialRC(z) <> Nothing Then
                    'RCfromArray = DenialRC(z)
                    RCfromArray = row(0)
                    ChargeFromDt = CDec(row(1))

                    For i As Integer = 6 To 18 Step 3
                        If Trim(GetText(i, 8, 2)) <> "" Then
                            LC += 1
                            If GetText(i, 45, 2) = RCfromArray And CDec(GetText(i, 49, 6) + "." + GetText(i, 56, 2)) = ChargeFromDt Then
                                'If GetText(i, 49, 6).Contains("-") Then
                                '    chrg = CDec(GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2))
                                'Else
                                '    chrg = CDec(GetText(i, 49, 6) + "." + GetText(i, 56, 2))
                                'End If
                                'If LC = LCArray(z) And charge(z) = chrg Then
                                '    SendKeyes("opn", i, 2)
                                '    WriteLog(GetTextRect)
                                '    autECLPSObj.SendKeys("[ENTER]")
                                '    apiChk()
                                '    Exit For
                                'End If
                                SendKeyes("opn", i, 2)
                                WriteLog(GetTextRect)
                                autECLPSObj.SendKeys("[ENTER]")
                                apiChk()
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                        If i = 18 Then
                            WriteLog(GetTextRect())
                            Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                            SendF8()
                            If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                Exit For
                            ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                Exit For
                            Else
                                SafetyPageCount = SafetyPageCount + 1

                                ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                If SafetyPageCount > 50 Then
                                    WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                    Exit For
                                End If
                            End If
                            i = 3
                        End If
                    Next
                    If GetText(2, 73, 7) = "EDS 6.5" Then
                        For ii As Integer = 8 To 17 Step 3
                            If Trim(GetText(ii, 7, 4)) <> "" Then
                                Dim dec As String = Trim(GetText(ii, 54, 2))
                                Dim charg As String = Trim(GetText(ii, 47, 6))
                                SendKeyes(charg, ii, 57)
                                apiChk()
                                SendKeyes(dec, ii, 64)
                                apiChk()
                                SendKeyes(RCfromArray, ii, 67)
                                apiChk()
                                SendKeyes("--------", ii, 71)
                                apiChk()
                                WriteLog(GetTextRect)

                            Else
                                SendEnter()
                                SendF3()
                                Exit For
                            End If
                            If ii = 17 Then
                                Dim screen1 As String = GetTextRect()
                                SendF8()
                                If GetTextRect() = screen1 Then
                                    SendEnter()
                                    SendF3()
                                    Exit For
                                End If
                                ii = 5
                            End If
                        Next
                        'autECLPSObj.sendkeys("[enter]")
                        apiChk()
                    End If
                Else
                    Exit For
                End If
            Next
        End If

        'WriteLog("-----------In EDS 6.5----------------------")

        'SendF3()
        WriteLog("-----------In EDS 6----------------------")
        WriteLog(GetTextRect)
        apiChk()
        SendEnter()
        SendKeyes("A", 1, 75)
        SendEnter()

    End Sub
    Public Sub MatchAttendingPhysician(ByVal HxICNList As String, ByRef FlagNotmatch As Boolean)
        WriteLog("================================ENTERING " + GetCurrentMethod.Name)
        Dim EDS1ANPI As String = String.Empty
        Dim HxMhiANPI As String = String.Empty
        Dim icn As String = String.Empty
        Dim MHIControlLine As String = String.Empty
        Dim EditWarning As String
        Dim count As Integer
        Dim FinalHxICNList As String = String.Empty
        Dim ERRORFLAG As Boolean = False
        Dim linesplitText As String = String.Empty
        'ReDim ListOfHxMatchingANPI(0)
        Dim Linecount As Integer = 0
        'Dim FlagNotmatch As Boolean = False
        Try
            pullEDSScreen(CurrICN, 1)
            EDS1ANPI = GetText(23, 57, 10)
            If EDS1ANPI.Contains("-") Then
                EDS1ANPI = EDS1ANPI.Replace("-", "")
            End If
            Dim HICNArray() As String = HxICNList.Split(",")
            For mm As Integer = LBound(HICNArray) To UBound(HICNArray)
                If HICNArray(mm).Trim <> "" Then
                    icn = HICNArray(mm)
                    'ListOfHxMatchingANPI(Linecount).HistoryICN(mm) = icn
                    If icn.Trim() = "" Then
                        Continue For
                    Else
                        SendClear()
                        SendClear()
                        SendKeyes("ret,i" + icn + ",m", 2, 2)
                        SendEnter()
                        MHIControlLine = Trim(Trim(GetText(2, 2, 75)).Replace("I" + icn, ""))
                        MHIControlLine = MHIControlLine + "BI" + icn
                        SendClear()
                        SendClear()
                        SendKeyesFor01(MHIControlLine, 2, 2)
                        SendKeyesFor01("I", 2, 4)
                        SendEnter()
                        apiChk()
                        Dim Linesplit() As String
                        Dim firstFlag As Boolean = False
                        EditWarning = String.Empty
                        count = 0
                        Do
                            If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                            If count > 20 Then Exit Do
                            If firstFlag = True Then
                                EditWarning += GetText(24, 1, 80, True)
                            End If
                            If InStr(GetText(24, 1, 80), "ATT") > 0 And firstFlag = False Then
                                EditWarning += GetText(24, 1, 80, True)
                                firstFlag = True
                                Exit Do
                            End If
                            SendEnter()
                            apiChk()
                            count += 1
                        Loop

                        If InStr(EditWarning, "ATT") > 0 Then
                            Linesplit = EditWarning.Split(",")
                            For I As Integer = LBound(Linesplit) To UBound(Linesplit)
                                linesplitText = Trim(Linesplit(I))
                                If linesplitText <> "" Then
                                    If InStr(linesplitText, "ATT") > 0 Then
                                        HxMhiANPI = Microsoft.VisualBasic.Right(linesplitText, 10)
                                        Exit For
                                    End If
                                End If
                            Next
                        Else
                            '''''''''''''''''''''
                            ERRORFLAG = True
                            '''''''''''''''''''''
                        End If

                        If EDS1ANPI <> HxMhiANPI Then
                            ''''''''''''''''''''''''''''''''''''''''''''''''
                            FlagNotmatch = True
                            'ListOfHxMatchingANPI(Linecount).AttPhyMatchFlag(mm) = False
                            dt_ListOfHxMatchingANPI.Rows.Add(icn, False)
                        Else
                            If InStr(FinalHxICNList, icn) = 0 Then
                                If FinalHxICNList = String.Empty Then
                                    FinalHxICNList = icn
                                Else
                                    FinalHxICNList = FinalHxICNList + "," + icn
                                End If
                            End If
                            'ListOfHxMatchingANPI(Linecount).AttPhyMatchFlag(mm) = True
                            dt_ListOfHxMatchingANPI.Rows.Add(icn, True)

                        End If

                    End If
                End If
            Next
        Catch ex As Exception
            WriteLog("exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try
    End Sub
    Sub GenerateDt()
        dt_ListOfHxMatchingANPI.Columns.Add("HxICN", GetType(String))
        dt_ListOfHxMatchingANPI.Columns.Add("ANPIMatchFlag", GetType(Boolean))

    End Sub

    Sub GenerateDt_Mod50()
        dt_LinesWith50Modifier.Columns.Add("EDS6LineNumber", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("EDS6.5LineNumber", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("REVCode", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("ProcCode", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("Charge", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("lineCount", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("PageNumber", GetType(String))
        dt_LinesWith50Modifier.Columns.Add("Mod50Flag", GetType(Boolean))
        dt_LinesWith50Modifier.Columns.Add("isItLastLineOfEDS6", GetType(Boolean))
    End Sub
    Sub Initialize_dt_denialRC(ByRef dt_denialRC)
        dt_denialRC.Columns.Add("Denial_RC", GetType(String))
        dt_denialRC.Columns.Add("Charge", GetType(String))
        dt_denialRC.columns.add("Suffix", GetType(String))
    End Sub

    Public Function GetCopayNBRfromHX(icnlist As String, CurrICN As String) As String      'Added 27/06/2023 --to calculate NBR from paid HxClaim and place in current claim
        Dim icn As String = String.Empty
        Dim NC As String = String.Empty
        Dim Charge As String = String.Empty
        Dim nbr As String = ""
        Dim CopayCount As Integer = 0
        Try
            Dim HxICN() = icnlist.Split(",")

            PullAHI()
            'If InStr(Trim(GetText(24, 1, 80)), "NO CLAIMS ON FILE") > 0 Then Return "NO CLAIMS ON FILE"
            For dd As Integer = 4 To 20 Step 2
                If Trim(GetText(dd, 4, 6)) = "------" Then Exit For

                If Trim(GetText(dd + 1, 54, 10)) = CurrICN Then
                    Continue For
                Else
                    For mm As Integer = LBound(HxICN) To UBound(HxICN)
                        If HxICN(mm).Trim <> "" Then
                            icn = HxICN(mm)

                            If Trim(GetText(dd, 19, 8)) <> Trim(GetText(dd, 54, 8)) And GetText(dd + 1, 54, 10) = icn And InStr("69,70,71,74", GetText(dd, 64, 2)) = 0 Then
                                SendKeyesFor01("S", dd, 2)
                                SendEnter()
                                apiChk()

                                If Trim(GetText(2, 2, 2)) = "MH" Then
                                    For ff As Integer = 12 To 20 Step 4
                                        If InStr(Trim(GetText(ff, 1, 3)), "ICN") > 0 And Trim(GetText(ff, 5, 10)) <> "0000000000" And ff <> 12 Then

                                            If InStr("69,70,71,74", GetText(ff - 6, 36, 2)) = 0 Then
                                                If CInt(Trim(GetText(10, 25, 3))) < 10 Then
                                                    nbr = "00" + Trim(GetText(10, 27, 1))
                                                ElseIf CInt(Trim(GetText(10, 25, 3))) > 9 And CInt(Trim(GetText(10, 25, 3))) < 100 Then
                                                    nbr = "0" + Trim(GetText(10, 26, 1))
                                                Else
                                                    nbr = Trim(GetText(10, 25, 3))
                                                End If
                                            End If
                                        End If

                                        'If InStr(Trim(GetText(ff, 1, 3)), "ICN") > 0 And Trim(GetText(ff, 5, 10)) <> "0000000000" And ff = 12 Then
                                        '    SendF7()
                                        '    If InStr("69,70,71,74", GetText(18, 36, 2)) = 0 Then
                                        '        If CInt(Trim(GetText(10, 25, 3))) < 10 Then
                                        '            nbr = "00" + Trim(GetText(10, 27, 1))
                                        '        ElseIf CInt(Trim(GetText(10, 25, 3))) > 9 And CInt(Trim(GetText(10, 25, 3))) < 100 Then
                                        '            nbr = "0" + Trim(GetText(10, 26, 1))
                                        '        Else
                                        '            nbr = Trim(GetText(10, 25, 3))
                                        '        End If
                                        '    End If
                                        '    SendF8()
                                        '    If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Return nbr
                                        'End If

                                        If ff = 20 Then
                                            'SendF8()
                                            If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Return nbr
                                            SendF8()
                                            ff = 8
                                        End If
                                    Next
                                End If
                            End If

                        End If
                    Next
                End If

                If dd = 20 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    SendF8()
                    dd = 2
                End If
            Next

        Catch ex As Exception
            WriteLog("exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try
    End Function

    Public Sub PlaceNCopayNBRinMPC()
        Try
            If GetText(2, 2, 2) = "MP" Then
                For I = 10 To 22 Step 2
                    If GetText(I, 4, 6) = "------" Then Exit For
                    If GetText(I, 4, 6) = "NCOPAY" Then
                        SendKeyesFor01(NCopayNBR, I, 25)
                        SendEnter()
                        WriteLog("++++======" + GetCurrentMethod.Name + "======" + vbNewLine + GetTextRect())
                        apiChk()
                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog("exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try
    End Sub
    Public Function CheckFinancialImpact(icn As String) As Boolean          '--Added 30th June by Sourav to determine financial impact of history icn
        Dim financialImpact As Boolean = False
        Dim totalBillChargeDifference As Boolean = False
        Dim totalAllowableDifference As Boolean = False
        Dim paidDifference As Boolean = False
        Dim hxtotalCharge As Double = 0.0
        Try
            If GetText(3, 75, 5) = "EDS 2" Or GetText(6, 2, 5) = "EDS 1" Then
                SendKeyes("A", 1, 76)
                SendEnter()
                apiChk()

            Else
                pullEDSScreen(icn, 2)
                SendKeyes("A", 1, 76)
                SendEnter()
                apiChk()
            End If
            'To check total bill charge difference

            If CurrentClaimDetails.TotalCharge = hxtotalCharge Then
                totalBillChargeDifference = True
            End If

            'To check total allowable difference
            Dim currentClaimTotalAllowable As Double = 0.0
            For i As Integer = LBound(CurrentClaimPaymentScreen(0).Allowable) To UBound(CurrentClaimPaymentScreen(0).Allowable)
                If CurrentClaimPaymentScreen(0).Allowable(i) = Nothing Then
                    Exit For
                Else
                    currentClaimTotalAllowable = currentClaimTotalAllowable + CDbl(CurrentClaimPaymentScreen(0).Allowable(i))
                End If
            Next

            'To check total paid difference
            Dim totalPaidinCurrentClaim As Double = 0.0

        Catch ex As Exception

        End Try

        Return financialImpact
    End Function
    Public Function IsOIOIM(ByRef MyList() As PossibleICN, ByRef MyClaim() As ClaimFromPaymentScreen) As Boolean
        For I As Integer = LBound(MyClaim) To UBound(MyClaim)
            For J As Integer = LBound(MyClaim(I).POS) To UBound(MyClaim(I).POS)
                If (InStr("OI,OIM,OIMEDI", MyClaim(I).ProcCode(J).Trim) > 0 And MyClaim(I).ProcCode(J).Trim <> "") OrElse
                   (InStr("59/UL/51/50/56/KN/P3/P5/P6/P7/P8/UR/SA/IP/3Q/4C/5M/57/F8/M9/L7/WW/L9/KR/23/KS/SY/59/G6", MyClaim(I).RemarkCode(J)) > 0 And MyClaim(I).RemarkCode(J).Trim <> "") Then
                    Return True
                End If
            Next
        Next
        For K As Integer = LBound(MyList) To UBound(MyList)
            For I As Integer = LBound(MyList(K).TempClaimFromPaymentScreen) To UBound(MyList(K).TempClaimFromPaymentScreen)
                For J As Integer = LBound(MyList(K).TempClaimFromPaymentScreen(I).POS) To UBound(MyList(K).TempClaimFromPaymentScreen(I).POS)
                    If (InStr("OI,OIM,OIMEDI", MyList(K).TempClaimFromPaymentScreen(I).ProcCode(J).Trim) > 0 And MyList(K).TempClaimFromPaymentScreen(I).ProcCode(J).Trim <> "") OrElse
                       (InStr("59/UL/51/50/56/KN/P3/P5/P6/P7/P8/UR/SA/IP/3Q/4C/5M/57/F8/M9/L7/WW/L9/KR/23/KS/SY/59/G6", MyList(K).TempClaimFromPaymentScreen(I).RemarkCode(J)) > 0 And MyList(K).TempClaimFromPaymentScreen(I).RemarkCode(J).Trim <> "") Then
                        Return True
                    End If
                Next
            Next
        Next
        Return False
    End Function
    Public Function CheckRC39() As Boolean
        WriteLog("Entering " + GetCurrentMethod.Name)
        Dim flag = False
        Dim list = "39,SF,UL,UW"
        Dim SafetyPageCount As Integer = 0

        Try
            'To check in EDS 6
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
                apiChk()
                For I = 6 To 18 Step 3
                    If GetText(I, 2, 3) <> "---" Then Exit For
                    If InStr(list, GetText(I + 1, 20, 2)) > 0 Then
                        flag = True
                        Return flag
                    End If

                    If I = 18 Then
                        WriteLog(GetTextRect())
                        Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                        SendF8()
                        If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                            Exit For
                        Else
                            SafetyPageCount = SafetyPageCount + 1

                            ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                            If SafetyPageCount > 50 Then
                                WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                Exit For
                            End If
                        End If
                        I = 3
                    End If
                Next
                WriteLog(GetTextRect)
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
                WriteLog(GetTextRect)
                removeNBAL()
                If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    apiChk()
                End If
            End If

            'To check in MPC screen
            For I = 10 To 22 Step 2
                If GetText(I, 4, 6) = "------" Then Exit For
                If InStr(list, GetText(I, 36, 2)) > 0 Then
                    flag = True
                    Return flag
                End If
            Next

        Catch ex As Exception
            WriteLog("Exception: " + ex.Message + "from " + GetCurrentMethod.Name)
        Finally
            WriteLog("Exiting " + GetCurrentMethod.Name)
        End Try

        Return flag
    End Function

    Public Sub ProcessClaimAsFreshSubmission(ByVal ICN As String, ByRef FlagError As Boolean)
        Try
            WriteLog("Entering " + GetCurrentMethod.Name)
            Dim EditWarning As String = String.Empty
            Dim flagOKToPay As Boolean = False

            'Pull MPC
            If GetText(2, 2, 2) <> "MP" Then
                pullEDSScreen(ICN, 2)
                SendKeyes("A", 1, 76)
                SendEnter()
                apiChk()
            End If

            CheckBypassHistoryField()
            Dim Count3 As Integer = 0
            Do
                SendEnter()
                apiChk()
                If Count3 > 50 Then Exit Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                EditWarning = EditWarning & GetText(24, 1, 80, True)
                Count3 = Count3 + 1
            Loop
            'Added below loop to fill empty RC fields  
            For I As Integer = 10 To 22 Step 2
                If GetText(I, 1, 80).Replace("-", "").Trim <> "" Then
                    If GetText(I, 36, 2) = "--" And CDbl(GetText(I, 41, 5).Replace("-", "0") + "." + GetText(I, 47, 2).Replace("-", "0")) <> CDbl(GetText(I, 50, 5).Replace("-", "0") + "." + GetText(I, 56, 2).Replace("-", "0")) Then
                        If InStr(EditWarning, "W1616") > 0 Then
                            SendKeyes("D2", I, 36)
                            apiChk()
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                            apiChk()
                        Else
                            strResearchStatus = "SKIPPED"
                            strResearchComments = "Skipped: RC field empty and W1616 not present"
                            FlagError = True
                            Exit Sub
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            If InStr(EditWarning, "W118 OK TO PAY") > 0 OrElse
           (InStr(EditWarning, "WE") = 0 AndAlso InStr(EditWarning, ",E") = 0 AndAlso InStr(EditWarning, ", E") = 0) Then
                flagOKToPay = True
            End If
            If flagOKToPay Then
                If isNCopayNBR = True Then

                    If GetText(2, 2, 2) = "MP" Then
                        For I = 10 To 22 Step 2
                            If GetText(I, 4, 6) = "------" Then Exit For
                            If GetText(I, 4, 6) = "NCOPAY" Then
                                SendKeyesFor01(NCopayNBR, I, 25)
                                SendEnter()
                                apiChk()
                            End If
                        Next
                    End If

                End If
                If ICESDenialFlag = False And flagDenial = False Then
                    If MatchAllowable() = False Then
                        strResearchComments = "Allowable mismatch before MPP"
                        FlagError = True
                        Exit Sub
                    End If
                End If
                WriteLog("-------------------------BEFORE MPP----------------------------")
                WriteLog(GetTextRect)
                SendF12()
                If CheckClaimProcessedBeforeMPP() = True Then
                    strResearchComments = "Current claim is already processed"
                    FlagError = True
                    Exit Sub
                End If
                SendF12()
                SendMPP()
                If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                    If isNCopayNBR = True Then
                        If isC68REV = True Then
                            strResearchComments = "Claim process by tool; Therapy with NCOPAY"
                        Else
                            strResearchComments = "Claim process by tool; NCopay present"
                        End If
                    Else
                        strResearchComments = "Claim process by tool"

                    End If

                    WriteLog("-------------------------AFTER MPP----------------------------")
                    WriteLog(GetTextRect)
                    PaymentComment = "FRESH CLAIM PAID"
                End If
                If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                    strResearchStatus = "SKIPPED"
                    strResearchComments = GetText(24, 1, 80, True)
                    FlagError = True
                    Exit Sub
                End If
                CaptureScreenLog()
            Else
                If InStr(EditWarning, "WE1434INVESTIGATE COB") > 0 Then
                    strResearchStatus = "SKIPPED"
                    strResearchComments = "WE1434INVESTIGATE COB"
                    FlagError = True
                    Exit Sub
                Else
                    strResearchStatus = "SKIPPED"
                    strResearchComments = EditWarning
                    FlagError = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            WriteLog("Exception: " + ex.Message + "from " + GetCurrentMethod.Name)
        End Try

    End Sub

    Public Sub MARSEDITINGAPPLIES()
        Try
            If GetText(2, 2, 3) = "MPC" Then
                SendKeyesFor01("MPR", 2, 2)
                apiChk()
                WriteLog(GetTextRect)
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
                WriteLog(GetTextRect)

                If InStr(GetText(24, 1, 80).Trim, "W3067SENT TO MARS EDITING") Then
                    strResearchComments = "Sent to MARS EDITING"
                Else
                    strResearchComments = "Unable to pend to MARS EDITING"
                End If
            End If
        Catch ex As Exception
            WriteLog("Exception: " + ex.Message + "from " + GetCurrentMethod.Name)
        End Try
    End Sub

    Public Function GetCopayNBRforTherapyClaims(ByRef therapyNBRError As String)      'Modified 29th July,2024: to determine based on revenue codes and corresponding modifiers
        WriteLog("=================Inside " + GetCurrentMethod.Name + " ==========================")
        Dim nbrCount As Integer = 0
        Dim nbr As String = String.Empty
        Dim DOS As String = ""
        Try
            For i As Integer = 0 To UBound(CurrentICNEDS1)
                If CurrentICNEDS1(i).ServiceLine.LineTranslation = "C68REV" Then
                    isC68REV = True
                    For j As Integer = 0 To UBound(CurrentICNEDS1(i).ServiceLine.DateofService)
                        If CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) >= 420 And CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) <= 429 Then
                            If CurrentICNEDS1(i).ServiceLine.M1(j) = "GP" Or CurrentICNEDS1(i).ServiceLine.M2(j) = "GP" Then
                                If InStr(DOS, CurrentICNEDS1(i).ServiceLine.DateofService(j)) = 0 Then
                                    nbrCount += 1
                                    DOS = DOS + CurrentICNEDS1(i).ServiceLine.DateofService(j) + ","
                                End If
                            Else
                                therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS"
                            End If
                        End If

                        If CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) >= 430 And CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) <= 439 Then
                            If CurrentICNEDS1(i).ServiceLine.M1(j) = "GO" Or CurrentICNEDS1(i).ServiceLine.M2(j) = "GO" Then
                                If InStr(DOS, CurrentICNEDS1(i).ServiceLine.DateofService(j)) = 0 Then
                                    nbrCount += 1
                                    DOS = DOS + CurrentICNEDS1(i).ServiceLine.DateofService(j) + ","
                                End If
                            Else
                                therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS"
                            End If
                        End If

                        If CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) >= 440 And CInt(CurrentICNEDS1(i).ServiceLine.RevCode(j)) <= 449 Then
                            If CurrentICNEDS1(i).ServiceLine.M1(j) = "GN" Or CurrentICNEDS1(i).ServiceLine.M2(j) = "GN" Then
                                If InStr(DOS, CurrentICNEDS1(i).ServiceLine.DateofService(j)) = 0 Then
                                    nbrCount += 1
                                    DOS = DOS + CurrentICNEDS1(i).ServiceLine.DateofService(j) + ","
                                End If
                            Else
                                therapyNBRError = "Therapy billed without GP/GO/GN MODIFIERS"
                            End If
                        End If

                        'If InStr(DOS, CurrentICNEDS1(i).ServiceLine.DateofService(j)) = 0 Then
                        '    nbrCount += 1
                        '    DOS = DOS + CurrentICNEDS1(i).ServiceLine.DateofService(j) + ","
                        'End If
                    Next
                End If
            Next
            If isC68REV = False Then
                nbrCount = 1
            End If
            WriteLog("DOS: " + DOS)
            WriteLog("Nbrcount = " + nbrCount.ToString)
            If nbrCount < 10 Then
                nbr = "00" + nbrCount.ToString
            ElseIf nbrCount > 9 And nbrCount < 100 Then
                nbr = "0" + nbrCount.ToString
            Else
                nbr = nbrCount.ToString
            End If
        Catch ex As Exception
            WriteLog("Exception: " + ex.Message + "from " + GetCurrentMethod.Name)

        End Try

        Return nbr
    End Function

    Public Function MatchAllowable() As Boolean
        Dim isMatch As Boolean = False
        Try
            WriteLog("==================ENTERING" + GetCurrentMethod.Name + " before MPP ========================")

            Dim Allowable As String = String.Empty
            Dim PSAllowable As Decimal = 0.0
            If GetText(2, 2, 2) = "MP" Then
                For dd As Integer = 11 To 23 Step 2
                    Dim CheckLine As String = String.Empty
                    CheckLine = Trim(GetText(dd - 1, 11, 6))
                    If CheckLine = "------" Then Exit For

                    If Trim(GetText(dd - 1, 4, 6)) = "NCOPAY" Then
                        Allowable = Trim(GetText(dd - 1, 50, 5)) + "." + Trim(GetText(dd - 1, 56, 2))
                    Else
                        Allowable = Trim(GetText(dd, 30, 5)) + "." + Trim(GetText(dd, 36, 2))
                    End If
                    If Allowable = ".00" Then
                        Allowable = "0.00"
                    End If
                    If InStr(Allowable, "-") > 0 Then
                        Allowable = Allowable.Replace("-", "0")
                    End If
                    PSAllowable += CDec(Allowable)
                Next
            End If
            WriteLog("Total allowable in MPP screen= " + PSAllowable.ToString())
            If CDec(allowableFromOCI) = PSAllowable Then
                isMatch = True
            Else
                isMatch = False
            End If
            WriteLog("IS Final allowable from MPP screen matching: " + isMatch.ToString())

        Catch ex As Exception

        End Try
        Return isMatch
    End Function

    Public Sub EditWarningCheck(Edit As String)

        Try
            If InStr(Edit, "W870") > 0 Then
                strResearchComments = "W870 present"
            End If
            If InStr(Edit, "W1444") > 0 Then
                strResearchComments = "W1444 present"
            End If
            If InStr(Edit, "W1445") > 0 Then
                strResearchComments = "W1445 present"
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub CheckBypassHistoryField()            'Added on 3rd March,24 to send Y in bypass history field
        Try
            If GetText(2, 2, 2) <> "MP" Then
                If GetText(6, 2, 5) <> "EDS 1" Then
                    pullEDSScreen(CurrICN, 1)
                End If
                SendKeyes("A", 1, 76)
                SendEnter()
                SendKeyes("EDS", 2, 2)
                SendEnter()
            Else
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If

            If GetText(3, 2, 5) = "EDS 6" Then
                For dd As Integer = 6 To 18 Step 3
                    If (GetText(dd, 11, 6, True) = "OI") OrElse
                   (GetText(dd, 11, 6, True) = "OIM") OrElse (GetText(dd, 11, 6, True) = "BLANK-") OrElse
                   (InStr(PISLLINE, GetText(dd, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(dd, 11, 6, True)) = 0) Then
                        RemovePISLline()
                        apiChk()
                        If GetText(2, 2, 2) = "MP" Then
                            autECLPSObj.SendKeys("EDS", 2, 2)
                            SendEnter()
                            apiChk()
                        End If
                    End If
                    SendKeyes("opn", 6, 2)
                    SendEnter()
                    apiChk()
                    If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                        For cc As Integer = 7 To 19 Step 3
                            If GetText(cc - 1, 8, 9) <> "         " Then
                                autECLPSObj.SendKeys("--------", cc, 6)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 20)
                                apiChk()
                                autECLPSObj.SendKeys("--------", cc, 28)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 42)
                                apiChk()
                                autECLPSObj.SendKeys("------", cc, 50)
                                apiChk()
                            End If
                        Next
                        SendEnter()
                        apiChk()
                    End If
                    If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                        SendKeyesFor01(OriginalDraft, 4, 32)
                        SendKeyesFor01(OriginalICN, 4, 15)
                        SendEnter()
                        apiChk()
                        SendKeyes("A", 1, 75)
                        SendEnter()
                        apiChk()
                        If GetText(2, 2, 2) = "MP" Then
                            autECLPSObj.SendKeys("EDS", 2, 2)
                            SendEnter()
                            apiChk()
                        End If
                    End If
                    If GetText(3, 2, 5) = "EDS 6" Then
                        SendKeyes("opn", 6, 2)
                        SendEnter()
                        apiChk()
                    End If
                    If GetText(2, 73, 7) = "EDS 6.5" Then
                        SendKeyes("Y", 2, 49)
                        SendEnter()
                        WriteLog(GetTextRect())
                        apiChk()
                        SendF3()
                    End If
                    Exit For
                Next
            End If
            SendKeyes("A", 1, 75)
            SendEnter()

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name)
        End Try
    End Sub

    Public Sub CheckBypassHistoryField(suffix As Integer)            'Added on 3rd March,24 to send Y in bypass history field
        Try
            pullEDSScreen(CurrICN, 1)
            SendKeyes("A" & suffix, 1, 76)
            SendEnter()
            SendKeyes("EDS", 2, 2)
            SendEnter()

            If GetText(3, 2, 5) = "EDS 6" Then
                For dd As Integer = 6 To 18 Step 3
                    If (GetText(dd, 11, 6, True) = "OI") OrElse
                   (GetText(dd, 11, 6, True) = "OIM") OrElse (GetText(dd, 11, 6, True) = "BLANK-") OrElse
                   (InStr(PISLLINE, GetText(dd, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(dd, 11, 6, True)) = 0) Then
                        RemovePISLline()
                        apiChk()
                        If GetText(2, 2, 2) = "MP" Then
                            autECLPSObj.SendKeys("EDS", 2, 2)
                            SendEnter()
                            apiChk()
                        End If
                    End If
                    SendKeyes("opn", 6, 2)
                    SendEnter()
                    apiChk()
                    If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Then
                        For cc As Integer = 7 To 19 Step 3
                            If GetText(cc - 1, 8, 9) <> "         " Then
                                autECLPSObj.SendKeys("--------", cc, 6)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 20)
                                apiChk()
                                autECLPSObj.SendKeys("--------", cc, 28)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 42)
                                apiChk()
                                autECLPSObj.SendKeys("------", cc, 50)
                                apiChk()
                            End If
                        Next
                        SendEnter()
                        apiChk()
                    End If
                    If InStr(GetText(24, 2, 80).Trim, "ORG ICN OR DRAFT INVALID VALUE") > 0 Then
                        SendKeyesFor01(OriginalDraft, 4, 32)
                        SendKeyesFor01(OriginalICN, 4, 15)
                        SendEnter()
                        apiChk()
                        SendKeyes("A", 1, 75)
                        SendEnter()
                        apiChk()
                        If GetText(2, 2, 2) = "MP" Then
                            autECLPSObj.SendKeys("EDS", 2, 2)
                            SendEnter()
                            apiChk()
                        End If
                    End If
                    If GetText(3, 2, 5) = "EDS 6" Then
                        SendKeyes("opn", 6, 2)
                        SendEnter()
                        apiChk()
                    End If
                    If GetText(2, 73, 7) = "EDS 6.5" Then
                        SendKeyes("Y", 2, 49)
                        SendEnter()
                        WriteLog(GetTextRect())
                        apiChk()
                        SendF3()
                    End If
                    Exit For
                Next
            End If
            SendKeyes("A", 1, 75)
            SendEnter()

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name)
        End Try
    End Sub
#Region "Medicare check 05.02.2024"
    Public Function MedicareCheck(CurrentOII As OIIDetails) As Boolean
        Dim _checkisOIM As Boolean = False
        Try
            For K As Integer = LBound(CurrentICNEDS1) To UBound(CurrentICNEDS1)
                Dim _FDOS As String = CurrentICNEDS1(K).FDate
                Dim MOnth_FDS As String = Convert.ToInt32(_FDOS.Substring(0, 2))
                Dim day1_FDS As String = Convert.ToInt32(_FDOS.Substring(2, 2))
                Dim year1_FDS As String = Convert.ToInt32(_FDOS.Substring(4, 2))
                Dim totalday_FDS As String = MOnth_FDS.ToString() + "/" + day1_FDS.ToString() + "/" + year1_FDS.ToString()
                Dim dtrehxicn_FDS As DateTime
                DateTime.TryParse(totalday_FDS, dtrehxicn_FDS)

                Dim _LDOS As String = CurrentICNEDS1(K).LDate
                Dim MOnth_LDOS As String = Convert.ToInt32(_LDOS.Substring(0, 2))
                Dim day1_LDOS As String = Convert.ToInt32(_LDOS.Substring(2, 2))
                Dim year1_LDOS As String = Convert.ToInt32(_LDOS.Substring(4, 2))
                Dim totalday_LDOS As String = MOnth_LDOS.ToString() + "/" + day1_LDOS.ToString() + "/" + year1_LDOS.ToString()
                Dim dtrehxicn_LDOS As DateTime
                DateTime.TryParse(totalday_LDOS, dtrehxicn_LDOS)

                If CurrentOII.OI IsNot Nothing Then
                    For I As Integer = LBound(CurrentOII.OI) To UBound(CurrentOII.OI)
                        If CurrentOII.OI(I).Indicator <> "N" AndAlso CurrentOII.OI(I).Indicator <> "-" Then
                            Dim dtrehxicn_FDS_Medi As DateTime
                            Dim _ash As String = CurrentOII.OI(I).EffectiveDate
                            DateTime.TryParse(_ash, dtrehxicn_FDS_Medi)
                            Dim _Canc As String = CurrentOII.OI(I).CancellationDate
                            Dim dtrehxicn_FDS_Can As DateTime
                            DateTime.TryParse(_Canc, dtrehxicn_FDS_Can)
                            If (dtrehxicn_FDS > dtrehxicn_FDS_Medi) AndAlso (dtrehxicn_LDOS < dtrehxicn_FDS_Can) Then
                                _checkisOIM = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If CurrentOII.Medicare IsNot Nothing Then
                    For I As Integer = LBound(CurrentOII.Medicare) To UBound(CurrentOII.Medicare)
                        If CurrentOII.Medicare(I).Indicator <> "N" AndAlso CurrentOII.Medicare(I).Indicator <> "-" Then
                            Dim dtrehxicn_FDS_Medi As DateTime
                            Dim _ash As String = CurrentOII.Medicare(I).EffectiveDate
                            DateTime.TryParse(_ash, dtrehxicn_FDS_Medi)
                            Dim _Canc As String = CurrentOII.Medicare(I).CancellationDate
                            Dim dtrehxicn_FDS_Can As DateTime
                            DateTime.TryParse(_Canc, dtrehxicn_FDS_Can)
                            If (dtrehxicn_FDS > dtrehxicn_FDS_Medi) AndAlso (dtrehxicn_LDOS < dtrehxicn_FDS_Can) Then
                                _checkisOIM = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If _checkisOIM Then
                    Exit For
                End If
            Next




        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name)
        End Try
        Return _checkisOIM
    End Function
#End Region

#Region "To verify iCES denial RC"
    Public Function iCESRcVerify(RC As String) As Boolean
        Dim verified As Boolean = False
        Try
            autECLPSObj.setcursorpos(2, 1)
            apiChk()
            SendKeyes("[pa2]")
            If InStr(GetText(1, 1, 80), "ENTER SIGN-ON ID") > 0 Then
                SendF4()
            End If
            SendKeyes("RC," & RC, 23, 12)
            SendEnter()
            If InStr(GetText(1, 1, 80).ToUpper(), "ICES") > 0 Or InStr(GetText(1, 1, 80), "iCES") > 0 Or InStr(GetText(1, 1, 80), "Practice Review System/iCES") > 0 Or InStr(GetText(1, 1, 80), "MCR -") > 0 Then
                verified = True
            End If

            SendF3()
        Catch ex As Exception
            WriteLog("Exception :" + ex.Message + " from method " + GetCurrentMethod.Name)
        End Try

        Return verified
    End Function
#End Region

#Region "To remove negative balance from EDS"
    Public Function removeNBAL() As Boolean
        Dim isRemoved As Boolean = False
        Try
            If GetText(3, 2, 5) = "EDS 6" Then
                If InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                    RemoveNotCoveredFromEDS6Point5()
                    SendEnter()
                    If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") Or InStr(GetText(24, 2, 80), "NBAL NC") > 0 Then
                        For cc As Integer = 7 To 19 Step 3
                            If GetText(cc - 1, 8, 9) <> "         " Then
                                autECLPSObj.SendKeys("--------", cc, 6)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 20)
                                apiChk()
                                autECLPSObj.SendKeys("--------", cc, 28)
                                apiChk()
                                autECLPSObj.SendKeys("---", cc, 42)
                                apiChk()
                                autECLPSObj.SendKeys("------", cc, 50)
                                apiChk()
                            End If
                        Next
                        SendEnter()
                        apiChk()
                    End If
                End If
                If GetText(24, 2, 80).Trim.Contains("NC1 AND RC1 CONFLICT") <> True And GetText(24, 2, 80).Trim.Contains("NO MORE RECORDS") <> True And GetText(24, 2, 80).Trim.Contains("PLEASE ENTER HIST INDICATOR") <> True And GetText(24, 2, 80).Trim <> "" And InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") = 0 Then
                    strResearchComments = "Manual intervention required: " + GetText(24, 2, 80).Trim
                    SendClear()
                    If GetText(2, 2, 2) <> "MP" Then SendClear()
                    Return isRemoved
                    Exit Function
                End If
            End If
            If GetText(3, 2, 5) = "EDS 6" Then
                SendKeyes("A", 1, 75)
                SendEnter()
                If InStr(GetText(24, 1, 80), "NBAL NC") = 0 Then
                    isRemoved = True
                End If
            End If
        Catch ex As Exception
            WriteLog("Exception as " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
        Return isRemoved
    End Function

#End Region

#Region "To check for iCES denial RC in history claims// If found skip claim"
    Public Function checkForDenialRCinHxClaim(hxClaimList As String) As Boolean     'Added 18th May,2024 by Sourav
        Dim isFound As Boolean = False
        Dim hxICN As String = String.Empty
        Dim historyClaims(20) As Claim
        Try
            Dim HICNArray() As String = hxClaimList.Split(",")
            For ff As Integer = LBound(HICNArray) To UBound(HICNArray)
                If HICNArray(ff) <> "" Then
                    hxICN = HICNArray(ff)
                    SendClear()
                    SendClear()
                    SendKeyes("MEI,I" + hxICN + ",1", 2, 2)
                    SendEnter()
                    apiChk()
                    apiChk()

                    PullMHI(hxICN)
                    ReadMHI(historyClaims)
                    For I As Integer = LBound(historyClaims) To UBound(historyClaims)
                        If historyClaims(I).ClaimNumber Is Nothing Then Exit For
                        For j As Integer = LBound(historyClaims(I).RemarkCode) To UBound(historyClaims(I).RemarkCode)
                            If InStr("69,70,71,74,D2", historyClaims(I).RemarkCode(j)) = 0 Then
                                isFound = iCESRcVerify(historyClaims(I).RemarkCode(j))
                                If isFound = True Then Return isFound
                            End If
                        Next
                    Next
                End If
            Next

        Catch ex As Exception
            WriteLog("Exception :" + ex.Message + " from method " + GetCurrentMethod.Name)
        End Try
        Return isFound
    End Function
#End Region

#Region "Generate hard edit so that AUTO doesn't pay history voided claims that are skipped"
    Public Sub GenerateHardEdit()
        Try
            If GetText(2, 2, 3) = "MPC" Then
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If
            apiChk()
            If GetText(3, 2, 5) = "EDS 6" Then
                For dd As Integer = 6 To 18 Step 3

                    If Trim(GetText(dd, 18, 6)) = "" Then
                        Exit For
                    End If
                    autECLPSObj.SendKeys("opn", dd, 2)
                    SendEnter()
                    apiChk()
                    For hh = 8 To 17 Step 3
                        If Trim(GetText(hh, 7, 4)) = "" Then Exit For

                        autECLPSObj.SendKeys("Y", hh + 2, 56)
                        SendEnter()
                        apiChk()
                        autECLPSObj.SendKeys("6", hh + 2, 43)
                        SendEnter()
                        apiChk()

                        If hh = 17 Then
                            Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                            SendF8()
                            If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                SendF3()
                                Exit For
                            ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                                SendF3()
                                Exit For
                            Else
                                hh = 5
                            End If
                        End If
                    Next

                    SendF3()
                    Exit For
                    If dd = 18 Then
                        SendEnter()
                        If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                            SendKeyes("N", 4, 79)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                            Exit For
                        Else
                            dd = 3
                        End If
                    End If
                Next

                'A over to MPC
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
                If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                    SendKeyes("N", 4, 79)
                    SendEnter()
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    apiChk()
                End If
            End If
        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name)
        End Try
    End Sub
#End Region

    Public Function CheckClaimProcessedBeforeMPP() As Boolean           'Added 17th June,24: To stop duplicate processing of same claim, we are checking if the current claim MHI is created before MPP
        Dim isProcessed As Boolean = False
        Try
            SendClear()
            SendKeyes("RET,I" + CurrICN + ",M", 2, 2)
            SendEnter()

            If GetText(4, 1, 80).Replace(" ", "").Trim <> "" Then
                isProcessed = True
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " :" + ex.Message)
        End Try
        Return isProcessed
    End Function


#Region "Split OCI with 50 modifier"
    Public Sub splitFor50Modifier()
        Try
            Dim lineCount As Integer = 0
            Dim pageNumber As Integer = 0
            Dim copiedPos = String.Empty
            Dim copiedRev = String.Empty
            Dim copiedProc = String.Empty
            Dim copiedFromDt = String.Empty
            Dim copiedToDt = String.Empty
            Dim copiedUnits = String.Empty
            Dim copiedCharge = String.Empty
            Dim chargeCoverted As Decimal = 0.0
            Dim finalCharge As Decimal = 0.0
            Dim copiedFinalChargeWhole As String = String.Empty
            Dim copiedFinalChargeDec As String = String.Empty
            Dim lineNumber As Integer = 0
            Dim dt_LinesWith50ModinOCI As New DataTable
            dt_LinesWith50ModinOCI.Columns.Add("sPOS", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sREVCode", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sProcCode", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sFromDt", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sTodt", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sUnits", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sChargeWhole", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("sChargeDec", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("iLineNumber", GetType(Integer))
            'dt_LinesWith50ModinOCI.Columns.Add("sPageNumber", GetType(String))
            dt_LinesWith50ModinOCI.Columns.Add("Mod50Flag", GetType(Boolean))
            dt_LinesWith50ModinOCI.Columns.Add("isItLastLineOfOCI", GetType(Boolean))

            'Go to first page
            SendKeyesFor01("OCB", 2, 2)
            While GetText(10, 78, 3) <> "001"
                SendEnter()
            End While

            For dd As Integer = 10 To 20 Step 2
                Dim isItLastLine As Boolean = False
                If dd = 10 Then pageNumber += 1
                If Trim(GetText(dd, 1, 47).Replace("-", "")) = "" Then Exit For
                If Trim(GetText(dd + 2, 2, 2)) = "--" Or dd = 20 Then
                    isItLastLine = True
                End If
                lineCount += 1
                If GetText(dd + 1, 6, 2) = "50" Or GetText(dd + 1, 9, 2) = "50" Or GetText(dd + 1, 12, 2) = "50" Or GetText(dd + 1, 15, 2) = "50" Then
                    copiedPos = GetText(dd, 2, 2)
                    copiedRev = GetText(dd, 5, 4)
                    copiedProc = GetText(dd, 10, 6).Replace("-", "")
                    copiedFromDt = GetText(dd, 17, 8)
                    copiedToDt = GetText(dd, 26, 8)
                    copiedUnits = GetText(dd, 35, 3)
                    If CDec(copiedUnits) > 1 Then
                        Dim tempUnit As String = CDec(copiedUnits) / 2
                        copiedUnits = "00" + tempUnit
                    End If
                    copiedCharge = GetText(dd, 40, 5) + "." + GetText(dd, 46, 2)
                    finalCharge = CDec(copiedCharge) / 2
                    Dim chargeArr = (finalCharge).ToString.Split(".")
                    If chargeArr(0).Length = 1 Then
                        copiedFinalChargeWhole = "0000" + chargeArr(0)
                    ElseIf chargeArr(0).Length = 2 Then
                        copiedFinalChargeWhole = "000" + chargeArr(0)
                    ElseIf chargeArr(0).Length = 3 Then
                        copiedFinalChargeWhole = "00" + chargeArr(0)
                    ElseIf chargeArr(0).Length = 4 Then
                        copiedFinalChargeWhole = "0" + chargeArr(0)
                    Else
                        copiedFinalChargeWhole = chargeArr(0)
                    End If
                    'Dim chargeArr = finalCharge.ToString.Split(".")
                    If CountDigits(chargeArr(1)) > 2 Then
                        Dim dec As String = Math.Round(finalCharge, 2).ToString.Split(".")(1)
                        copiedFinalChargeDec = (Convert.ToInt32(dec) + 1).ToString()
                        SendKeyesFor01(copiedFinalChargeWhole, dd, 40)
                        SendKeyesFor01(dec, dd, 46)
                        SendKeyesFor01(copiedUnits, dd, 35)
                    Else
                        copiedFinalChargeDec = chargeArr(1)
                        SendKeyesFor01(copiedFinalChargeWhole, dd, 40)
                        SendKeyesFor01(copiedFinalChargeDec, dd, 46)
                        SendKeyesFor01(copiedUnits, dd, 35)
                    End If

                    lineNumber = CInt(GetText(dd, 78, 3))
                    dt_LinesWith50ModinOCI.Rows.Add(copiedPos, copiedRev, copiedProc, copiedFromDt, copiedToDt, copiedUnits, copiedFinalChargeWhole,
                                                    copiedFinalChargeDec, lineNumber, True, isItLastLine)

                End If

                If dd = 20 Then
                    Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                    SendKeyesFor01("OCN", 2, 2)
                    'SendEnter()
                    SendF8()
                    lineCount = 0
                    If GetText(10, 2, 47, True).Replace("-", "").Trim = "" Then '////////////////  ****Change this line****
                        Exit For
                    ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                        Exit For
                    Else
                        dd = 8
                    End If
                End If
            Next

            checkForBlankLineAndPasteCopiedLineInOCIScreen(dt_LinesWith50ModinOCI)
        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name)
        End Try
    End Sub

    'Public Sub checkForBlankLineAndPasteCopiedLineInOCIScreen(copiedPos As String, copiedRev As String, copiedProc As String, copiedFromDt As String, copiedToDt As String, copiedUnits As String, copiedFinalChargeWhole As String, copiedFinalChargeDec As String)
    Public Sub checkForBlankLineAndPasteCopiedLineInOCIScreen(dtMod50Info As DataTable)
        Try
            For Each row In dtMod50Info.Rows

                For i As Integer = 10 To 20 Step 2
                    If Trim(GetText(i, 1, 47).Replace("-", "")) = "" Then
                        SendKeyesFor01(row("sPOS"), i, 2)
                        SendKeyesFor01(row("sREVCode"), i, 5)
                        SendKeyesFor01(row("sProcCode"), i, 10)
                        SendKeyesFor01(row("sFromDt"), i, 17)
                        SendKeyesFor01(row("sTodt"), 26)
                        SendKeyesFor01(row("sUnits"), i, 35)
                        SendKeyesFor01(row("sChargeWhole"), i, 40)
                        SendKeyesFor01(row("sChargeDec"), i, 46)
                        SendKeyesFor01("50", i + 1, 6)
                        Exit For
                    End If
                    If i = 20 Then
                        Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                        SendF8()
                        If GetText(10, 2, 47, True).Replace("-", "").Trim = "" Then '////////////////  ****Change this line****
                            Exit For
                        ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                            Exit For
                        Else
                            i = 8
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            WriteLog(ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub
#End Region
    '======================================================================================================
#Region "50Modifier operation in EDS"
    Public Sub splitLineOperationFormodifier50InEDS6()
        Dim m50Flag As Boolean = False
        'Dim units As String
        'Dim copiedCharge As Decimal
        'Dim chargeArr As String()
        Dim copiedFinalChargeWhole As String
        Dim copiedFinalChargeDec As String
        Dim lineCount As Integer = 0
        Dim pageNumber As Integer = 0
        Dim SafetyPageCount As Integer = 0
        Try
            WriteLog(">>>>>>>>>>>>>>Entering " + GetCurrentMethod.Name + " <<<<<<<<<<<<<<<<")
            If GetText(2, 2, 2) = "MP" Then
                autECLPSObj.SendKeys("EDS", 2, 2)
                SendEnter()
                apiChk()
            ElseIf GetText(3, 2, 5) = "EDS 6" Then

            Else
                SendClear()
                SendClear()
                pullEDSScreen(CurrICN, 2)
                SendKeyes("A", 1, 76)
                SendEnter()
                apiChk()
                SendKeyes("EDS", 2, 2)
                SendEnter()
            End If
            GenerateDt_Mod50()
            Dim isItLastLine As Boolean = False
            'Dim splitLineNumber As String = GetManualInput("Enter the line number for splitting")
            If GetText(3, 2, 5) = "EDS 6" Then
                WriteLog(GetTextRect())
                For dd As Integer = 6 To 18 Step 3
                    If dd = 6 Then pageNumber += 1
                    If Trim(GetText(dd, 18, 6)) = "" Then
                        Exit For
                    End If
                    If Trim(GetText(dd + 3, 2, 3)) <> "---" Or dd = 18 Then
                        isItLastLine = True
                    End If
                    autECLPSObj.SendKeys("opn", dd, 2)
                    SendEnter()
                    apiChk()
                    WriteLog(GetTextRect())
                    For hh = 8 To 17 Step 3
                        If Trim(GetText(hh, 7, 4)) = "" Then Exit For
                        lineCount += 1
                        If GetText(hh, 20, 2) = "50" Or GetText(hh, 23, 2) = "50" Or GetText(hh, 25, 2) = "50" Or GetText(hh, 29, 2) = "50" Then
                            dt_LinesWith50Modifier.Rows.Add(dd.ToString, hh.ToString, GetText(hh, 7, 4), GetText(hh, 13, 5), GetText(hh, 47, 6) + GetText(hh, 54, 2), lineCount.ToString(), pageNumber.ToString(), True, isItLastLine)
                        End If
                        If hh = 17 Then
                            Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                            SendF8()
                            If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                WriteLog(GetTextRect())
                                SendF3()
                                Exit For
                            ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                                SendF3()
                                Exit For
                            Else
                                hh = 5
                            End If
                        End If
                    Next

                    'SendKeyes("Y", 2, 49)
                    'SendEnter()
                    SendF3()
                    If m50Flag = True Then
                        'Dim lineNum = GetManualInput("the line number which should be carved out?")
                        Dim lineNum = getLineNumForCarving(lineCount)
                        'carvingOutSplitLine(lineNum, dd,)
                        'placing split charges 
                        SendKeyesFor01(copiedFinalChargeWhole, dd, 49)
                        SendKeyesFor01(copiedFinalChargeDec, dd, 56)

                        SendKeyesFor01(copiedFinalChargeWhole, dd + 3, 49)
                        SendKeyesFor01(copiedFinalChargeDec, dd + 3, 56)
                        SendEnter()
                        m50Flag = False
                        dd = dd + 3
                    End If

                    If dd = 18 Then
                        SendEnter()
                        If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                            SendKeyes("N", 4, 79)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                            Exit For
                        Else
                            SafetyPageCount = SafetyPageCount + 1

                            ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                            If SafetyPageCount > 50 Then
                                WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                Exit For
                            End If

                            dd = 3
                        End If
                    End If
                Next

            End If

            'Looping on all rows of datatable dt_LinesWith50Modifier
            For Each row As DataRow In dt_LinesWith50Modifier.Rows
                If dt_LinesWith50Modifier.Rows.Count > 1 Then
                    strResearchComments = "more than 1 lines with 50 modifier"
                    Exit Sub
                End If
                Dim eds6Row As String = row(0)
                Dim eds6point5Row As String = row(1)
                Dim eds6PageNum As Integer = CInt(row(6))
                goToEDS6Page(eds6PageNum)
                If row.Table.Rows.IndexOf(row) = 0 Then
                    SendKeyes("opn", eds6Row, 2)
                    SendEnter()
                    apiChk()
                Else
                    FindMatchingLineForSplitting(row)
                End If

                For jj = 8 To 17 Step 3
                    If Trim(GetText(jj, 7, 4)) = "" Then Exit For
                    If row(2) = GetText(jj, 7, 4) And row(3) = GetText(jj, 13, 5) And (row(4) = GetText(jj, 47, 6) + GetText(jj, 54, 2)) Then
                        SplitAndRepeatLine(eds6point5Row)
                        Dim lineNum = getLineNumForCarving(row(5))
                        carvingOutSplitLine(lineNum, eds6Row, row)

                        'placing split charges 
                        'SendKeyesFor01(copiedFinalChargeWhole, eds6Row, 49)
                        'SendKeyesFor01(copiedFinalChargeDec, eds6Row, 56)

                        'SendKeyesFor01(copiedFinalChargeWhole, eds6Row + 3, 49)
                        'SendKeyesFor01(copiedFinalChargeDec, eds6Row + 3, 56)
                        SendEnter()
                    End If

                    If jj = 17 Then
                        Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                            SendF3()
                            Exit For
                        ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                            SendF3()
                            Exit For
                        Else
                            jj = 5
                        End If
                    End If
                Next
            Next

            '====================Going to MPC so that everything gets saved===================================
            If GetText(4, 15, 10) = "9999999999" Or GetText(4, 15, 10) = "AA12345678" Or GetText(4, 15, 10) = "----------" Then
                SendKeyesFor01(OriginalDraft, 4, 32)
                SendKeyesFor01(OriginalICN, 4, 15)
                SendEnter()
                apiChk()
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
                If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                    SendKeyes("N", 4, 79)
                    SendEnter()
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    apiChk()
                End If
            Else
                'A over to MPC
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
                If InStr(GetText(24, 1, 80), "ENTER HIST INDICATOR") Then
                    SendKeyes("N", 4, 79)
                    SendEnter()
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    apiChk()
                End If
            End If

            'Check edit present: "E2570 RSC APLY JK"
            Dim EditWarning As String = ""
            Dim Count As Integer = 0

            Do
                If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                If Count > 50 Then Exit Do
                EditWarning += GetText(24, 1, 80, True)
                SendEnter()
            Loop
            If InStr(EditWarning, "E2570") > 0 Then
                SendOverrideInScreen("18")
            End If

        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub

    Public Sub carvingOutSplitLine(lineNum As String, currentlineNumber As String, row As DataRow)
        Dim actualLine As String = ""
        Try
            Dim SafetyPageCount As Integer = 0
            WriteLog(">>>>>>>>>>>>>>Entering " + GetCurrentMethod.Name + " <<<<<<<<<<<<<<<<")
            'Carving process
            actualLine = lineNum
            While GetText(currentlineNumber + 2, 4, 2) <> lineNum.Substring(1) And GetText(currentlineNumber + 2, 30, 2) <> lineNum.Substring(1)

                SendF6()

            End While
            SendKeyesFor01(actualLine, currentlineNumber, 2)
            WriteLog(GetTextRect())
            SendEnter()
            apiChk()
            If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                SendKeyes("N", 4, 79)
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
            End If
            SendKeyesFor01("00000000", currentlineNumber, 49)
            WriteLog(GetTextRect())
            autECLPSObj.SendKeys("[ENTER]")
            apiChk()
            'While InStr(GetText(24, 1, 80), "NBAL") > 0
            'End While
            Dim amtArr As String()
            If InStr(GetText(24, 1, 80), "NBAL") > 0 Then
                amtArr = getNBALamount()
                WriteLog("amtArr(0) =" + amtArr(0) + " amtArr(1) =" + amtArr(1))
            End If
            SendKeyesFor01(amtArr(0) + amtArr(1), currentlineNumber, 49)
            WriteLog(GetTextRect())
            autECLPSObj.SendKeys("[ENTER]")

            Dim amtArr2 As String()
            If InStr(GetText(24, 1, 80), "NBAL") > 0 Then
                amtArr2 = getNBALamount()
            End If

            If CurrentEDS6Point5Screen.Length > 1 Then ' To check whether MORE THAN 1 LINES ARE PRESENT IN eds6, ONLY THEN THE CARVED OUT LINE IS REQUIRED TO MOVE
                If row(8) = False Then  '****************Checking from datatable row whether the line is lastLine, if it is lastline then no need to move
                    ''''''''Pushing carved line next to original line
                    SendKeyesFor01("a", currentlineNumber, 2)
                    autECLPSObj.SendKeys("[ENTER]")
                End If
            End If

            For j As Integer = currentlineNumber To 18 Step 3
                If Trim(GetText(j + 3, 2, 80)) = "" And j <> 18 Then
                    'Dim amtArr2 As String()
                    'If InStr(GetText(24, 1, 80), "NBAL") > 0 Then
                    '    amtArr2 = getNBALamount()
                    'End If
                    SendKeyesFor01(amtArr2(0) + amtArr2(1), j, 49)
                    autECLPSObj.SendKeys("[ENTER]")

                    If CurrentEDS6Point5Screen.Length > 1 Then  ' To check whether MORE THAN 1 LINES ARE PRESENT IN eds6, ONLY THEN THE CARVED OUT LINE IS REQUIRED TO MOVE
                        If row(8) = False Then
                            SendKeyesFor01("m", j, 2)
                            WriteLog(GetTextRect())
                            SendEnter()
                        End If
                    End If

                    Exit Sub
                End If
                If j = 18 Then  'to check if 18th line is the last line
                    'SendF8()
                    SendEnter()
                    If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                        SendKeyes("N", 4, 79)
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                    Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                    SendF8()
                    If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Or InStr(GetText(24, 1, 80, True), "L5") > 0 Then
                        SendKeyesFor01(amtArr2(0) + amtArr2(1), j, 49)
                        autECLPSObj.SendKeys("[ENTER]")

                        SendKeyesFor01("m", j, 2)
                        SendEnter()
                        Exit Sub
                    ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                        Exit For
                    Else
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                        j = 3
                    End If
                End If
            Next

        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub

    Public Sub SplitAndRepeatLine(eds6point5Row As String)
        Dim m50Flag As Boolean = False
        Dim units As String
        Dim copiedCharge As Decimal
        Dim chargeArr As String()
        Dim copiedFinalChargeWhole As String
        Dim copiedFinalChargeDec As String
        Dim lineCount As Integer = 0
        Try
            WriteLog(">>>>>>>>>>>>>>Entering " + GetCurrentMethod.Name + " <<<<<<<<<<<<<<<<")
            If GetText(eds6point5Row, 20, 2) = "50" Or GetText(eds6point5Row, 23, 2) = "50" Or GetText(eds6point5Row, 25, 2) = "50" Or GetText(eds6point5Row, 29, 2) = "50" Then
                SendKeyesFor01("--------", eds6point5Row, 57)
                SendKeyesFor01("--------", eds6point5Row, 71)
                SendKeyesFor01("r", eds6point5Row, 2)
                WriteLog(GetTextRect)
                SendEnter()
                apiChk()

                '==========================
                units = GetText(eds6point5Row, 38, 5)
                If CDec(units) > 1 Then
                    Dim tempUnit As String = CDec(units) / 2
                    units = "0000" + tempUnit
                End If
                copiedCharge = CDec(GetText(eds6point5Row, 47, 6) + "." + GetText(eds6point5Row, 54, 2)) / 2
                'chargeArr = copiedCharge.ToString.Split(".")
                chargeArr = addZeroesToBillCharge(copiedCharge).ToString.Split(".")
                copiedFinalChargeWhole = chargeArr(0)
                copiedFinalChargeDec = chargeArr(1)
                '===============================================

                SendKeyesFor01(copiedFinalChargeWhole, eds6point5Row, 47)
                SendKeyesFor01(copiedFinalChargeDec, eds6point5Row, 54)
                SendKeyesFor01(units, eds6point5Row, 38)

                SendKeyesFor01(units, eds6point5Row + 3, 38)
                SendKeyesFor01(copiedFinalChargeWhole, eds6point5Row + 3, 47)
                SendKeyesFor01(copiedFinalChargeDec, eds6point5Row + 3, 54)
                SendEnter()
                If InStr(GetText(24, 1, 80), "NBAL") > 0 Then
                    Dim amt = getNBALamount()
                    Dim finalDecPart As String = (Convert.ToInt32(copiedFinalChargeDec) + Convert.ToInt32(amt(1))).ToString
                    SendKeyesFor01(finalDecPart, eds6point5Row + 3, 54)
                    SendEnter()
                End If
                WriteLog(">>>>>>>>>>>>>After spplitting<<<<<<<<<<<<<<<<<" + vbNewLine + GetTextRect())
                SendF3()
                'm50Flag = True
                'Exit For
            End If

        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub
    Public Function getLineNumForCarving(lineCount As Integer) As String
        Dim lineNum As String
        Try
            If lineCount < 10 Then
                lineNum = "m0" + lineCount.ToString
            ElseIf lineCount < 100 And lineCount > 9 Then
                lineNum = "m" + lineCount.ToString
            End If
        Catch ex As Exception
            WriteLog(ex.Message + " from " + GetCurrentMethod.Name)
        End Try
        Return lineNum
    End Function

    Public Function getNBALamount() As String()
        Dim amount As String
        Dim amtArr As String()
        Try
            If InStr(GetText(24, 1, 80), "NBAL") > 0 Then
                Dim edit = GetText(24, 1, 80).Trim
                Dim arr() = edit.Split(" ")

                For Each p In arr
                    If p.Contains("$") Then
                        amount = p.Replace("$", "").ToString()
                        amtArr = amount.Split(".")
                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
        Return amtArr
    End Function

    Public Sub goToEDS6Page(pageNum As Integer)
        Try
            If GetText(3, 2, 5) = "EDS 6" Then
                Do
                    SendF7()
                Loop Until Trim(GetText(24, 1, 80)) = "ALREADY IN FIRST PAGE"
                If pageNum > 1 Then
                    For i As Integer = 1 To pageNum - 1
                        SendF8()
                        apiChk()
                    Next
                End If

            End If
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub

    Public Sub FindMatchingLineForSplitting(dt_row As DataRow)
        Try
            Dim SafetyPageCount As Integer = 0
            If GetText(3, 2, 5) = "EDS 6" Then
                For i As Integer = dt_row(0) To 18 Step 3
                    If GetText(i, 49, 6) + GetText(i, 56, 2) = dt_row(4) Then
                        SendKeyes("opn", i, 2)
                        SendEnter()
                        apiChk()

                        For jj = 8 To 17 Step 3
                            If Trim(GetText(jj, 7, 4)) = "" Then Exit For

                            If GetText(jj, 7, 4) = dt_row(2) And GetText(jj, 13, 5) = dt_row(3) Then

                                Exit Sub
                            End If
                            If jj = 17 Then
                                Dim MyEDS6point5 As String = GetTextRect(8, 1, 17, 79)
                                SendF8()
                                If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                                    SendF3()
                                    Exit For
                                ElseIf GetTextRect(8, 1, 17, 79) = MyEDS6point5 Then
                                    SendF3()
                                    Exit For
                                Else
                                    jj = 5
                                End If
                            End If
                        Next
                    End If



                    If i = 18 Then
                        SendEnter()
                        If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                            SendKeyes("N", 4, 79)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                        SendF8()
                        If InStr(GetText(24, 1, 80, True), "NO MORE RECORDS") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                            Exit For
                        Else
                            SafetyPageCount = SafetyPageCount + 1

                            ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                            If SafetyPageCount > 50 Then
                                WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                Exit For
                            End If
                            i = 3
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub
#End Region

#Region "iCES DENIAL"
    Sub GenerateDt_iCESDenial()         '*********** Addded on 1st jun,24 **********
        dt_iCESCheck.Clear()
        dt_icesMPCScreen.Clear()
        dt_iCESCheck.Columns.Add("REVCode", GetType(String))
        dt_iCESCheck.Columns.Add("CPTCode", GetType(String))
        dt_iCESCheck.Columns.Add("M1", GetType(String))
        dt_iCESCheck.Columns.Add("M2", GetType(String))
        dt_iCESCheck.Columns.Add("Charge", GetType(String))
        dt_iCESCheck.Columns.Add("lineNumber", GetType(Integer))
        dt_iCESCheck.Columns.Add("DOS", GetType(String))
        dt_iCESCheck.Columns.Add("Suffix", GetType(Integer))
        '**************************************************************
        dt_icesMPCScreen.Columns.Add("PS", GetType(String))
        dt_icesMPCScreen.Columns.Add("SVC", GetType(String))
        dt_icesMPCScreen.Columns.Add("InitialRC", GetType(String))
        dt_icesMPCScreen.Columns.Add("DenialRC", GetType(String))
        dt_icesMPCScreen.Columns.Add("Charge", GetType(String))
        dt_icesMPCScreen.Columns.Add("lineNumber", GetType(Integer))
        dt_icesMPCScreen.Columns.Add("linePosition", GetType(Integer))
        dt_icesMPCScreen.Columns.Add("Suffix", GetType(Integer))
    End Sub


    '    Public Sub ICESCheck(icn As String)                                                  'Added 23rd Dec,21; Modified 23rd June,22; Mod completed on 31st Jan,23; Completely changed to datatable on 1st jun,24
    '        Dim editWarning As String = String.Empty
    '        Dim Linecount As Integer = 0
    '        Dim count As Integer
    '        Dim c As Integer = 0
    '        Dim count2 As Integer = 0
    '        Dim denialRC As String = String.Empty
    '        WriteLog("========Entering " + GetCurrentMethod.Name + " ======")
    '        Try
    '            WriteLog("*******************Pulling EDS1 to fetch REVE, HCPC/MD, CHARGE AND SX***********************************")
    '            Dim LCArray(100) As Integer
    '            pullEDSScreen(icn, 1)
    '            WriteLog("======IN EDS 1=========")
    '            WriteLog(GetTextRect())
    '            If InStr(GetText(24, 1, 80), "PF4 TO REVIEW EDC/ICES DECISION") > 0 Then
    '                ICESDenialFlag = True
    '                SendF4()
    '                WriteLog(GetTextRect)
    '                For Y As Integer = 8 To 21 Step 1
    '                    If Trim(GetText(Y, 2, 4)) = "" Then Exit For
    '                    If Trim(GetText(Y, 66, 6)) <> "" Then
    '                        c += 1
    '                    End If
    '                    If Trim(GetText(Y, 19, 8)) <> "" Then
    '                        dt_iCESCheck.Rows.Add(Trim(GetText(Y, 2, 4)), Trim(GetText(Y, 36, 5)), Trim(GetText(Y, 41, 2)), Trim(GetText(Y, 43, 2)), addZeroesToBillCharge(Trim(GetText(Y, 51, 8))), c)
    '                    End If
    '                    If Y = 21 Then
    '                        Dim screen1 As String = GetTextRect(8, 2, 21, 80)
    '                        SendF8()
    '                        If GetTextRect(8, 2, 21, 80) = screen1 Then Exit For
    '                        Y = 7
    '                    End If
    '                Next

    '                SendKeyes("A", 1, 76)
    '                SendEnter()

    '                If GetText(2, 2, 2) = "MP" Then
    '                    'RemovePZ1()
    '                    IncreementMPCCount()
    '                Else
    '                    strResearchComments = "Unable to pull MPC from method " + GetCurrentMethod.Name
    '                    Exit Sub
    '                End If

    '                Dim ifGotColorScheme As Boolean = False
    '                If GetText(2, 2, 2) = "MP" Then
    '                    WriteLog("=====================In MPC======================")
    '                    WriteLog(GetTextRect)
    '                    For I As Integer = 10 To 22 Step 2
    '                        If GetText(I, 1, 2) <> "--" Then
    '                            Linecount += 1
    '                            If GetColorScheme(I, 36) = 3 Then
    '                                ifGotColorScheme = True
    '                                LCArray(count2) = Linecount
    '                                count2 += 1
    '                                dt_icesMPCScreen.Rows.Add(GetText(I, 1, 2), Trim(GetText(I, 4, 5)), GetText(I, 36, 2), "", Convert.ToDecimal(addZeroesToBillCharge(Trim(GetText(I, 41, 5)) + "." + Trim(GetText(I, 47, 2)))), Linecount, I)
    '                            End If
    '                        Else Exit For
    '                        End If
    '                    Next
    '                End If
    '                If ifGotColorScheme = False Then    'Skip claim if no RC is red in MPC screen
    '                    strResearchComments = "Skip ICES:No Denial RC pulled by system"
    '                    Exit Sub
    '                End If
    '                editWarning = String.Empty
    '                Do
    '                    If editWarning = String.Empty Then
    '                        editWarning = String.Empty
    '                        count = 0
    '                    End If
    '                    If InStr(editWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
    '                    If count > 50 Then Exit Do
    '                    editWarning += GetText(24, 1, 80, True)
    '                    autECLPSObj.SendKeys("[ENTER]")
    '                    apiChk(count)
    '                Loop
    '                For Each r As DataRow In dt_icesMPCScreen.Rows
    '                    If r(2) = "" Then
    '                        WriteLog("Skipped from ICES denial; can't capture initialRC")
    '                        strResearchComments = "Skip ICES initialRC"
    '                        Exit Sub
    '                    ElseIf GetText(r(6), 36, 2) <> r(2) And Trim(GetText(r(6), 50, 5)) + "." + GetText(r(6), 56, 2) = 0 Then
    '                        r(3) = GetText(r(6), 36, 2)
    '                        WriteLog("Final RC=" + r(3))
    '                    Else
    '                        WriteLog("Skipped from ICES denial; can't capture finalRC")
    '                        strResearchComments = "Skip ICES finalRC"
    '                        Exit Sub
    '                    End If
    '                    If iCESRcVerify(r(3)) = False Then        '**********Added on 07/03/2024 to verify RC is iCES or not***************
    '                        strResearchComments = "iCES logic applies based on EDS1 FLAG"
    '                        Exit Sub
    '                    End If
    '                Next

    '                'Going to MPC then EDS6.5 to place not covered amount
    '                If GetText(2, 2, 2) = "MP" Then
    '                    SendKeyes("EDS", 2, 2)
    '                    apiChk()
    '                    autECLPSObj.SendKeys("[ENTER]")
    '                    apiChk()
    '                End If

    '                If GetText(3, 2, 5) = "EDS 6" Then
    '                    For Each line As DataRow In dt_iCESCheck.Rows
    '                        icesSuccess = False
    '                        Dim LC As Integer = 0
    '                        Dim chrg As String = String.Empty

    '                        For i As Integer = 6 To 18 Step 3
    '                            If Trim(GetText(i, 8, 2)) = "" Then Exit For
    '                            LC += 1
    '                            If GetText(i, 49, 6).Contains("-") Then
    '                                chrg = (GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2))
    '                            Else
    '                                chrg = (GetText(i, 49, 6) + "." + GetText(i, 56, 2))
    '                            End If
    '                            'If LC <> line(5) Then
    '                            '    GoTo NextRow
    '                            'End If
    '                            If chrg = line(4) And LC = line(5) And (GetText(i, 59, 6) + "." + GetText(i, 66, 2)) <> chrg Then
    '                                SendKeyes("opn", i, 2)
    '                                WriteLog(GetTextRect)
    '                                autECLPSObj.SendKeys("[ENTER]")
    '                                apiChk()

    '                                If GetText(2, 73, 7) = "EDS 6.5" Then
    '                                    For ii As Integer = 8 To 17 Step 3
    '                                        If Trim(GetText(ii, 7, 4)) = "" Then Exit For
    '                                        If line(0) = Trim(GetText(ii, 7, 4)) And line(1) = Trim(GetText(ii, 13, 5)) And line(2) = Trim(GetText(ii, 20, 2)) And line(3) = Trim(GetText(ii, 23, 2)) And Convert.ToDecimal(line(4)) = Convert.ToDecimal(GetText(ii, 47, 9).Replace(" ", ".")) Then
    '                                            Dim rows() As DataRow = dt_icesMPCScreen.Select("lineNumber = " & LC)
    '                                            If rows.Length > 0 Then
    '                                                Dim foundRow As DataRow = rows(0)
    '                                                icesSuccess = True
    '                                                Dim dec As String = Trim(GetText(ii, 54, 2))
    '                                                Dim charg As String = Trim(GetText(ii, 47, 6))
    '                                                SendKeyes(charg, ii, 57)
    '                                                apiChk()
    '                                                SendKeyes(dec, ii, 64)
    '                                                apiChk()
    '                                                SendKeyes(foundRow(3), ii, 67)
    '                                                apiChk()
    '                                                SendKeyes("--------", ii, 71)
    '                                                apiChk()
    '                                                SendEnter()
    '                                                Exit For
    '                                            End If
    '                                        End If
    '                                        If ii = 17 Then
    '                                            Dim screen1 As String = GetTextRect()
    '                                            SendF8()
    '                                            If GetTextRect() = screen1 Then Exit For
    '                                            ii = 5
    '                                        End If
    '                                    Next
    '                                End If
    '                                WriteLog("-----------In EDS 6.5----------------------")
    '                                WriteLog(GetTextRect)
    '                                SendF3()
    '                                WriteLog("-----------In EDS 6----------------------")
    '                                WriteLog(GetTextRect)
    '                                apiChk()
    '                                If icesSuccess = True Then GoTo NextRow
    '                            End If
    '                            If i = 18 Then
    '                                WriteLog(GetTextRect())
    '                                Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
    '                                SendF8()
    '                                If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
    '                                    Exit For
    '                                ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
    '                                    Exit For
    '                                End If
    '                                i = 3
    '                            End If
    '                        Next
    'NextRow:            Next

    '                End If

    '                SendKeyes("A", 1, 75)
    '                SendEnter()
    '                If GetText(2, 2, 2) = "MP" Then
    '                    'RemovePZ1()
    '                    IncreementMPCCount()
    '                End If
    '                If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
    '                    SendKeyes("N", 4, 79)
    '                    autECLPSObj.SendKeys("[ENTER]")
    '                    apiChk()
    '                    SendKeyes("A", 1, 75)
    '                    If GetText(2, 2, 2) = "MP" Then
    '                        'RemovePZ1()
    '                        IncreementMPCCount()
    '                    End If
    '                End If

    '                If GetText(3, 2, 5) = "EDS 6" Then
    '                    SendKeyes("N", 4, 79)
    '                    autECLPSObj.sendkeys("[enter]")
    '                    apiChk()
    '                    SendKeyes("A", 1, 75)
    '                    SendEnter()
    '                    If GetText(2, 2, 2) = "MP" Then
    '                        'RemovePZ1()
    '                        IncreementMPCCount()
    '                    End If
    '                End If

    '            Else
    '                WriteLog("No iCES edit line found")
    '                WriteLog("========EXITING ICESCHECK ======")
    '                'strResearchComments = "RC changed, Exception Processing"
    '                Exit Sub
    '            End If

    '        Catch ex As Exception
    '            WriteLog(GetCurrentMethod.Name + " :" + ex.Message)
    '        Finally
    '            WriteLog("========EXITING ICESCHECK ======")
    '        End Try
    '    End Sub

    Public Sub ICESCheck(icn As String, SuffixCount As Integer)                                                  'Added 23rd Dec,21; Modified 23rd June,22; Mod completed on 31st Jan,23; Completely changed to datatable + using selenium to find denial RC on 11th jun,24
        Dim editWarning As String = String.Empty
        Dim Linecount As Integer = 0
        'Dim count As Integer
        Dim c As Integer = 0
        Dim count2 As Integer = 0
        Dim denialRC As String = String.Empty
        Dim SafetyPageCount As Integer = 0
        Dim sx As Integer = 1
        WriteLog("========Entering " + GetCurrentMethod.Name + " ======")
        Try
            WriteLog("*******************Pulling EDS1 to fetch REVE, HCPC/MD, CHARGE AND SX***********************************")
            Dim LCArray(100) As Integer
            pullEDSScreen(icn, 1)
            WriteLog("======IN EDS 1=========")
            WriteLog(GetTextRect())
            If InStr(GetText(24, 1, 80), "PF4 TO REVIEW EDC/ICES DECISION") > 0 Then
                If SuffixCount > 1 Then
                    'strResearchComments = "ICES denial in multiple suffix"
                    'Exit Sub
                End If
                ICESDenialFlag = True

                SendF4()
                WriteLog(GetTextRect)
                GenerateDt_iCESDenial()
                dt_iCESCheck.Clear()
                For Y As Integer = 8 To 21 Step 1
                    If Trim(GetText(Y, 2, 80)) = "" Then Exit For

                    If Trim(GetText(Y, 2, 4)) <> "" Then
                        If Trim(GetText(Y, 63, 2)) <> "" And CDec(GetText(Y, 60, 2).Trim) > sx Then
                            c = 0
                            sx += 1
                        End If
                    End If
                    If Trim(GetText(Y, 66, 6)) <> "" Then
                        c += 1
                    End If
                    If Trim(GetText(Y, 19, 8)) <> "" Then
                        dt_iCESCheck.Rows.Add(Trim(GetText(Y, 2, 4)), Trim(GetText(Y, 36, 5)), Trim(GetText(Y, 41, 2)), Trim(GetText(Y, 43, 2)), addZeroesToBillCharge(Trim(GetText(Y, 51, 8))), c, GetText(Y, 46, 4), GetText(Y, 60, 2))
                    End If
                    If Y = 21 Then
                        Dim screen1 As String = GetTextRect(8, 2, 21, 80)
                        SendF8()
                        If GetTextRect(8, 2, 21, 80) = screen1 Then Exit For
                        Y = 7
                    End If
                Next

                SendKeyes("A", 1, 76)
                SendEnter()

                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                Else
                    strResearchComments = "Unable to pull MPC from method " + GetCurrentMethod.Name
                    Exit Sub
                End If

                Dim totalCharge As Decimal = dt_iCESCheck.AsEnumerable().
                    Sum(Function(row) Convert.ToDecimal(row.Field(Of String)("Charge")))

                '*******************************************************************************************************************************************

                Dim cls1 As Class1 = New Class1()
                cls1.generateDtJOCTable()
                cls1.generateDtForClaimRC()
                dtClaimInfo = cls1.getEchoValue(icn)

                If InStr(strResearchComments, "No matching row found in JOC") > 0 Then
                    Exit Sub
                End If
                If dtClaimInfo Is Nothing Then Exit Sub
                '*********************************************************************************************************************************************

                'Joining datatables
                Dim dtCombinedICESInfo As New DataTable()
                dtCombinedICESInfo.Columns.Add("sREVCode", GetType(String))
                dtCombinedICESInfo.Columns.Add("sCPTCode", GetType(String))
                dtCombinedICESInfo.Columns.Add("sCharge", GetType(String))
                dtCombinedICESInfo.Columns.Add("iLineNumber", GetType(Integer))
                dtCombinedICESInfo.Columns.Add("sM1", GetType(String))
                dtCombinedICESInfo.Columns.Add("sM2", GetType(String))
                dtCombinedICESInfo.Columns.Add("sRemarkCode", GetType(String))
                dtCombinedICESInfo.Columns.Add("sDOS", GetType(String))

                'dtCombinedICESInfo.Clear()
                Dim totalChargeFromJOC As Decimal = dtClaimInfo.AsEnumerable().Sum(Function(row) Convert.ToDecimal(row.Field(Of String)("dChargedBilledAmount")))
                '(FOR A SINGLE LINE PRESENT IN JOC INSTEAD OF MULTIPLE LINES LIKE EDS1)
                If totalCharge = totalChargeFromJOC Then
                    Dim denialSuccess = PerformIcesDenialForSingleTotalLineinJOC(dtCombinedICESInfo, SuffixCount)
                    If denialSuccess Then
                        Exit Sub
                    End If
                End If

                For Each row2 As DataRow In dt_iCESCheck.Rows
                    For Each row1 As DataRow In dtClaimInfo.Rows
                        Dim parsedDate As Date = Date.Parse(row1("FirstDateEcho"))
                        Dim resultDate As String = parsedDate.ToString("MMdd")

                        If row1("sRevCode").ToString() = row2("REVCode").ToString AndAlso row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso
                            CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) And row2("DOS").ToString = resultDate Then
                            'row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) Then

                            dtCombinedICESInfo.Rows.Add(row2("REVCode"), row1("sCPTCode"), row2("Charge"), row2("LineNumber"), row2("M1"), row2("M2"), row1("sRemarkCode"))
                        End If
                    Next
                Next
                icesSuccess = False

                For sfx As Integer = 1 To SuffixCount

                    SendClear()
                    SendClear()
                    pullEDSScreen(CurrentClaimDetails.ICN, 1)
                    SendKeyes("A" & sfx, 1, 76)
                    SendEnter()
                    apiChk()

                    'Going to MPC then EDS6.5 to place not covered amount
                    If GetText(2, 2, 2) = "MP" Then
                        SendKeyes("EDS", 2, 2)
                        apiChk()
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                    End If
                    If GetText(3, 2, 5) = "EDS 6" Then
                        For Each line As DataRow In dtCombinedICESInfo.Rows
                            Dim LC As Integer = 0
                            Dim chrg As String = String.Empty

                            For i As Integer = 6 To 18 Step 3
                                If Trim(GetText(i, 8, 2)) = "" Then Exit For
                                LC += 1
                                If GetText(i, 49, 6).Contains("-") Then
                                    chrg = (GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2).Replace("-", "0"))
                                Else
                                    chrg = (GetText(i, 49, 6) + "." + GetText(i, 56, 2))
                                End If
                                'If LC <> line(5) Then
                                '    GoTo NextRow
                                'End If
                                'If chrg = line(2) And LC = line(3) And (GetText(i, 59, 6) + "." + GetText(i, 66, 2)) <> chrg Then
                                If (LC = line(3)) Or CDbl(GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2).Replace("-", "0")) = line(2) Then
                                    SendKeyes("opn", i, 2)
                                    WriteLog(GetTextRect)
                                    autECLPSObj.SendKeys("[ENTER]")
                                    apiChk()

                                    If GetText(2, 73, 7) = "EDS 6.5" Then
                                        For ii As Integer = 8 To 17 Step 3
                                            If Trim(GetText(ii, 7, 4)) = "" Then Exit For
                                            If line(0) = Trim(GetText(ii, 7, 4)) And line(1) = Trim(GetText(ii, 13, 5)).Replace("-", "") And
                                            line(4) = Trim(GetText(ii, 20, 2)) And line(5) = Trim(GetText(ii, 23, 2)) And
                                            Convert.ToDecimal(line(2)) = Convert.ToDecimal(GetText(ii, 47, 9).Replace(" ", ".")) Then

                                                icesSuccess = True
                                                Dim dec As String = Trim(GetText(ii, 54, 2))
                                                Dim charg As String = Trim(GetText(ii, 47, 6))
                                                SendKeyes(charg, ii, 57)
                                                apiChk()
                                                SendKeyes(dec, ii, 64)
                                                apiChk()
                                                SendKeyes(line(6), ii, 67)
                                                apiChk()
                                                SendKeyes("--------", ii, 71)
                                                apiChk()
                                                SendKeyes("Y", 2, 49)
                                                SendEnter()
                                                Exit For
                                            End If
                                            If ii = 17 Then
                                                Dim screen1 As String = GetTextRect()
                                                SendF8()
                                                If GetTextRect() = screen1 Then Exit For
                                                ii = 5
                                            End If
                                        Next
                                    End If
                                    WriteLog("-----------In EDS 6.5----------------------")
                                    WriteLog(GetTextRect)
                                    SendF3()
                                    WriteLog("-----------In EDS 6----------------------")
                                    WriteLog(GetTextRect)
                                    apiChk()
                                    If icesSuccess = True Then GoTo NextRow
                                End If
                                If i = 18 Then
                                    WriteLog(GetTextRect())
                                    Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                                    SendF8()
                                    If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                        Exit For
                                    ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                        Exit For
                                    Else
                                        SafetyPageCount = SafetyPageCount + 1

                                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                        If SafetyPageCount > 50 Then
                                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                            Exit For
                                        End If
                                    End If
                                    i = 3
                                End If
                            Next
NextRow:                Next

                    End If

                    SendKeyes("A", 1, 75)
                    SendEnter()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                    If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                        SendKeyes("N", 4, 79)
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()
                        SendKeyes("A", 1, 75)
                        If GetText(2, 2, 2) = "MP" Then
                            'RemovePZ1()
                            IncreementMPCCount()
                        End If
                    End If

                    If GetText(3, 2, 5) = "EDS 6" Then
                        SendKeyes("N", 4, 79)
                        autECLPSObj.sendkeys("[enter]")
                        apiChk()
                        SendKeyes("A", 1, 75)
                        SendEnter()
                        If GetText(2, 2, 2) = "MP" Then
                            'RemovePZ1()
                            IncreementMPCCount()
                        End If
                    End If
                Next
            Else
                WriteLog("No iCES edit line found")
                SendKeyes("A", 1, 76)
                SendEnter()
                WriteLog("========EXITING ICESCHECK ======")
                'strResearchComments = "RC changed, Exception Processing"
                Exit Sub
            End If

        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " :" + ex.Message)
        Finally
            WriteLog("========EXITING ICESCHECK ======")
        End Try
    End Sub

    Public Function PerformIcesDenialForSingleTotalLineinJOC(byref dtCombinedICESInfo As DataTable, SuffixCount As Integer) As Boolean
        Dim SafetyPageCount As Integer = 0
        Dim denialPerformed = False
        Try
            dtCombinedICESInfo.Clear()
            dtCombinedICESInfo.Columns.Add("iSuffix", GetType(Integer))

            For Each row2 As DataRow In dt_iCESCheck.Rows
                For Each row1 As DataRow In dtClaimInfo.Rows
                    Dim parsedDate As Date = Date.Parse(row1("FirstDateEcho"))
                    Dim resultDate As String = parsedDate.ToString("MMdd")

                    If row2("DOS").ToString = resultDate Then
                        'row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) Then

                        dtCombinedICESInfo.Rows.Add(row2("REVCode"), row2("CPTCode"), row2("Charge"), row2("LineNumber"), row2("M1"), row2("M2"), dtClaimInfo(0)("sRemarkCode"), resultDate, row2("Suffix"))
                    End If
                Next
            Next
            icesSuccess = False

            For sfx As Integer = 1 To SuffixCount

                SendClear()
                SendClear()
                pullEDSScreen(CurrentClaimDetails.ICN, 1)
                SendKeyes("A" & sfx, 1, 76)
                SendEnter()
                apiChk()

                ' Filter rows for this suffix
                Dim combinedRows As DataTable = dtCombinedICESInfo.Select($"iSuffix = {sfx}").CopyToDataTable()
                'Dim newtable As DataTable = dtCombinedICESInfo.Select("iSuffix = '" & sfx & "'").CopyToDataTable()

                ' If there are no rows, skip
                If combinedRows Is Nothing OrElse combinedRows.Rows.Count = 0 Then
                    Continue For
                End If

                'Going to MPC then EDS6.5 to place not covered amount
                If GetText(2, 2, 2) = "MP" Then
                    SendKeyes("EDS", 2, 2)
                    apiChk()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                End If
                If GetText(3, 2, 5) = "EDS 6" Then
                    For Each line As DataRow In combinedRows.Rows
                        Dim LC As Integer = 0
                        Dim chrg As String = String.Empty

                        For i As Integer = 6 To 18 Step 3
                            If Trim(GetText(i, 8, 2)) = "" Then Exit For
                            LC += 1
                            If GetText(i, 49, 6).Contains("-") Then
                                chrg = (GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2)).Replace("-", "0")
                            Else
                                chrg = (GetText(i, 49, 6) + "." + GetText(i, 56, 2))
                            End If
                            'If LC <> line(5) Then
                            '    GoTo NextRow
                            'End If
                            'If chrg = line(2) And LC = line(3) And (GetText(i, 59, 6) + "." + GetText(i, 66, 2)) <> chrg Then
                            If (LC = CInt(line(3))) Then
                                SendKeyes("opn", i, 2)
                                WriteLog(GetTextRect)
                                autECLPSObj.SendKeys("[ENTER]")
                                apiChk()

                                If GetText(2, 73, 7) = "EDS 6.5" Then
                                    For ii As Integer = 8 To 17 Step 3
                                        If Trim(GetText(ii, 7, 80)) = "" Then Exit For
                                        If line(0) = Trim(GetText(ii, 7, 4)) And line(1) = Trim(GetText(ii, 13, 5)).Replace("-", "") And
                                        Convert.ToDecimal(line(2)) = Convert.ToDecimal(GetText(ii, 47, 9).Replace(" ", ".")) Then

                                            icesSuccess = True
                                            Dim dec As String = Trim(GetText(ii, 54, 2))
                                            Dim charg As String = Trim(GetText(ii, 47, 6))
                                            SendKeyes(charg, ii, 57)
                                            apiChk()
                                            SendKeyes(dec, ii, 64)
                                            apiChk()
                                            SendKeyes(line(6), ii, 67)
                                            apiChk()
                                            SendKeyes("--------", ii, 71)
                                            apiChk()
                                            SendKeyes("Y", 2, 49)
                                            SendEnter()
                                            Exit For
                                        End If
                                        If ii = 17 Then
                                            Dim screen1 As String = GetTextRect()
                                            SendF8()
                                            If GetTextRect() = screen1 Then Exit For
                                            ii = 5
                                        End If
                                    Next
                                End If
                                WriteLog("-----------In EDS 6.5----------------------")
                                WriteLog(GetTextRect)
                                SendF3()
                                WriteLog("-----------In EDS 6----------------------")
                                WriteLog(GetTextRect)
                                apiChk()
                                If icesSuccess = True Then GoTo NextRow
                            End If
                            If i = 18 Then
                                WriteLog(GetTextRect())
                                Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                                SendF8()
                                If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                    Exit For
                                ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                    Exit For
                                Else
                                    SafetyPageCount = SafetyPageCount + 1

                                    ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                    If SafetyPageCount > 50 Then
                                        WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                        Exit For
                                    End If
                                End If
                                i = 3
                            End If
                        Next
NextRow:            Next

                End If

                SendKeyes("A", 1, 75)
                SendEnter()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                End If

                If GetText(3, 2, 5) = "EDS 6" Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.sendkeys("[enter]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                End If
            Next
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)

        End Try
        Return icesSuccess
    End Function
#End Region

#Region "Deny with RC = 7G for not timely filed claims"
    Public Sub denyWith7G()             'Added 24/08/2024
        Try
            Dim SafetyPageCount As Integer = 0
            WriteLog(">>>>> Entering " + GetCurrentMethod.Name + " <<<<<<")
            If GetText(3, 2, 5) = "EDS 6" Then
                For i As Integer = 6 To 18 Step 3

                    If Trim(GetText(i, 8, 2)) <> "" Then

                        If (GetText(i, 11, 6, True) = "OI") OrElse (GetText(i, 11, 6, True) = "OIM") OrElse (GetText(i, 11, 6, True) = "BLANK-") OrElse
                       (InStr(PISLLINE, GetText(i, 11, 6, True)) > 0 AndAlso InStr("OI,OIM", GetText(i, 11, 6, True)) = 0) Then
                            Continue For
                        Else
                            If GetText(i, 49, 6) <> GetText(i, 59, 6) AndAlso GetText(i, 56, 2) <> GetText(i, 66, 2) Then
                            End If
                            SendKeyes("opn", i, 2)
                            SendEnter()
                            apiChk()
                            If GetText(2, 73, 7) = "EDS 6.5" Then
                                For ii As Integer = 8 To 17 Step 3
                                    If Trim(GetText(ii, 7, 4)) <> "" Then

                                        Dim dec As String = Trim(GetText(ii, 54, 2))
                                        Dim charg As String = Trim(GetText(ii, 47, 6))
                                        SendKeyes(charg, ii, 57)
                                        apiChk()
                                        SendKeyes(dec, ii, 64)
                                        apiChk()
                                        SendKeyes("T8", ii, 67)
                                        apiChk()
                                        SendKeyes("--------", ii, 71)
                                        apiChk()
                                        'Exit For
                                        WriteLog(GetTextRect)
                                    Else
                                        SendF3()
                                        Exit For
                                    End If
                                    If ii = 17 Then
                                        Dim screen1 As String = GetTextRect()
                                        SendF8()
                                        If GetTextRect() = screen1 Then
                                            SendF3()
                                            Exit For
                                        Else
                                            ii = 5
                                        End If
                                    End If
                                Next
                                'autECLPSObj.sendkeys("[enter]")
                                apiChk()
                            End If

                        End If
                    Else
                        Exit For
                    End If
                    If i = 18 Then
                        WriteLog(GetTextRect())
                        Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                        SendF8()
                        If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                            Exit For
                        ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                            Exit For
                        End If
                        SafetyPageCount = SafetyPageCount + 1

                        ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                        If SafetyPageCount > 50 Then
                            WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                            Exit For
                        End If
                        i = 3
                    End If
                Next
            End If
            WriteLog("-----------In EDS 6----------------------")
            WriteLog(GetTextRect)
            apiChk()

            SendKeyes("A", 1, 75)
            SendEnter()
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub
#End Region


    Public Function DetermineHxFilingDate(matchICNList As String) As String
        Dim hxdate As String = ""
        Try
            PullAHI()
            goToAHILastPage()
            For dd As Integer = 20 To 4 Step -2
                If Trim(GetText(dd, 4, 6)) = "------" Then Continue For
                If InStr(matchICNList, GetText(dd + 1, 54, 10)) > 0 And (Trim(GetText(dd, 73, 2)) <> "C" Or Trim(GetText(dd, 73, 2)) <> "CN") Then
                    hxdate = GetText(dd + 1, 73, 6)
                    Exit For
                End If
                If dd = 4 Then

                    SendF7()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    dd = 22
                End If
            Next
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
        Return hxdate
    End Function

    Public Sub goToAHILastPage()
        Try
            Dim count = 0
            If GetText(2, 2, 2) = "AH" Then
                Do
                    count += 1
                    SendF8()
                    If count > 10 Then Exit Do
                Loop Until instr(Trim(GetText(24, 1, 80)), "NO MORE RECORDS") > 0


            End If
        Catch ex As Exception
            WriteLog("Exception " + ex.Message + " from " + GetCurrentMethod.Name)
        End Try
    End Sub
    Public Function denyInOCI(dt_762Denial As DataTable) As Boolean
        Dim flag As Boolean = False
        Try


            'Iterate on each row of datatable
            For Each row As DataRow In dt_762Denial.Rows
                'Go to first page of OCI
                SendKeyesFor01("OCB", 2, 2)
                While GetText(10, 78, 3) <> "001"
                    SendEnter()
                End While

                For dd As Integer = 10 To 20 Step 2     'loop to match lines from OCI screen
                    If Trim(GetText(dd, 1, 56).Replace("-", "")) = "" Then Exit For
                    If GetText(dd, 5, 4) = "0762" And CInt(Trim(GetText(dd, 40, 5))) = CInt(row("BillCharge").split(" ")(0)) And Trim(GetText(dd, 46, 2)) = row("BillCharge").split(" ")(1) And GetText(dd, 10, 5) = row("CPTCode") Then
                        SendKeyes(GetText(dd, 40, 5), dd, 49)
                        SendKeyes(GetText(dd, 46, 2), dd, 55)
                    End If

                    If dd = 20 Then
                        Dim Myscreen6 As String = GetTextRect(5, 1, 19, 79)
                        SendKeyesFor01("OCN", 2, 2)
                        'SendEnter()
                        SendF8()
                        If GetText(10, 2, 47, True).Replace("-", "").Trim = "" Then '////////////////  ****Change this line****
                            Exit For
                        ElseIf GetTextRect(5, 1, 19, 79) = Myscreen6 Then
                            Exit For
                        Else
                            dd = 8
                        End If
                    End If
                Next
            Next

        Catch ex As Exception

        End Try
        Return flag
    End Function
#Region "Second call for iCES denial"
    Public Sub placingiCESDenialInfo(icn As String, SuffixCount As Integer)                                    'Added 21/10/2024 to use the existing datatable and place the required iCES info in EDS screen
        Try
            Dim SafetyPageCount As Integer = 0
            WriteLog("Entering  " + GetCurrentMethod.Name())
            If GetText(2, 2, 2) <> "MP" Then
                pullEDSScreen(icn, 1)
                SendKeyes("A", 1, 76)
                SendEnter()
            End If
            'Going to MPC then EDS6.5 to place not covered amount
            If GetText(2, 2, 2) = "MP" Then
                SendKeyes("EDS", 2, 2)
                apiChk()
                autECLPSObj.SendKeys("[ENTER]")
                apiChk()
            End If

            'Joining datatables
            Dim dtCombinedICESInfo As New DataTable()
            dtCombinedICESInfo.Columns.Add("sREVCode", GetType(String))
            dtCombinedICESInfo.Columns.Add("sCPTCode", GetType(String))
            dtCombinedICESInfo.Columns.Add("sCharge", GetType(String))
            dtCombinedICESInfo.Columns.Add("iLineNumber", GetType(Integer))
            dtCombinedICESInfo.Columns.Add("sM1", GetType(String))
            dtCombinedICESInfo.Columns.Add("sM2", GetType(String))
            dtCombinedICESInfo.Columns.Add("sRemarkCode", GetType(String))

            'For Each row2 As DataRow In dt_iCESCheck.Rows
            '    For Each row1 As DataRow In dtClaimInfo.Rows
            '        If row1("sRevCode").ToString() = row2("REVCode").ToString AndAlso
            '            row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) Then

            '            dtCombinedICESInfo.Rows.Add(row1("sRevCode"), row1("sCPTCode"), row2("Charge"), row2("LineNumber"), row2("M1"), row2("M2"), row1("sRemarkCode"))
            '        End If
            '    Next
            'Next
            For Each row2 As DataRow In dt_iCESCheck.Rows
                For Each row1 As DataRow In dtClaimInfo.Rows
                    Dim parsedDate As Date = Date.Parse(row1("FirstDateEcho"))
                    Dim resultDate As String = parsedDate.ToString("MMdd")

                    If row1("sRevCode").ToString() = row2("REVCode").ToString AndAlso row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso
                            CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) And row2("DOS").ToString = resultDate Then
                        'row1("sCPTCode").ToString() = row2("CPTCode").ToString().Replace("-", "") AndAlso CDec(row1("dChargedBilledAmount")) = CDec(row2("Charge")) Then

                        dtCombinedICESInfo.Rows.Add(row1("sRevCode"), row1("sCPTCode"), row2("Charge"), row2("LineNumber"), row2("M1"), row2("M2"), row1("sRemarkCode"))
                    End If
                Next
            Next
            icesSuccess = False


            For sfx As Integer = 1 To SuffixCount

                SendClear()
                SendClear()
                pullEDSScreen(CurrentClaimDetails.ICN, 1)
                SendKeyes("A" & sfx, 1, 76)
                SendEnter()
                apiChk()

                'Going to MPC then EDS6.5 to place not covered amount
                If GetText(2, 2, 2) = "MP" Then
                    SendKeyes("EDS", 2, 2)
                    apiChk()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                End If
                Dim totalcharge = dtCombinedICESInfo.AsEnumerable().Sum(Function(row) Convert.ToDecimal(row.Field(Of String)("sCharge")))

                If GetText(3, 2, 5) = "EDS 6" Then
                    For Each line As DataRow In dtCombinedICESInfo.Rows
                        icesSuccess = False
                        Dim LC As Integer = 0
                        Dim chrg As String = String.Empty
                        For i As Integer = 6 To 18 Step 3
                            If Trim(GetText(i, 8, 2)) = "" Then Exit For
                            LC += 1
                            If GetText(i, 49, 6).Contains("-") Then
                                chrg = (GetText(i, 49, 6).Replace("-", "0") + "." + GetText(i, 56, 2).Replace("-", "0"))
                            Else
                                chrg = (GetText(i, 49, 6) + "." + GetText(i, 56, 2))
                            End If
                            'If LC <> line(5) Then
                            '    GoTo NextRow
                            'End If
                            'If chrg = line(2) And LC = line(3) And (GetText(i, 59, 6) + "." + GetText(i, 66, 2)) <> chrg Then
                            If (chrg = line(2) Or CDec(chrg) = totalcharge) Or LC = line(3) Then
                                SendKeyes("opn", i, 2)
                                WriteLog(GetTextRect)
                                autECLPSObj.SendKeys("[ENTER]")
                                apiChk()

                                If GetText(2, 73, 7) = "EDS 6.5" Then
                                    For ii As Integer = 8 To 17 Step 3
                                        If Trim(GetText(ii, 7, 4)) = "" Then Exit For
                                        If line(0) = Trim(GetText(ii, 7, 4)) And line(1) = Trim(GetText(ii, 13, 5)).Replace("-", "") And
                                        line(4) = Trim(GetText(ii, 20, 2)) And line(5) = Trim(GetText(ii, 23, 2)) And
                                        Convert.ToDecimal(line(2)) = Convert.ToDecimal(GetText(ii, 47, 9).Replace(" ", ".")) And
                                        (GetText(ii, 57, 6).Replace("-", "") <> GetText(ii, 47, 6) And GetText(ii, 54, 2).Replace("-", "") <> GetText(ii, 64, 2)) Then

                                            'Dim rows() As DataRow = dt_icesMPCScreen.Select("lineNumber = " & LC)
                                            'If rows.Length > 0 Then
                                            '    Dim foundRow As DataRow = rows(0)
                                            'End If

                                            icesSuccess = True
                                            Dim dec As String = Trim(GetText(ii, 54, 2))
                                            Dim charg As String = Trim(GetText(ii, 47, 6))
                                            SendKeyes(charg, ii, 57)
                                            apiChk()
                                            SendKeyes(dec, ii, 64)
                                            apiChk()
                                            SendKeyes(line(6), ii, 67)
                                            apiChk()
                                            SendKeyes("--------", ii, 71)
                                            apiChk()
                                            SendKeyes("Y", 2, 49)
                                            SendEnter()
                                            Exit For
                                        End If
                                        If ii = 17 Then
                                            Dim screen1 As String = GetTextRect()
                                            SendF8()
                                            If GetTextRect() = screen1 Then Exit For
                                            ii = 5
                                        End If
                                    Next
                                End If
                                WriteLog("-----------In EDS 6.5----------------------")
                                WriteLog(GetTextRect)
                                SendF3()
                                WriteLog("-----------In EDS 6----------------------")
                                WriteLog(GetTextRect)
                                apiChk()
                                If icesSuccess = True Then GoTo NextRow
                            End If
                            If i = 18 Then
                                WriteLog(GetTextRect())
                                Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                                SendF8()
                                If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                    Exit For
                                ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                    Exit For
                                Else
                                    SafetyPageCount = SafetyPageCount + 1

                                    ' If we have paged down more than 50 times, assume we are stuck and kill the process.
                                    If SafetyPageCount > 50 Then
                                        WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                        Exit For
                                    End If
                                End If
                                i = 3
                            End If
                        Next
NextRow:            Next

                End If

                SendKeyes("A", 1, 75)
                SendEnter()
                If GetText(2, 2, 2) = "MP" Then
                    'RemovePZ1()
                    IncreementMPCCount()
                End If
                If InStr(GetText(24, 2, 80), "PLEASE ENTER HIST INDICATOR") > 0 Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                End If

                If GetText(3, 2, 5) = "EDS 6" Then
                    SendKeyes("N", 4, 79)
                    autECLPSObj.sendkeys("[enter]")
                    apiChk()
                    SendKeyes("A", 1, 75)
                    SendEnter()
                    If GetText(2, 2, 2) = "MP" Then
                        'RemovePZ1()
                        IncreementMPCCount()
                    End If
                End If
            Next

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name())
        End Try
    End Sub
#End Region

    Public Function DetermineHxDateBasedOnClaimClosure() As String

        Dim icn = "EY85693212"
        Dim PAYLOC = "358"
        Dim engine = "L"
        Dim hxDate As String = ""
        Dim I As Integer, J As Integer, MyAHI(100) As AHI
        Try
            GetClaimDetails(CurrentClaimDetails, icn, PAYLOC)
            PullAHI()
            J = ReadAHI(MyAHI)

            For indx = UBound(MyAHI) To LBound(MyAHI)
                If MyAHI(indx).OV <> "CN" And MyAHI(indx).OV <> "C" Then

                End If
            Next

        Catch ex As Exception
            Return hxDate

        End Try
        Return hxDate
    End Function

    Public Function FindClosedICN() As String
        Dim OrICN As String = CurrentClaimDetails.ICN
        Dim icn As String
        Dim AHIICN As String
        icn = String.Empty
        AHIICN = String.Empty
        'Dim HCFC As Boolean
        Dim UB As Boolean
        Dim isMHI As Boolean = True
        Dim isSkip As Boolean           'Added by Sourav to skip the entire claim if HH or IJ is present in HxICN:(19.05.2023)
        Try
            PullAHI()
            If InStr(Trim(GetText(24, 1, 80)), "NO CLAIMS ON FILE") > 0 Then Return "NO CLAIMS ON FILE"
            For dd As Integer = 4 To 20 Step 2
                If Trim(GetText(dd, 4, 6)) = "------" Then Exit For

                If Trim(GetText(dd + 1, 54, 10)) = OrICN Then
                    Continue For
                Else
                    If GetText(dd + 1, 52, 1) = "F" Then
                        Continue For

                    ElseIf GetText(dd + 1, 52, 1) = "I" Then
                        UB = False
                        SendKeyes("S", dd, 2)
                        SendEnter()
                        apiChk()
                        If InStr(Trim(GetText(2, 2, 3)), "GHI") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHN") > 0 OrElse InStr(Trim(GetText(2, 2, 3)), "GHB") > 0 Then
                            isMHI = False
                            SendKeyes("MHI", 2, 2)
                            autECLPSObj.SendKeys("[ENTER]")
                            apiChk()
                        End If
                        If Trim(GetText(2, 2, 2)) = "MH" Then
                            isMHI = True
                        End If
                        If InStr(Trim(GetText(10, 4, 1)), "C") > 0 OrElse InStr(Trim(GetText(10, 7, 3)), "REV") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "MISC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "IC") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "PR") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "SP") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "NB") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "OPS") > 0 OrElse InStr(Trim(GetText(10, 4, 6)), "OA") > 0 Then
                            UB = True
                            If isMHI = False Then
                                PullAHI()
                                isMHI = True
                            Else
                                For ff As Integer = 12 To 20 Step 4
                                    If Trim(GetText(ff - 2, 29, 2)) = "C" Or GetText(ff - 2, 29, 2) = "CN" Then
                                        strResearchComments = "C or CN present in HxICN"
                                        isSkip = True
                                        Return ""
                                    End If
                                    If InStr(Trim(GetText(ff, 1, 3)), "ICN") > 0 And Trim(GetText(ff, 5, 10)) <> "0000000000" Then
                                        AHIICN = Trim(GetText(ff, 5, 10))
                                    End If

                                    If ff = 20 Then
                                        'SendF8()
                                        If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                                        SendF8()
                                        ff = 8
                                    End If
                                Next
                                SendKeyes("A", 3, 15)
                                SendEnter()
                                apiChk()
                                'IncreementMPCCount()

                                If Trim(GetText(2, 2, 2)) <> "AH" Then
                                    PullAHI()
                                End If
                            End If
                        Else
                            'HCFC = True
                        End If

                    End If


                End If
                If UB = True AndAlso Trim(GetText(dd + 1, 54, 10)) <> OrICN Then
                    'AHIICN = Trim(GetText(dd + 1, 54, 10))
                    If InStr(icn, AHIICN) = 0 Then
                        icn = AHIICN + "," + icn
                    End If
                Else
                    SendKeyes("A", 3, 15)
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                    'IncreementMPCCount()
                    If Trim(GetText(2, 2, 2)) <> "AH" Then
                        PullAHI()
                    End If
                End If
                If dd = 20 Then
                    'SendF8()
                    If InStr(autECLPSObj.GetText(24, 1, 80), "W758 NO MORE RECORDS") > 0 Then Exit For
                    SendF8()
                    dd = 2
                End If
            Next
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")

        End Try

        Return icn
    End Function

    Public Function Check30IfLinePasted() As Boolean
        Dim is30LinePresent As Boolean = False
        Try
            If GetText(2, 2, 2) = "MP" Then
                For i As Integer = 10 To 22 Step 2
                    If GetText(i, 4, 6) = "------" Then Exit For
                    If GetText(i, 29, 2) = "30" Then
                        is30LinePresent = True
                    End If
                Next

            End If
        Catch ex As Exception
            WriteLog(GetCurrentMethod.Name + " exception")
        End Try
        Return is30LinePresent
    End Function
    Private Function GetRemarkCodeByRevCode(ByVal revCode As String) As String

        Dim dr() As DataRow = dtNonServiceRemarkCodes.Select(
        "RevenueCode = '" & revCode.Trim() & "'")

        If dr.Length > 0 Then
            Return dr(0)("RemarkCode").ToString()
        End If

        Return String.Empty

    End Function
    Public Sub implicitDenialInEDS(MaxSuffix As Integer)
        WriteLog($">>>>Inside {GetCurrentMethod.Name}<<<<<")
        Try
            Dim SafetyPageCount As Integer = 0
            Dim LC As Integer = 0
            Dim dtRevCode As String = String.Empty
            Dim dtCharge As Decimal = 0.0
            Dim dtRemarkCode As String = ""
            For sfx As Integer = 1 To MaxSuffix
                pullEDSScreen(CurrICN, 1)
                SendKeyes("A" & sfx, 1, 76)
                apiChk()
                autECLPSObj.SendKeys("[ENTER]")
                WriteLog(GetTextRect)
                apiChk()

                If GetText(2, 2, 2) = "MP" Then
                    SendKeyes("EDS", 2, 2)
                    apiChk()
                    autECLPSObj.SendKeys("[ENTER]")
                    apiChk()
                End If

                If GetText(3, 2, 5) = "EDS 6" Then

                    For i As Integer = 6 To 18 Step 3           '/// EDS 6 LOOP 
                        If Trim(GetText(i, 8, 2)) = "" Then Exit For

                        LC += 1
                        SendKeyes("opn", i, 2)
                        WriteLog(GetTextRect)
                        autECLPSObj.SendKeys("[ENTER]")
                        apiChk()

                        If GetText(2, 73, 7) = "EDS 6.5" Then   '/// EDS 6.5 LOOP
                            For ii As Integer = 8 To 17 Step 3
                                If Trim(GetText(ii, 7, 4)) = "" Then Exit For
                                Dim revCode As String = Trim(GetText(ii, 7, 4))
                                Dim remarkCode As String = GetRemarkCodeByRevCode(revCode)
                                If remarkCode <> "" Then

                                    Dim dec As String = Trim(GetText(ii, 54, 2))
                                    Dim charg As String = Trim(GetText(ii, 47, 6))
                                    SendKeyes(charg, ii, 57)
                                    apiChk()
                                    SendKeyes(dec, ii, 64)
                                    apiChk()
                                    SendKeyes(remarkCode, ii, 67)
                                    apiChk()
                                    SendKeyes("--------", ii, 71)
                                    SendKeyes("Y", 2, 49)

                                    WriteLog(GetTextRect)
                                    SendEnter()
                                    apiChk()

                                End If
                                If ii = 17 Then
                                    Dim screen1 As String = GetTextRect()
                                    SendF8()
                                    If GetTextRect() = screen1 Then
                                        SendEnter()
                                        SendF3()
                                        Exit For
                                    End If
                                    ii = 5
                                End If
                            Next
                            'autECLPSObj.sendkeys("[enter]")
                        End If
                        apiChk()
                        SendF3()
                        WriteLog("-----------In EDS 6----------------------")
                        WriteLog(GetTextRect)
                        apiChk()


                        If i = 18 Then
                            WriteLog(GetTextRect())
                            Dim Myscreen6 As String = GetTextRect(5, 1, 18, 79)
                            SendF8()
                            If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 OrElse InStr(autECLPSObj.GetText(24, 1, 80), "ALREADY IN LAST PAGE") > 0 Then
                                Exit For
                            ElseIf GetTextRect(5, 1, 18, 79) = Myscreen6 Then
                                Exit For
                            Else
                                SafetyPageCount = SafetyPageCount + 1

                                ' If we have paged down more than 20 times, assume we are stuck and kill the process.
                                If SafetyPageCount > 20 Then
                                    WriteLog(GetCurrentMethod.Name + " Error: Force exit due to infinite loop detection (Max Pages Exceeded).")
                                    Exit For
                                End If
                            End If
                            i = 3
                        End If
                    Next

                End If

                'WriteLog("-----------In EDS 6.5----------------------")

                'SendF3()
                WriteLog("-----------In EDS 6----------------------")
                WriteLog(GetTextRect)
                apiChk()
                SendEnter()
                SendKeyes("A", 1, 75)
                SendEnter()
                apiChk()
            Next
        Catch ex As Exception

        End Try

    End Sub

#Region "Priority 2 claims protection"
    Public Function ProtectPriorityTwoClaims(ByVal Draft As String, ByVal HxICNList As String) As Boolean
        WriteLog("-------------Entering " + GetCurrentMethod.Name + "-------------")
        Dim isSuccess As Boolean = False
        Dim MyDraft() As String
        Dim AdjConLine As String
        Dim icn As String
        Dim ThirtyLine As String
        Dim serviceLine As New List(Of ServiceLineData)
        Dim serviceLineCount As Integer
        Dim newdata As String = ""
        Dim olddata As String = ""
        Dim serviceLinesForDraft As New Dictionary(Of String, List(Of ServiceLineData))
        Dim draftNumber As String
        Dim thirtyLineAmt As String = ""

        Try
            Dim HICNArray() As String = HxICNList.Split(",")

            For mm As Integer = LBound(HICNArray) To UBound(HICNArray)
                If isSuccess = True Then Exit For
                icn = HICNArray(mm)
                If icn = "" Then Continue For
                If icn.Trim <> "" Then
                    'PullMHI(icn)
                    'IsMHIRead = ReadMHI(HistoryClaim)
                End If

                PullMHI(icn)
                serviceLineCount = 0
                WriteLog("Storing data from MHI")
                For dd As Integer = 10 To 18 Step 4
                    If Trim(GetText(dd, 1, 80)) = "" Then Exit For
                    If draftNumber <> "" And InStr(Trim(GetText(24, 1, 80)), "NO MORE CLMS ON FILE") > 0 And (GetColorScheme(dd + 2, 1) <> "1" And GetColorScheme(dd + 2, 1) <> "2") Then
                        Exit For
                    End If
                    draftNumber = ""

                    If GetText(dd, 36, 2) = "70" And GetText(dd + 2, 1, 3) <> "ICN" And GetText(dd, 4, 6) <> "NCOPAY" Then
                        ThirtyLine = GetTextRect(dd, 1, dd + 1, 80)
                        thirtyLineAmt = Trim(GetText(dd + 1, 30, 8))
                    End If

                    If GetText(dd, 36, 2) <> "69" And GetText(dd, 36, 2) <> "70" And GetText(dd, 36, 2) <> "74" And GetText(dd + 2, 1, 3) <> "ICN" Then
                        serviceLine.Add(BuildFormattedServiceLine(dd, dd + 1))
                        serviceLineCount += 1
                    End If
                    If GetText(dd + 2, 1, 3) = "ICN" Then
                        draftNumber = Trim(GetText(dd, 28, 10))

                    End If
                    ' Store the serviceLine under the draft number
                    If Not String.IsNullOrEmpty(draftNumber) AndAlso serviceLine.Count > 0 Then
                        serviceLinesForDraft(draftNumber) = New List(Of ServiceLineData)(serviceLine)
                        serviceLine.Clear()
                    End If

MHIPageChange:      If dd = 18 Then
                        If InStr(autECLPSObj.GetText(24, 3, 80), "E065NO MORE CLMS ON FILE") > 0 Then Exit For
                        olddata = GetTextRect(10, 1, 20, 80)
                        SendF8()
                        newdata = GetTextRect(10, 1, 20, 80)
                        Dim count As Integer = 0
                        While olddata = newdata
                            SendF8()
                            count += 1
                            If count > 10 Then Exit While
                            newdata = GetTextRect(10, 1, 20, 80)
                            If newdata <> olddata Then Exit While
                        End While
                        dd = 6
                    End If
                Next

                'open Payment screen of history ICN
                If Draft IsNot Nothing Then
                    MyDraft = Draft.Split(",")
                Else
                    MyDraft = New String() {}
                End If
                SendF12()
                For J As Integer = LBound(MyDraft) To UBound(MyDraft)
                    If MyDraft(J).Trim() = "" Then Continue For
                    AdjConLine = "ADJ," + CurrentClaimDetails.Policy + "," + CurrentClaimDetails.SSN + "," +
                                                               CurrentClaimDetails.PatientName + "," + CurrentClaimDetails.PatientRelationship + "," +
                                                               MyDraft(J) + ",I" + icn

                    SendClear()
                    SendClear()
                    SendKeyes(AdjConLine, 2, 2)
                    WriteLog("ADJ line is :" + AdjConLine)
                    SendEnter()
                    WriteLog(GetTextRect)
                    Dim EditWarning1 = GetText(24, 1, 80)
                    If InStr(EditWarning1, "INVALID ICN") > 0 Or InStr(EditWarning1, "DRAFT NOT ON FILE") > 0 Or InStr(EditWarning1, "E124INVALID CODE") > 0 Or
                        InStr(EditWarning1, "E944INCOMPLETE DATA") > 0 Then
                        Continue For
                    End If
                    If GetText(2, 2, 2) = "MP" And Trim(GetText(10, 1, 80).Replace("-", "")) = "" Then Exit For

                    WriteLog("Placing data in MPC")
                    'Placing data in MPC
                    If serviceLinesForDraft.ContainsKey(MyDraft(J)) Then
                        Dim row As Integer = 10
                        For Each line In serviceLinesForDraft(MyDraft(J))
                            For Each kvp In line.Line1
                                SendKeyes(kvp.Value, row, kvp.Key)
                            Next
                            For Each kvp In line.Line2
                                SendKeyes(kvp.Value, row + 1, kvp.Key)
                            Next
                            row += 2
                            If row > 22 Then Exit For
                        Next
                    End If
                    Dim totalServiceLines As Integer = 0
                    Dim totalDeniedLine As Integer = 0
                    For I As Integer = 10 To 22 Step 2
                        If Trim(GetText(I, 1, 23).Replace("-", "")) = "" Then Exit For
                        If CDec(GetText(I, 41, 5) + "." + GetText(I, 47, 2)) = CDec(GetText(I, 50, 5) + "." + GetText(I, 56, 2).Replace("-", "0")) Then
                            totalDeniedLine += 1
                            totalServiceLines += 1
                        Else
                            totalServiceLines += 1
                        End If
                        If GetText(I, 29, 2) = "00" Then
                            SendKeyes("--", I, 29)
                        End If
                        If I = 22 Then Exit For
                    Next
                    SendEnter()
                    SendOverrideInScreen("02")
                    WriteLog("After Override 02")
                    WriteLog(GetTextRect)
                    If InStr(GetText(24, 1, 80), "NBAL NC") > 0 Then
                        WriteLog("NBAL edit present in HX icn")
                        strResearchComments = "NBAL edit present in HX icn"
                        Return False
                    End If

                    'Check TC field is same as the total of all the bill charge added
                    Dim totalBillCharge = addTotalBillCharge()
                    Dim wholenum = addZeroesToBillCharge(totalBillCharge).Split(".")(0)
                    Dim decimalNum = totalBillCharge.Split(".")(1)
                    If GetColorScheme(6, 4) = 3 Or InStr(GetText(24, 1, 80), "WE985TOT CHG BAL") > 0 Then
                        SendKeyes(wholenum, 6, 4)
                        SendKeyes(decimalNum, 6, 11)
                        SendEnter()
                    End If

                    'Check for interest line
                    'If present dashout the complete line and create 30line in the same place
                    Dim isTrue = checkForInterestLineAndDashOut()

                    Dim serviceLineRC As String = ""

                    'Create 30Line if required
                    If Not String.IsNullOrEmpty(ThirtyLine) Then
                        If CDec(thirtyLineAmt) > 0 Then
                            create30LineInMPCScreen(isSuccess, thirtyLineAmt, "0H")
                        End If
                    End If
                    WriteLog(GetTextRect)

                    serviceLineRC = GetText(10, 36, 2)
                    Dim reconCodeDict As New Dictionary(Of String, String) From {{"D2", "QG"}, {"W6", "22"}, {"IC", "QG"}, {"YL", "QG"}, {"YM", "QG"}}

                    'Reconsideration Code not applicable on the following:
                    '1. Bill Charge = Not Coverage
                    '2. Any service line Is having RC - UG Or W1
                    Dim ReconCodenotrequiredforRC As String = "0H,UG,W1,B9,QG"
                    For I As Integer = 10 To 22 Step 2
                        If Trim(GetText(I, 1, 23).Replace("-", "")) = "" Then Exit For
                        Dim reconCode As String = ""

                        If GetText(I, 50, 5).Replace(" ", "") <> "" Then
                            If InStr(ReconCodenotrequiredforRC, GetText(I, 36, 2)) = 0 And CInt(GetText(I, 41, 5) + "." + GetText(I, 47, 2)) <> CInt(GetText(I, 50, 5).Replace("-", "0") + "." + GetText(I, 56, 2)) Then
                                If reconCodeDict.ContainsKey(GetText(I, 36, 2)) Then
                                    reconCode = reconCodeDict(GetText(I, 36, 2))
                                Else
                                    reconCode = ObtainReconCode(GetText(I, 36, 2))

                                End If
                                WriteLog("Recon code=" + reconCode)
                                SendKeyes(reconCode, I, 36)
                                WriteLog(GetTextRect)
                            End If
                        Else
                            If InStr(ReconCodenotrequiredforRC, GetText(I, 36, 2)) = 0 And CInt(GetText(I, 41, 5) + "." + GetText(I, 47, 2)) > 0 Then
                                If reconCodeDict.ContainsKey(GetText(I, 36, 2)) Then
                                    reconCode = reconCodeDict(GetText(I, 36, 2))
                                Else
                                    reconCode = ObtainReconCode(GetText(I, 36, 2))

                                End If
                                WriteLog("Recon code=" + reconCode)
                                SendKeyes(reconCode, I, 36)
                                WriteLog(GetTextRect)
                            End If
                        End If
                    Next

                    Dim EditWarning As String = ""
                    Dim COUNT2 As Integer = 0
                    Dim flagOktopay As Boolean = False
                    Do
                        SendEnter()
                        If COUNT2 > 50 Then Exit Do
                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        EditWarning = EditWarning & GetText(24, 1, 80, True)
                        COUNT2 = COUNT2 + 1
                    Loop
                    Dim pdAmt As Decimal
                    If InStr(EditWarning, "W169") > 0 Then
                        Dim word As String() = New String() {"W169P1-"}
                        Dim wordArr As String() = EditWarning.Split(word, StringSplitOptions.RemoveEmptyEntries)
                        Dim pdAmtArr As String()
                        If wordArr.Length > 1 Then
                            pdAmtArr = wordArr(1).Trim.Split(",")
                        Else
                            pdAmtArr = wordArr(0).Trim.Split(",")
                        End If
                        If CheckIfAlphaCharacterIsPresent(pdAmtArr(0)) Then
                            strResearchComments = "Cannot extract paid amount from EDIT"
                            Return False
                        Else
                            pdAmt = pdAmtArr(0).Trim
                        End If

                    End If
                    WriteLog("Edit: " + EditWarning)
                    If InStr(EditWarning, "W118 OK TO PAY") > 0 OrElse (InStr(EditWarning, "WE") = 0 AndAlso InStr(EditWarning, ",E") = 0 AndAlso
                        InStr(EditWarning, ", E") = 0) Then
                        flagOktopay = True
                    End If

                    If InStr(EditWarning, "W1616PRV") > 0 And totalServiceLines <> totalDeniedLine Then 'W1616 - In Network - Float D2 / W515 -Out of Network - Float nothing
                        serviceLineRC = "D2"
                        SendKeyes("G" + serviceLineRC.ToString(), 2, 2)
                        SendEnter()
                        WriteLog(GetTextRect())
                        If InStr(GetText(24, 1, 80), "INVALID CONTROL LINE") > 0 Then
                            Dim controline = ExtractUntilTripleCommaCharByChar(Trim(GetText(2, 1, 80)))
                            SendKeyes("                                                           ", 2, 2)
                            SendKeyes(controline + ",,,", 2, 2)
                            SendKeyes("G" + serviceLineRC.ToString(), 2, 2)
                            SendEnter()
                        End If
                        If InStr(GetText(2, 1, 80).Trim(), "GD2") = 0 Then
                            strResearchComments = "GD2 float didn't complete"
                            Return False
                        End If
                    End If
                    If pdAmt = 0.00 And flagOktopay = True Then
                        SendMPP()
                        GetTextRect()
                        isSuccess = True

                        If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") > 0 Then
                            Return True
                        End If
                    Else
                        WriteLog("Pd Amt=" + pdAmt.ToString + ", flagOktoPay=" + flagOktopay.ToString)
                    End If
                Next
            Next

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + ": " + ex.Message)
            ' Optional: log error
            isSuccess = False

        End Try
        Return isSuccess
    End Function

    Function ExtractUntilTripleCommaCharByChar(input As String) As String
        Dim result As String = ""
        For i As Integer = 0 To input.Length - 1
            result &= input(i)
            If result.EndsWith(",,,") Then
                Return result.Substring(0, result.Length - 3)
            End If
        Next
        Return result ' If ',,,' not found, return full string
    End Function

    Public Function addTotalBillCharge() As String
        Dim totalCharge As Decimal = 0.0

        Try
            For I As Integer = 10 To 22 Step 2
                If GetText(I, 1, 2) <> "--" Then
                    totalCharge += CDec(GetText(I, 41, 5) + "." + GetText(I, 47, 2))
                End If
            Next
        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try
        Return totalCharge.ToString()
    End Function

    Private Function BuildFormattedServiceLine(row1 As Integer, row2 As Integer) As ServiceLineData
        Dim line1 As New Dictionary(Of Integer, String)()
        Dim line2 As New Dictionary(Of Integer, String)()

        Try
            ' Define coordinates and lengths
            Dim arrPOS As New Dictionary(Of Integer, Integer) From {
                {1, 2}, {4, 6}, {11, 6}, {18, 6}, {25, 3}, {29, 2}, {32, 1}, {34, 1}, {36, 2}, {39, 1},
                {41, 5}, {47, 2}, {50, 5}, {56, 2}, {59, 5}, {65, 2}, {68, 5}, {74, 2}, {77, 2}
            }

            Dim arrPOS2 As New Dictionary(Of Integer, Integer) From {
                {3, 3}, {9, 5}, {15, 2}, {20, 4}, {25, 2}, {30, 5}, {36, 2}, {41, 5}, {47, 2}, {50, 2},
                {55, 3}, {61, 5}, {67, 2}
            }

            Dim line1Coords As New HashSet(Of Integer) From {25, 29, 41, 50, 59, 65, 68}
            Dim line2Coords As New HashSet(Of Integer) From {3, 9, 20, 30, 41, 55, 61, 67, 72}

            Dim GetFormattedText = Function(r As Integer, c As Integer, l As Integer, pad As Boolean) As String
                                       Dim raw As String = Trim(GetText(r, c, l))
                                       If String.IsNullOrWhiteSpace(raw) Then raw = "0"
                                       If pad AndAlso IsNumeric(raw) Then
                                           Return raw.PadLeft(l, "0"c)
                                       Else
                                           Return raw
                                       End If
                                   End Function

            ' Fill line 1
            For Each kvp In arrPOS
                Dim col = kvp.Key
                Dim len = kvp.Value
                line1(col) = GetFormattedText(row1, col, len, line1Coords.Contains(col))
            Next

            ' Fill line 2
            For Each kvp In arrPOS2
                Dim col = kvp.Key
                Dim len = kvp.Value
                line2(col) = GetFormattedText(row2, col, len, line2Coords.Contains(col))
            Next

        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + ": " + ex.Message)
        End Try

        Return New ServiceLineData With {.Line1 = line1, .Line2 = line2}
    End Function

    Public Function ObtainReconCode(RC As String) As String
        Dim reconCode As String = ""
        WriteLog("Getting Recon Code from Unet")
        Try
            SendF12()
            autECLPSObj.setcursorpos(2, 1)
            apiChk()
            SendKeyes("[pa2]")
            If InStr(GetText(1, 1, 80), "ENTER SIGN-ON ID") > 0 Then
                SendF4()
            End If
            SendKeyes("RC," & RC, 23, 12)
            SendEnter()
            WriteLog(GetTextRect)
            Dim pattern As String = "(?i)code=(.{2})"

            For I As Integer = 5 To 19
                If GetText(I, 2, 2) = RC Then
                    For J As Integer = I To 19
                        Dim line As String = Trim(GetText(J, 3, 77))


                        Dim match As Match = Regex.Match(line, pattern)

                        If match.Success Then
                            reconCode = match.Groups(1).Value
                            Exit For
                        Else
                            reconCode = String.Empty
                        End If

                    Next
                    Exit For
                End If
            Next

            SendF3()
            SendF12()
        Catch ex As Exception
            WriteLog("Exception :" + ex.Message + " from method " + GetCurrentMethod.Name)
        End Try

        Return reconCode
    End Function

    Private Function checkForInterestLineAndDashOut() As Boolean
        Dim isDashedOut As Boolean = False
        WriteLog("inside " + GetCurrentMethod.Name)
        WriteLog(GetTextRect)
        Try
            For I As Integer = 10 To 22 Step 2

                If GetText(I, 4, 5) = "CXINT" Then
                    WriteLog("interest line present")
                    WriteLog(GetTextRect())
                    SendKeyes("------------------------------------------------------------", I, 1)

                    SendKeyes("-------------------------------------------------", I + 1, 3)
                    WriteLog(GetTextRect)
                    isDashedOut = True
                End If
            Next
        Catch ex As Exception
            WriteLog("Exception from " + GetCurrentMethod.Name + ": " + ex.Message)
            Return False
        End Try
        WriteLog("============Exit " + GetCurrentMethod.Name + "=============")
        Return isDashedOut
    End Function

    Public Function CreateServiceLineDT() As DataTable
        Dim claimTable As New DataTable()
        Try

            With claimTable.Columns
                .Add("Draft", GetType(String))
                .Add("Charge", GetType(String))
                .Add("RemarkCode", GetType(String))
                .Add("NotCovered", GetType(String))
                .Add("Deductible", GetType(String))
                .Add("Allowable", GetType(String))
                .Add("CoIns", GetType(String))
                .Add("Paid", GetType(String))
                .Add("TotalAllowable", GetType(String))
                .Add("TotalDeductible", GetType(String))
                .Add("TotalNotCovered", GetType(String))
                .Add("Copay", GetType(String))
                .Add("Modifier", GetType(String))
                .Add("NBR", GetType(String))
            End With

            For i As Integer = LBound(HistoryClaim) To UBound(HistoryClaim)
                If HistoryClaim(i).BillCharge = Nothing Then Exit For
                For j As Integer = LBound(HistoryClaim(i).RemarkCode) To UBound(HistoryClaim(i).RemarkCode)
                    If HistoryClaim(i).RemarkCode(j) = "69" Or HistoryClaim(i).RemarkCode(j) = "70" Then Continue For

                    ' Add the row to the DataTable
                    claimTable.Rows.Add(
                        HistoryClaim(i).Draft,
                        HistoryClaim(i).BillCharge(j),
                        HistoryClaim(i).RemarkCode(j),
                            HistoryClaim(i).NotCovered(j),
                            HistoryClaim(i).Deductible(j),
                             HistoryClaim(i).Allowable(j),
                                 HistoryClaim(i).CoIns(j),
                                  HistoryClaim(i).Paid(j),
                        HistoryClaim(i).TotalAllowable(j),
                       HistoryClaim(i).TotalDeductible(j),
                       HistoryClaim(i).TotalNotCovered(j),
                                 HistoryClaim(i).Copay(j),
                              HistoryClaim(i).Modifier(j),
                                   HistoryClaim(i).NBR(j)
                     )

                Next
            Next

        Catch ex As Exception
            WriteLog(ex.Message + "From " + GetCurrentMethod.Name)
        End Try
        Return claimTable
    End Function

    Public Function placeHXDataInMPC(serviceLinesForDraft As Dictionary(Of String, List(Of String)), MyDraft As String) As Boolean
        ' Define the column positions
        Dim arrPOS() As String = Split("1,4,11,18,25,29,32,34,36,39,41,47,50,56,59,65,68,74,77", ",")
        Dim arrPOS2() As String = Split("3,9,15,20,25,30,36,41,47,50,55,59,61,67,72,78", ",")

        Try
            ' Simulated dictionary of draft lines
            'Dim MyDraft As New Dictionary(Of Integer, List(Of String))()
            ' Example: MyDraft(1) = New List(Of String) From { "OH OPS    041525 042425   1 01 2 1 D2 N 50855.00      .00      .00      .00     %   0 B      .00 S     .00 M  8352.54 D      .00    % 100 P  8352.54 D     .00" }


            ' Check if the draft exists in the dictionary
            If Not serviceLinesForDraft.ContainsKey(MyDraft) Then Return False

            Dim mainLines As List(Of String) = serviceLinesForDraft(MyDraft)
            Dim row As Integer = 10

            For Each mainLine As String In mainLines
                If row > 22 Then Exit For

                ' Split the mainLine into two parts at the first '%'
                Dim part1 As String = ""
                Dim part2 As String = ""
                Dim percentIndex As Integer = mainLine.IndexOf("%"c)
                If percentIndex >= 0 Then
                    part1 = mainLine.Substring(0, percentIndex)
                    part2 = mainLine.Substring(percentIndex)
                Else
                    part1 = mainLine
                End If

                ' Normalize spacing and split into elements
                part1 = System.Text.RegularExpressions.Regex.Replace(part1, " {2,}", " ").Trim()
                part2 = System.Text.RegularExpressions.Regex.Replace(part2, " {2,}", " ").Trim()
                Dim arrTempStr1() As String = part1.Split(" "c)
                Dim arrTempStr2() As String = part2.Split(" "c)

                ' Place first part
                For L As Integer = 0 To Math.Min(arrTempStr1.Length - 1, arrPOS.Length - 1)
                    SendKeyes(arrTempStr1(L), row, CInt(arrPOS(L)))
                Next

                ' Place second part
                For L As Integer = 0 To Math.Min(arrTempStr2.Length - 1, arrPOS2.Length - 1)
                    SendKeyes(arrTempStr2(L), row + 1, CInt(arrPOS2(L)))
                Next

                row += 2

            Next

        Catch ex As Exception

        End Try
        Return True
    End Function

#End Region

End Module

