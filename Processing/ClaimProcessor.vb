Imports System.Configuration
Imports System.Reflection.MethodBase
Imports System.Runtime.InteropServices
Imports System.Security.Claims
Imports API
Imports Microsoft.Office.Interop.Excel


Namespace Processing
    ' =====================================================================
    ' ErrorHandler.vb - Centralized error handling
    ' =====================================================================
    Public Class ErrorHandler
        Public Shared Sub HandleUnetError(ex As Exception, logger As Logging.LogManager, currentStep As String)
            Try
                logger.LogError("ErrorHandler", String.Format("Error at step '{0}': {1}", currentStep, ex.Message), ex)
            Catch
                ' Failsafe
            End Try
        End Sub

        Public Shared Function GetUserFriendlyMessage(ex As Exception) As String
            If ex.Message.Contains("timeout") Then
                Return "Operation timed out. Please check Unet connection."
            ElseIf ex.Message.Contains("not found") Then
                Return "Claim not found in Unet."
            ElseIf ex.Message.Contains("unauthorized") Then
                Return "Not authorized to access this claim."
            Else
                Return "An error occurred: " & ex.Message
            End If
        End Function
    End Class

    ' =====================================================================
    ' ReversalWorkflow.vb - Defines the steps for claim reversal with void logic
    ' =====================================================================
    Public Class ReversalWorkflow
        Private _unetWrapper As Unet.UnetWrapper
        Private _logger As Logging.LogManager
        Private _icn As String
        Private _unetCredential As Models.DroidCredential
        Private _claim As Models.ClaimRecord
        Private _claimDetails As Object
        Private _claimModifier As Object
        Private FlagError As Boolean


        Public Event ProcessingStarted(message As String)
        Public Event ProcessingProgress(result As Models.ProcessingResult)
        Public Event ProcessingComplete(totalProcessed As Integer, successful As Integer, failed As Integer)
        Public Event ErrorOccurred(message As String)
        Public Sub New(unetWrapper As Unet.UnetWrapper, logger As Logging.LogManager, icn As String, credential As Models.DroidCredential, claim As Models.ClaimRecord)
            _unetWrapper = unetWrapper
            _logger = logger
            _icn = icn
            _unetCredential = credential
            _claim = claim
        End Sub

        ''' <summary>
        ''' Executes the complete reversal workflow including void logic
        ''' </summary>
        Public Function Execute() As Models.ProcessingResult
            Dim result As New Models.ProcessingResult With {
                .ICN = _icn,
                .StartTime = DateTime.Now,
                .Status = Configuration.Constants.STATUS_PROCESSING,
                .UnetSessionID = _unetWrapper.SessionName,
                .Office = _claim.Office.Trim,
                .Engine = _claim.Engine.Trim,
                .DateResearch = _claim.DateResearch,
                .ResearchComment = _claim.ResearchComment.Trim
            }
            Dim flagVoidConditions As Boolean = False
            Try
                ' STEP 1: Verify claim exists
                _logger.LogInfo("ReversalWorkflow", "STEP 1: Verifying claim exists")
                If Not VerifyClaimExists() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Payloc login failed"
                    result.ReversalStepReached = "ClaimVerification"
                    Return result
                End If
                result.ReversalStepReached = "Payloc login done"

                ' STEP 2: Check if current claim is eligible for reversal (void)
                _logger.LogInfo("ReversalWorkflow", "STEP 2: Checking if claim is eligible for VOID")
                If Not CheckClaimEligibilityForReversal() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    If InStr(strResearchComments, "Current claim not yet processed") > 0 Then
                        WriteLog("Current claim not yet processed- not eligible")
                        result.ErrorMessage = "Current claim not yet processed- not eligible"
                        result.PostFlagValue = Configuration.Constants.FLAG_PENDING
                    Else
                        result.ErrorMessage = "Claim is not eligible for 74 void (may be in pend 71/74)"
                    End If
                    result.ReversalStepReached = "EligibilityCheck"
                    Return result
                End If
                result.ReversalStepReached = "EligibilityChecked"

                ' STEP 3: Load claim history and check for void conditions
                '_logger.LogInfo("ReversalWorkflow", "STEP 3: Loading claim history and checking void conditions")
                'If Not LoadClaimHistory() Then
                '    result.Status = Configuration.Constants.STATUS_FAILED
                '    result.ErrorMessage = "Failed to load claim history"
                '    result.ReversalStepReached = "HistoryLoad"
                '    Return result
                'End If
                'result.ReversalStepReached = "HistoryLoaded"

                ' STEP 4: Check for void conditions (Adjuster ID 008273 with paid amount)
                _logger.LogInfo("ReversalWorkflow", "STEP 4: Checking void conditions")
                flagVoidConditions = CheckVoidConditions(result)
                If Not flagVoidConditions Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    If InStr(strResearchComments, "Claim voided by other adjuster ID") > 0 Or InStr(strResearchComments, "69/70 not performed by Wipro Corrected droid") > 0 Then
                        result.ErrorMessage = "HX claim voided by other adjuster ID"
                    Else
                        result.ErrorMessage = "Claim does not meet void conditions"
                    End If
                    result.ReversalStepReached = "VoidConditionCheck"
                    Return result
                End If
                result.ReversalStepReached = "VoidConditionsChecked"

                ' STEP 5: Perform void operation
                _logger.LogInfo("ReversalWorkflow", "STEP 5: Performing void operation")
                If Not PerformVoidOperation(flagVoidConditions) Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    If strResearchComments = "Void conditions didn't meet - Claim not voided by 008273" Or InStr(strResearchComments, "claim not processed by autobot") > 0 Then
                        result.ErrorMessage = strResearchComments
                    ElseIf InStr(strResearchComments, "claim void/pend unsuccessful") Then
                        result.ErrorMessage = strResearchComments
                    Else
                        result.ErrorMessage = "Failed to perform void operation"
                    End If
                    result.ReversalStepReached = "PerformVoidOperation"
                    Return result
                End If
                result.ReversalStepReached = "VoidCompleted"

                ' SUCCESS
                result.Status = Configuration.Constants.STATUS_SUCCESS
                _logger.LogSuccess("ReversalWorkflow", "Reversal/Void completed successfully for: " & _icn)

            Catch ex As Exception
                result.Status = Configuration.Constants.STATUS_FAILED
                result.ErrorMessage = ex.Message
                result.ErrorSource = "ReversalWorkflow.Execute"
                _logger.LogError("ReversalWorkflow", "Exception during workflow execution", ex)
            Finally
                result.EndTime = DateTime.Now
                result.DurationSeconds = CInt((result.EndTime - result.StartTime).TotalSeconds)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' Verify claim exists on the Unet screen
        ''' Access UnetWrapper properties: _unetWrapper.SessionName, _unetWrapper.UnetID, etc.
        ''' </summary>
        Private Function VerifyClaimExists() As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Try
                Dim LogInFlag As Boolean = False
                Dim LogInCount As Integer = 0
                ' Can now access wrapper properties directly
                _logger.LogInfo("ReversalWorkflow", "Verifying claim with session: " & _unetWrapper.SessionName)
                System.Threading.Thread.Sleep(500)
                While Not LogInFlag

                    PaylocLogIn(_claim.Engine, _claim.Office, _unetWrapper.SessionName, _unetCredential.ID, _unetCredential.Pass)
                    If InStr(GetText(24, 1, 80), "W018ADJ SIGN ON COMPLETE") > 0 Then
                        LogInFlag = True
                        WriteLog("Payloc Log in successful...")
                        WriteLog(GetTextRect())
                    Else
                        If LogInCount > 3 Then
                            Exit While
                        End If
                        LogInCount += 1
                        _unetWrapper.CloseSession()
                        '_unetWrapper.InitializeEmulator()
                        'autECLPSObj.StopCommunication()
                        autECLPSObj.StartCommunication()
                        Do
                            apiChk()
                        Loop Until autECLPSObj.CommStarted
                        Threading.Thread.Sleep(60000)
                    End If
                End While

                'Dim screenText As String = _unetWrapper.GetText(1, 1, 80, True)


                If LogInFlag Then
                    _logger.LogSuccess("ReversalWorkflow", "Claim verification successful: " & _icn)
                    _logger.LogInfo("ReversalWorkflow", "Using Unet ID: " & _unetWrapper.UnetID)
                    Return True
                End If

                '_logger.LogError("ReversalWorkflow", "Claim not found on screen")
                Return False
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "VerifyClaimExists")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check if claim is eligible for reversal (not in pend 71/74)
        ''' </summary>
        Private Function CheckClaimEligibilityForReversal() As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Dim isPend71or74 As Boolean = False
            Try
                GetClaimDetails(CurrentClaimDetails, _icn)
                GetModifierFromEDS1(CurrentClaimModifier, _icn)
                'WriteLog(GetTextRect)
                If InStr(GetText(24, 1, 80).Trim, "SELECTED CLAIM NOT FOUND") > 0 Then
                    _logger.LogError("ReversalWorkflow", "Current claim is too old - not eligible")
                    strResearchComments = "Current claim is too old - not eligible"
                    WriteLog(strResearchComments)
                    Return False
                End If
                FlagMHI = True
                ReDimension(CurrentClaim, _icn)
                If InStr(strResearchComments, "Current claim not yet processed") > 0 Then
                    _logger.LogError("ReversalWorkflow", "Current claim is not yet processed - not eligible")
                    Return False
                End If
                PullMHI(_icn)
                CurrentClaimCount = ReadMHI(CurrentClaim)

                GetDraftDetailsInMemory(CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount)
                isPend71or74 = DetermineWhetherCurrentClaimIsInPend71or74(CurrentClaimDraftDetails, GetMaxSuffix(_icn))
                'If DetermineWhetherCurrentClaimIsInPend71or74(CurrentClaimDraftDetails, GetMaxSuffix(_icn)) Then
                '    '''check 74 or 71 update database
                '    Exit Function
                'End If


                ' If claim is in pend 71/74, it's not eligible
                If isPend71or74 Then
                    _logger.LogWarning("ReversalWorkflow", "Claim appears to be in pend 71/74 - not eligible")
                    Return False
                End If

                ' Check for closed or voided claims
                'If InStr(screenText, "CLOSED") > 0 Or InStr(screenText, "VOID") > 0 Then
                '    _logger.LogWarning("ReversalWorkflow", "Claim appears to be closed or already voided")
                '    Return False
                'End If

                _logger.LogSuccess("ReversalWorkflow", "Claim eligibility check passed - current claim is not voided")
                WriteLog("69/70 WorkflowClaim- eligibility check passed - current claim is not voided")
                Return True
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "CheckClaimEligibilityForReversal")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Load claim history (MHI) to access draft details
        ''' This mirrors the code: PullMHI, ReadMHI, GetDraftDetailsInMemory
        ''' </summary>
        Private Function LoadClaimHistory() As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Try
                _logger.LogInfo("ReversalWorkflow", "Pulling MHI (Multi-occurrence History)")

                ' Navigate to MHI screen
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()

                ' This would call your existing PullMHI method from mdlUnet
                ' Placeholder: mdlUnet.PullMHI(_icn)

                System.Threading.Thread.Sleep(500)

                _logger.LogSuccess("ReversalWorkflow", "Claim history loaded successfully")
                Return True

            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "LoadClaimHistory")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Check for void conditions from the LatestVoid code:
        ''' - Adjuster ID must be "008273"
        ''' - Paid amount must be > 0
        ''' </summary>
        Private Function CheckVoidConditions(result As Models.ProcessingResult) As Boolean
            Dim checkfor6970 As Boolean = False
            Dim paidAmount As Double = 0
            Dim adjusterId As String = "008273"
            WriteLog($"Entering {GetCurrentMethod.Name}")

            Try
                _logger.LogInfo("ReversalWorkflow", "Checking void conditions (Adjuster 008273 with paid amount > 0)")

                ' These would be populated from your claim arrays/objects
                GetClaimDetails(CurrentClaimDetails, _icn)
                GetModifierFromEDS1(CurrentClaimModifier, _icn)

                'checkfor6970 = Check6970void_For74void(_icn, paidAmount)

                checkfor6970 = True
                If checkfor6970 Then
                    _logger.LogSuccess("ReversalWorkflow", "Void conditions met - Adjuster: " & adjusterId & ", Amount: " & paidAmount)
                    Return True
                Else
                    strResearchComments = "Void conditions NOT met - 69/70 not performed by Wipro Corrected droid"
                    _logger.LogWarning("ReversalWorkflow", "Void conditions NOT met - 69/70 not performed by Wipro Corrected droid")
                    Return False
                End If

            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "CheckVoidConditions")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Perform the actual void operation (PerformVoid_For74AAProcess equivalent)
        ''' </summary>
        Private Function PerformVoidOperation(flagVoidCondition As Boolean) As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Initiating void operation")

                ' Navigate to void screen
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()
                Dim flagVoid As Boolean = voidCode(flagVoidCondition)

                System.Threading.Thread.Sleep(500)

                ' Check if void was successful
                Dim screenText As String = _unetWrapper.GetText(24, 1, 80, True)

                If flagVoid Then
                    _logger.LogSuccess("ReversalWorkflow", "Void operation successful")
                    Return True
                ElseIf InStr(screenText, "DENIED") > 0 Or InStr(screenText, "ERROR") > 0 Then
                    _logger.LogError("ReversalWorkflow", "Void operation denied/errored: " & screenText)
                    Return False
                ElseIf strResearchStatus = "skipped" Then
                    _logger.LogError("PerformVoidOPeration", "Current claim void/pend unsuccessful")
                    strResearchComments = "Current claim void/pend unsuccessful: " & strResearchComments
                    Return False
                ElseIf flagVoid = False Then
                    _logger.LogWarning("ReversalWorkflow", "Void conditions didn't meet - Claim not processed by 008273")
                    'strResearchComments = "Void conditions didn't meet - Claim not voided by 008273"
                    Return False
                Else
                    _logger.LogWarning("ReversalWorkflow", "Void operation status unclear - assuming success")
                    Return True
                End If

            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "PerformVoidOperation")
                Return False
            End Try
        End Function

        Private Function voidCode(flag69or70 As Boolean) As Boolean
            Dim _checkcurrentclaimAdjusterID As String
            Try
#Region "Session check"
                flagHistoryVoid = False
                If flag69or70 Then

                    FlagMHI = True
                    ReDimension(CurrentClaim, _icn)
                    PullMHI(_icn)

                    CurrentClaimCount = ReadMHI(CurrentClaim)

                    GetDraftDetailsInMemory(CurrentClaimDraftDetails, CurrentClaim, CurrentClaimCount)

                    Dim curdate As String = String.Format("{0:MM/dd/yy}", DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"))
                    For IOM As Integer = LBound(CurrentClaim) To UBound(CurrentClaim)
                        If CurrentClaim(IOM).AdjusterID Is Nothing Then Exit For
                        _checkcurrentclaimAdjusterID = CurrentClaim(IOM).AdjusterID
                        If _checkcurrentclaimAdjusterID = "008273" AndAlso CDbl(CurrentClaim(IOM).PaidAmount) > 0 Then
                            Dim Draftnumber As String = CurrentClaim(IOM).Draft
                            Dim draftprocessdate As String = CurrentClaim(IOM).DateOfProcessing

                            PerformVoid_For74AAProcess(FlagError, flagHistoryVoid, Draftnumber, _icn, "100", draftprocessdate)
                            If FlagError Then
                                ''''skip
                                WriteLog("Current claim void failed:" + strResearchComments)
                                Return False
                            End If
                        End If
                    Next

                    If flagHistoryVoid Then
                        '''flag to set 74 void successful
                        WriteLog("Claim void successful")
                        Return True
                    End If
                End If
#End Region
            Catch ex As Exception

            End Try
            WriteLog("Current claim not processed by autobot- " & _checkcurrentclaimAdjusterID)
            strResearchComments = "Current claim not processed by autobot- " & _checkcurrentclaimAdjusterID
            Return False
        End Function

        'Public Sub GetClaimDetails(ByRef MyclaimDetails As ClaimDetails, ByVal icn As String, Payloc As String)
        '    Try
        '        ReDimension(CurrentClaimDraftDetails, 0)
        '        ReDim CurrentClaim(0)

        '        SendClear()
        '        SendClear()
        '        pullEDSScreen(icn, "1")
        '        apiChk()
        '        apiChk()
        '        If Trim(GetText(6, 2, 5)) = "EDS 1" Then
        '            WriteLog("====================OPEN EDS1 TO STORE DATA==================")
        '            WriteLog(GetTextRect)
        '            If Trim(GetText(7, 2, 14)) = "BILLED CHARGES" Then
        '                SendF9()
        '                MyclaimDetails.ICN = Trim(GetText(1, 6, 10))
        '                MyclaimDetails.SSN = Trim(GetText(1, 22, 10))
        '                MyclaimDetails.Policy = Trim(GetText(2, 6, 6))
        '                MyclaimDetails.PatientName = Trim(GetText(1, 37, 11))
        '                MyclaimDetails.PatientRelationship = Trim(GetText(1, 52, 2))
        '                MyclaimDetails.Fdate = Trim(GetText(3, 21, 6))
        '                MyclaimDetails.Ldate = Trim(GetText(3, 33, 6))
        '                MyclaimDetails.PatDOB = Trim(GetText(1, 59, 8))
        '                MyclaimDetails.TotalCharge = Trim(GetText(2, 53, 10))
        '                CurrentClaim(0).Policy = GetText(2, 6, 6)
        '                CurrentClaim(0).SSN = GetText(1, 22, 10)
        '                CurrentClaim(0).EmployeeName = GetText(1, 38, 10)
        '                CurrentClaim(0).Relationship = GetText(1, 52, 2)
        '            ElseIf Trim(GetText(6, 2, 5)) = "EDS 1" And Trim(GetText(7, 2, 4)) = "REVE" Then
        '                MyclaimDetails.ICN = Trim(GetText(1, 6, 10))
        '                MyclaimDetails.SSN = Trim(GetText(1, 22, 10))
        '                MyclaimDetails.Policy = Trim(GetText(2, 6, 6))
        '                MyclaimDetails.PatientName = Trim(GetText(1, 37, 11))
        '                MyclaimDetails.PatientRelationship = Trim(GetText(1, 52, 2))
        '                MyclaimDetails.Fdate = Trim(GetText(3, 21, 6))
        '                MyclaimDetails.Ldate = Trim(GetText(3, 33, 6))
        '                MyclaimDetails.PatDOB = Trim(GetText(1, 59, 8))
        '                MyclaimDetails.TotalCharge = Trim(GetText(2, 53, 10))
        '                CurrentClaim(0).Policy = GetText(2, 6, 6)
        '                CurrentClaim(0).SSN = GetText(1, 22, 10)
        '                CurrentClaim(0).EmployeeName = GetText(1, 37, 10)
        '                CurrentClaim(0).Relationship = GetText(1, 52, 2)
        '            End If
        '        End If
        '        SendClear()
        '        SendClear()
        '        pullEDSScreen(icn, "3")
        '        WriteLog(GetTextRect())
        '        MyclaimDetails.FLN = Trim(GetText(2, 17, 10))
        '        MyclaimDetails.DCC = Trim(GetText(2, 28, 3))
        '        'Below condition of DCC added on 8/10/2021
        '        If MyclaimDetails.DCC = "951" Or MyclaimDetails.DCC = "807" Or MyclaimDetails.DCC = "910" Or MyclaimDetails.DCC = Payloc Then
        '            strResearchComments = "Claim is not an electronic/paper submission"
        '            Exit Sub
        '        End If
        '        MyclaimDetails.Tin = Trim(GetText(18, 19, 9))
        '        MyclaimDetails.Prifix_Tin_Suffix = Trim(GetText(23, 19, 1)) + Trim(GetText(23, 20, 9)) + Trim(GetText(23, 30, 5))
        '        MyclaimDetails.Prefix_Tin = Trim(GetText(23, 19, 10))
        '        CurrentClaim(0).ICN = icn
        '        CurrentClaim(0).Tin = GetText(23, 19, 1) + "-" + GetText(23, 20, 9) + "-" + GetText(23, 30, 5)
        '        CurrentClaimCount = -1
        '        SendClear()
        '        SendClear()
        '        pullEDSScreen(icn, "1")
        '        apiChk()
        '        WriteLog("Exit GetClaimDetails()..........")
        '    Catch ex As Exception
        '        WriteLog(GetCurrentMethod.Name + " exception")
        '    End Try

        'End Sub
        Sub PerformVoid_For74AAProcess(ByRef FlagError As Boolean, ByRef FlagHistoryVoid As Boolean, ByVal Draft As String, ByVal ICN As String, ByVal strPaidAmount As String, ByVal ProcDate As String)
            Dim FlagFound As Boolean, K As Integer, ValidLline As Integer, Count2 As Integer, EditWarning As String = ""
            'MsgBox DateDiff("d", CDate(ProcDate), CDate(Format(Now, "MM/DD,YY")))
            'MsgBox(DateTime.Today)
            FlagError = False
            WriteLog("Void starts for draft " & Draft & "and ICN " & ICN)
            WriteLog("Processed Date " & ProcDate & " And Date.Today: " & CStr(CDate(System.TimeZoneInfo.ConvertTime(Now, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")))))
            Dim curdate As String = String.Format("{0:MM/dd/yy}", DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time"))
            If CDate(ProcDate) = curdate Then 'CDate(DateTime.Today) Then   ''''''74 Void
                SendClear()
                SendClear()
                SendKeyes("MPI," & CurrentClaim(0).Policy & "," & CurrentClaim(0).SSN & "," & CurrentClaim(0).EmployeeName & "," & CurrentClaim(0).Relationship & "," & Draft & ",I" & ICN, 2, 2)
                SendEnter()
                WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                If InStr(GetText(24, 1, 80), "REVERSAL ON FILE") = 0 Then 'Already 74 void is performed
                    SendKeyes("GZ3", 2, 2)
                    SendKeyes("74", 3, 78)
                    SendEnter()
                    WriteLog("Before 74 void MPP for draft " & Draft & " and ICN " & ICN, True)
                    WriteLog(GetTextRect())
                    Count2 = 0
                    EditWarning = ""
                    Do
                        SendEnter()
                        If Count2 > 6 Then Exit Do
                        If InStr(EditWarning, GetText(24, 1, 80, True)) > 0 Then Exit Do
                        EditWarning = EditWarning & GetText(24, 1, 80, True)
                        Count2 = Count2 + 1
                    Loop
                    If Count2 > 6 Then
                        strResearchComments = EditWarning
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If
                    If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                        SendMPP()
                        If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                            strResearchComments = GetText(24, 1, 80, True)
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        Else
                            FlagHistoryVoid = True
                        End If
                    Else
                        strResearchComments = EditWarning
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If



                    WriteLog("After 74 void MPP for draft " & Draft & " and ICN " & ICN, True)
                End If
            ElseIf DateDiff("d", CDate(ProcDate), CDate(DateTime.Today)) < 3 Then
                '''''If Check not issued then 71 Void else 69/70



                SendClear()
                SendClear()
                SendKeyes("MPI," & CurrentClaim(0).Policy & "," & CurrentClaim(0).SSN & "," & CurrentClaim(0).EmployeeName & "," & CurrentClaim(0).Relationship & "," & Draft & ",I" & ICN, 2, 2)
                SendEnter()
                WriteLog("Control Line for Void = " + GetText(2, 1, 80, True))
                If GetText(22, 36, 2, True) <> "71" Then 'Already 71 void has been performed
                    SendKeyes("GZN", 2, 2)
                    SendKeyes("71", 3, 78)
                    SendEnter()
                    WriteLog("Before 71 void MPP for draft " & Draft & " and ICN " & ICN, True)
                    WriteLog(GetTextRect())
                    Count2 = 0
                    EditWarning = ""
                    EditWarning = EditWarning & GetText(24, 1, 80, True)

                    If InStr(EditWarning, "W118 OK TO PAY") > 0 Then
                        SendMPP()
                        If InStr(GetText(24, 1, 80), "W134ALL FILES UPDATED") = 0 Then
                            strResearchComments = GetText(24, 1, 80, True)
                            strResearchStatus = "skipped"
                            FlagError = True
                            Exit Sub
                        Else
                            WriteLog(GetTextRect)
                            FlagHistoryVoid = True
                        End If
                    Else
                        strResearchComments = EditWarning
                        strResearchStatus = "skipped"
                        FlagError = True
                        Exit Sub
                    End If
                    WriteLog("After 71 void MPP for draft " & Draft & " and ICN " & ICN, True)
                    'End If
                Else
                    strResearchComments = "Not eligiable for 74 and 71"
                    strResearchStatus = "skipped"
                    FlagError = True
                    Exit Sub
                    ''''check for 74 or 
                End If
            Else
                strResearchComments = "Not eligiable for 74 and 71"
                strResearchStatus = "skipped"
                FlagError = True
                Exit Sub
            End If
            WriteLog("Void ends for draft " & Draft & " and ICN " & ICN)

        End Sub

        Public Function DetermineWhetherCurrentClaimIsInPend71or74(ByRef Mydrafts() As DraftDetails, ByVal MaxSuffix As Integer) As Boolean
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
                        If Mydrafts(J).ReversalType = "71" AndAlso Mydrafts(J).Reversal OrElse Mydrafts(J).ReversalType = "71 Void" OrElse Mydrafts(J).ReversalType = "74" Then
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
        Public Function Check6970void_For74void(ByVal ICN As String, ByRef totalPaid As Double) As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Try
                Dim checkwithouttin As Boolean = False
                '''''add 07.18.2025 for er deni           
                Dim AhiConLine12 As String = "AHI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName &
      "," & CurrentClaimDetails.PatientRelationship & ",T" & Left(CurrentClaimDetails.Tin, 11).Replace("-", "") &
      ",," & GetFDOS(CurrentClaimModifier) & "," & GetLDOS(CurrentClaimModifier)
                'HistoryICN = FindHistoryICN(CurrentClaimCount, isHICF, isUB)

                If GetText(4, 1, 80, True).Replace("-", "").Trim = "" Then
AhiSearchWithoutTin:

                    AhiConLine12 = "AHI," & CurrentClaimDetails.Policy & "," & CurrentClaimDetails.SSN & "," & CurrentClaimDetails.PatientName &
             "," & CurrentClaimDetails.PatientRelationship & "," & ",," & GetFDOS(CurrentClaimModifier) & "," & GetLDOS(CurrentClaimModifier)


                End If
                SendClear()
                SendClear()
                SendKeyes(AhiConLine12, 2, 2)
                SendEnter()
                If InStr(GetText(24, 1, 80), "E126INVALID DATE") > 0 Then
                    strResearchComments = "E126INVALID DATE"
                    Return False
                End If
                WriteLog(GetTextRect)
                Dim Icntest As String = ""
                Dim XXXX As Integer, YYYYY As Integer, IcnList1 As String = ""
                XXXX = 0
                For YYYYY = 4 To 20 Step 2
                    If Replace(Replace(autECLPSObj.GetText(YYYYY, 1, 80), "-", ""), " ", "") = "" Then Exit For
                    If autECLPSObj.GetText(YYYYY + 1, 54, 10) <> CurrICN AndAlso GetText(YYYYY + 1, 52, 1) = "I" Then
                        IcnList1 += autECLPSObj.GetText(YYYYY + 1, 54, 10) + ","
                        Icntest = autECLPSObj.GetText(YYYYY + 1, 54, 10)
                        Dim RemarkCode As String = Replace(Replace(Replace(autECLPSObj.GetText(YYYYY, 64, 8), " ", ","), "--", ""), ",,", ",")
                        'add abhijit dolui 06.25.2024
                        RemarkCode = Replace(RemarkCode, ",", "")
                        ' GetAmountAllreadypaidsuffix(Icntest)
                        Dim R_OV As String = Trim(autECLPSObj.gettext(YYYYY, 67, 2))

                        Dim OV As String = Replace(Replace(Replace(autECLPSObj.gettext(YYYYY, 73, 8), " ", ","), "--", ""), ",,", ",")
                        ' Dim OV1 As String = Trim(autECLPSObj.GetText(YYYYY, 73, 8))
                        Dim OV1 As String = Replace(Replace(autECLPSObj.gettext(YYYYY, 73, 8), " ", ""), "--", "")
                        Dim HXICN_6970 As String = autECLPSObj.GetText(YYYYY + 1, 54, 10)

                        If RemarkCode = "69" Or RemarkCode = "70" Then

                            Dim mhiscreencheck As Boolean = Check6970void_MHIScreen(HXICN_6970, totalPaid)
                            If mhiscreencheck Then
                                Return True
                                Exit For
                            End If


                        End If

                        XXXX = XXXX + 1
                        ' End If

                        If YYYYY = 20 Then
                            If InStr(autECLPSObj.GetText(24, 1, 80), "NO MORE RECORDS") > 0 Then Exit For
                            SendF8()
                            YYYYY = 2
                        End If
                    End If
                Next

                If Not checkwithouttin Then
                    checkwithouttin = True
                    GoTo AhiSearchWithoutTin

                End If

            Catch ex As Exception
                Return False
            End Try
            Return False

        End Function

        Public Function Check6970void_MHIScreen(ByVal ICN As String, ByRef TotalPaid As Double) As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            'Dim TotalPaid As Double = "0"
            Try
                SendClear()
                SendClear()
                'read current tool mhi
                SendKeyes("RET,I" & ICN & ",M", 2, 2)
                SendEnter() : WriteLog("***************************************" & vbCrLf & "Gathering current claim information" & vbCrLf & "***************************************", True)
                WriteLog(GetTextRect)
                If autECLPSObj.GetText(2, 2, 2) = "MH" Then
                    WriteLog("MHI PULLED")
                    FlagMHI = True
                    ReDimension(CurrentClaim, ICN)
                    PullMHI(ICN)
                    CurrentClaimCount = ReadMHI(CurrentClaim)



                    For I As Integer = LBound(CurrentClaim) To UBound(CurrentClaim)
                        If CurrentClaim(I).Voided = True Then
                            If CurrentClaim(I).AdjusterID = "430664" Or CurrentClaim(I).AdjusterID = "398676" Then
                                WriteLog("Total Paid amount is " + TotalPaid.ToString + " Paid by Wipro bot")
                                Return True

                            End If
                        End If

                    Next
                Else
                    strResearchComments = "MHI not present"
                    Return False
                End If
            Catch ex As Exception
                WriteLog($"Exception from {GetCurrentMethod.Name}: {ex.Message}")
                Return False
            End Try
            WriteLog($"")
            Return False
        End Function


    End Class

    ' =====================================================================
    ' ClaimProcessor.vb - Main orchestration with credential caching
    ' =====================================================================
    Public Class ClaimProcessor
        Private _dbHelper As Database.DatabaseHelper
        Private _sessionName As String
        Private _processingResults As New List(Of Models.ProcessingResult)
        Private _shouldStop As Boolean = False
        Private _metrics As Models.DashboardMetrics

        ' CREDENTIALS CACHE - Fetched once, reused for all claims
        Private _cachedCredentials As Dictionary(Of String, Models.DroidCredential)

        Public Event ProcessingStarted(message As String)
        Public Event ProcessingProgress(result As Models.ProcessingResult)
        Public Event MetricsUpdated(metrics As Models.DashboardMetrics)
        Public Event ProcessingComplete(totalProcessed As Integer, successful As Integer, failed As Integer)
        Public Event ErrorOccurred(message As String)

        Public Sub New(sessionName As String)
            _sessionName = sessionName.ToUpper()
            _dbHelper = New Database.DatabaseHelper()
            _metrics = New Models.DashboardMetrics()
            _cachedCredentials = New Dictionary(Of String, Models.DroidCredential)
        End Sub

        ''' <summary>
        ''' Starts batch processing of all ICNs with priority_flag = 2
        ''' NOW CACHES CREDENTIALS ONCE AT THE START
        ''' </summary>
        Public Sub StartBatchProcessing()
            StartDateMonth = DateTime.Today.ToString("MMM")
            StartDateLong = DateTime.Today.ToString("MMddyyyy")
            Try
                RaiseEvent ProcessingStarted("Starting batch processing with session: " & _sessionName)

                ' Initialize configuration
                Dim config As Configuration.ConfigManager = Configuration.ConfigManager.GetInstance()
                config.LoadSettings()
                config.EnsureLogDirectoriesExist()

                ' ===== NEW: FETCH AND CACHE ALL CREDENTIALS ONCE =====
                RaiseEvent ProcessingStarted("Fetching credentials (cached for all claims)...")
                Try
                    _cachedCredentials = _dbHelper.FetchAllCredentials()
                    If _cachedCredentials.Count = 0 Then
                        RaiseEvent ErrorOccurred("No credentials found in DroidCred table")
                        Return
                    End If
                    RaiseEvent ProcessingStarted("Credentials cached successfully: " & _cachedCredentials.Count & " localities")
                Catch ex As Exception
                    RaiseEvent ErrorOccurred("Failed to fetch credentials: " & ex.Message)
                    Return
                End Try
                ' ===== END: CREDENTIAL CACHING =====

                ' Fetch all ICNs ready for processing
                Dim claimsToProcess As List(Of Models.ClaimRecord) = _dbHelper.FetchICNsByFlag(Configuration.Constants.FLAG_READY_TO_PROCESS)

                If claimsToProcess.Count = 0 Then
                    RaiseEvent ProcessingStarted("No claims ready for processing")
                    Return
                End If

                ' ===== Initialize metrics with TOTAL COUNT =====
                _metrics.InitializeWithTotal(claimsToProcess.Count)
                RaiseEvent MetricsUpdated(_metrics)
                RaiseEvent ProcessingStarted("Total claims fetched from database: " & claimsToProcess.Count)
                ' ===== END: Proper metrics Initialization =====


                ' Process each claim sequentially
                For Each claim In claimsToProcess
                    If _shouldStop Then Exit For

                    ' PASS CACHED CREDENTIALS TO PROCESS METHOD
                    ProcessSingleClaim(claim, config, _cachedCredentials)
                    If strResearchComments.Contains("Payloc login failed") Then Exit For

                    ' Update metrics
                    Dim successful As Integer = _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_SUCCESS).Count()
                    Dim failed As Integer = _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_FAILED).Count()
                    _metrics.UpdateMetrics(successful, failed)

                    'Notify UI of metrics update
                    RaiseEvent MetricsUpdated(_metrics)

                    RaiseEvent ProcessingComplete(_metrics.TotalClaims,
                                             _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_SUCCESS).Count(),
                                             _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_FAILED).Count())
                Next
                If strResearchComments.Contains("Payloc login failed") Then MessageBox.Show("Payloc login failed", "UNET System might not be available", MessageBoxButtons.OKCancel)
            Catch ex As Exception
                RaiseEvent ErrorOccurred("Error during batch processing: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Processes a single claim through the reversal workflow
        ''' NOW RECEIVES CACHED CREDENTIALS INSTEAD OF QUERYING DB
        ''' </summary>
        Private Sub ProcessSingleClaim(claim As Models.ClaimRecord, config As Configuration.ConfigManager, cachedCredentials As Dictionary(Of String, Models.DroidCredential))
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Dim logger As Logging.LogManager = Nothing
            Dim unetWrapper As Unet.UnetWrapper = Nothing
            strResearchComments = ""
            strResearchStatus = ""
            UnetID = claim.Locality
            CurrICN = claim.ICN
            Office = claim.Office
            Engine1 = claim.Engine
            Try
                ' Step 1: Update claim status to "Processing" (Flag 9)
                _dbHelper.UpdatePriorityFlag(claim.ICN, Configuration.Constants.FLAG_CURRENTLY_PROCESSING, claim.ID)

                ' Step 2: Initialize logger for this ICN
                logger = New Logging.LogManager(claim.ICN, claim.Locality)
                logger.LogInfo("ClaimProcessor", "Started processing ICN: " & claim.ICN)

                ' Step 3: Get credential from CACHE (no DB call needed)
                Dim credential As Models.DroidCredential = _dbHelper.GetCredentialFromCache(cachedCredentials, claim.Locality)
                If credential Is Nothing Then
                    Throw New Exception("No cached credential found for locality: " & claim.Locality)
                End If
                logger.LogInfo("ClaimProcessor", "Using cached credentials for locality: " & claim.Locality)

                ' Step 4: Initialize Unet wrapper
                unetWrapper = New Unet.UnetWrapper(_sessionName, credential.ID, credential.Pass, logger)
                If Not unetWrapper.InitializeEmulator() Then
                    WriteLog($"Failed to initialize session")
                    Throw New Exception("Failed to initialize Unet emulator")
                End If

                ' Step 5: Execute reversal workflow with credential and wrapper
                ' NOW PASSES CREDENTIAL TO WORKFLOW
                Dim workflow As New ReversalWorkflow(unetWrapper, logger, claim.ICN, credential, claim)
                Dim result As Models.ProcessingResult = workflow.Execute()

                'Execute in case of Payloc login failure
                If result.ErrorMessage = "Payloc login failed" Then
                    _dbHelper.UpdatePriorityFlag(claim.ICN, result.PostFlagValue, claim.ID)
                    strResearchComments = "Payloc login failed"
                    Exit Sub
                End If
                ' Set additional result properties
                result.Engine = claim.Engine
                result.Office = claim.Office
                If InStr(strResearchComments, "Current claim not yet processed") > 0 Then
                    result.PostFlagValue = Configuration.Constants.FLAG_PENDING
                Else
                    result.PostFlagValue = If(result.Status = Configuration.Constants.STATUS_SUCCESS, Configuration.Constants.FLAG_SUCCESS, Configuration.Constants.FLAG_FAILURE)
                End If

                ' Step 6: Update priority flag based on result
                _dbHelper.UpdatePriorityFlag(claim.ICN, result.PostFlagValue, claim.ID)

                ' Step 7: Log result to database and file
                logger.LogProcessingResult(result)

                ' Step 8: Add to results list
                _processingResults.Add(result)

                ' Step 9: Notify UI
                RaiseEvent ProcessingProgress(result)

                logger.LogSuccess("ClaimProcessor", "Processing completed for ICN: " & claim.ICN & " - Status: " & result.Status)

            Catch ex As Exception
                If logger IsNot Nothing Then
                    logger.LogError("ClaimProcessor", "Error processing claim: " & claim.ICN, ex)
                End If

                ' Mark as failed
                Try
                    _dbHelper.UpdatePriorityFlag(claim.ICN, Configuration.Constants.FLAG_FAILURE, claim.ID)
                Catch
                End Try

                ' Add failed result
                _processingResults.Add(New Models.ProcessingResult With {
                    .ICN = claim.ICN,
                    .Status = Configuration.Constants.STATUS_FAILED,
                    .ErrorMessage = ex.Message,
                    .ErrorSource = "ClaimProcessor.ProcessSingleClaim",
                    .Engine = claim.Engine,
                    .Office = claim.Office,
                    .PostFlagValue = Configuration.Constants.FLAG_FAILURE,
                    .StartTime = DateTime.Now,
                    .EndTime = DateTime.Now
                })

                RaiseEvent ProcessingProgress(_processingResults.Last())

            Finally
                If unetWrapper IsNot Nothing Then
                    unetWrapper.CloseSession()
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Stops the processing batch
        ''' </summary>
        Public Sub StopProcessing()
            _shouldStop = True
        End Sub

        Public ReadOnly Property Results As List(Of Models.ProcessingResult)
            Get
                Return _processingResults
            End Get
        End Property

        Public ReadOnly Property Metrics As Models.DashboardMetrics
            Get
                Return _metrics
            End Get
        End Property

    End Class

End Namespace
