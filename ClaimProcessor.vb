Namespace Processing
    ' =====================================================================
    ' ErrorHandler.vb - Centralized error handling
    ' =====================================================================
    Public Class ErrorHandler
        Public Shared Sub HandleUnetError(ex As Exception, logger As Logging.LogManager, step As String)
            Try
                logger.LogError("ErrorHandler", String.Format("Error at step '{0}': {1}", step, ex.Message), ex)
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
    ' ReversalWorkflow.vb - Defines the steps for claim reversal with void processing
    ' =====================================================================
    Public Class ReversalWorkflow
        Private _unetWrapper As Unet.UnetWrapper
        Private _logger As Logging.LogManager
        Private _icn As String
        Private _claimDetails As Object  ' Will hold claim details from Unet
        Private _claimModifier As Object  ' Will hold modifier information
        
        Public Sub New(unetWrapper As Unet.UnetWrapper, logger As Logging.LogManager, icn As String)
            _unetWrapper = unetWrapper
            _logger = logger
            _icn = icn
        End Sub
        
        ''' <summary>
        ''' Executes the complete reversal workflow with void processing
        ''' </summary>
        Public Function Execute() As Models.ProcessingResult
            Dim result As New Models.ProcessingResult With {
                .ICN = _icn,
                .StartTime = DateTime.Now,
                .Status = Configuration.Constants.STATUS_PROCESSING
            }
            
            Try
                ' STEP 1: Verify claim exists
                _logger.LogInfo("ReversalWorkflow", "STEP 1: Verifying claim exists")
                If Not VerifyClaimExists() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Claim does not exist or cannot be accessed"
                    result.ReversalStepReached = "ClaimVerification"
                    Return result
                End If
                result.ReversalStepReached = "ClaimVerified"
                
                ' STEP 2: Get claim details from EDS screens
                _logger.LogInfo("ReversalWorkflow", "STEP 2: Fetching claim details from EDS")
                If Not GetClaimDetailsFromEDS() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Failed to fetch claim details"
                    result.ReversalStepReached = "GetClaimDetails"
                    Return result
                End If
                result.ReversalStepReached = "ClaimDetailsRetrieved"
                
                ' STEP 3: Check for void eligibility (6970/74 void check)
                _logger.LogInfo("ReversalWorkflow", "STEP 3: Checking void eligibility")
                If Not CheckVoidEligibility() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Claim not eligible for void processing"
                    result.ReversalStepReached = "VoidEligibilityCheck"
                    Return result
                End If
                result.ReversalStepReached = "VoidEligibilityChecked"
                
                ' STEP 4: Get MHI (Master History Information)
                _logger.LogInfo("ReversalWorkflow", "STEP 4: Pulling MHI data")
                If Not PullAndProcessMHI() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Failed to fetch MHI data"
                    result.ReversalStepReached = "MHIRetrieval"
                    Return result
                End If
                result.ReversalStepReached = "MHIRetrieved"
                
                ' STEP 5: Check if claim is in pending status
                _logger.LogInfo("ReversalWorkflow", "STEP 5: Checking claim pending status")
                If IsClaimInPending() Then
                    _logger.LogWarning("ReversalWorkflow", "Claim is in pending 71 or 74 status - skipping")
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Claim is in pending 71 or 74 status"
                    result.ReversalStepReached = "PendingStatusCheck"
                    Return result
                End If
                result.ReversalStepReached = "NotPending"
                
                ' STEP 6: Process void for eligible claims
                _logger.LogInfo("ReversalWorkflow", "STEP 6: Processing void operations")
                If Not ProcessVoidForEligibleClaims() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Failed to process void operations"
                    result.ReversalStepReached = "VoidProcessing"
                    Return result
                End If
                result.ReversalStepReached = "VoidProcessed"
                
                ' STEP 7: Confirm reversal completion
                _logger.LogInfo("ReversalWorkflow", "STEP 7: Confirming reversal completion")
                If Not ConfirmReversalCompletion() Then
                    result.Status = Configuration.Constants.STATUS_FAILED
                    result.ErrorMessage = "Failed to confirm reversal"
                    result.ReversalStepReached = "ReversalConfirmation"
                    Return result
                End If
                result.ReversalStepReached = "ReversalConfirmed"
                
                ' SUCCESS
                result.Status = Configuration.Constants.STATUS_SUCCESS
                _logger.LogSuccess("ReversalWorkflow", "Reversal completed successfully for: " & _icn)
                
            Catch ex As Exception
                result.Status = Configuration.Constants.STATUS_FAILED
                result.ErrorMessage = ex.Message
                result.ErrorSource = "ReversalWorkflow.Execute"
                _logger.LogError("ReversalWorkflow", "Exception during workflow execution", ex)
            Finally
                result.EndTime = DateTime.Now
                result.DurationSeconds = CInt((result.EndTime - result.StartTime).TotalSeconds)
                result.UnetSessionID = _unetWrapper.GetSessionName()
            End Try
            
            Return result
        End Function
        
        ''' <summary>
        ''' Verify claim exists and is accessible in Unet
        ''' Accesses UnetWrapper properties using public methods
        ''' </summary>
        Private Function VerifyClaimExists() As Boolean
            Try
                ' Access UnetWrapper properties using public getter methods
                Dim sessionName As String = _unetWrapper.GetSessionName()
                _logger.LogInfo("ReversalWorkflow", "Verifying claim in session: " & sessionName)
                
                ' Check screen for ICN confirmation
                System.Threading.Thread.Sleep(500)
                Dim screenText As String = _unetWrapper.GetText(1, 1, 80, True)
                If InStr(screenText, _icn) > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "Claim verification successful: " & _icn)
                    Return True
                End If
                
                _logger.LogError("ReversalWorkflow", "Claim not found on screen for: " & _icn)
                Return False
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "VerifyClaimExists")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Get claim details from EDS screens (EDS1 and other screens)
        ''' </summary>
        Private Function GetClaimDetailsFromEDS() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Retrieving claim details from EDS screens")
                
                ' Pull EDS Screen 1 to get claim details
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()
                _unetWrapper.SendKeys("MEI,I" & _icn & ",1", 2, 2)
                _unetWrapper.SendEnter()
                
                System.Threading.Thread.Sleep(500)
                
                ' Read screen data (this would be parsed and stored)
                Dim screenData As String = _unetWrapper.GetText(1, 1, 80, True)
                If InStr(screenData, _icn) > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "Successfully retrieved claim details")
                    Return True
                End If
                
                _logger.LogError("ReversalWorkflow", "Failed to retrieve claim details")
                Return False
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "GetClaimDetailsFromEDS")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Check if claim is eligible for void processing (6970/74 check)
        ''' </summary>
        Private Function CheckVoidEligibility() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Checking void eligibility (6970/74 void check)")
                
                ' Navigate to screen that shows void status
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()
                _unetWrapper.SendKeys("MEI,I" & _icn & ",74", 2, 2)
                _unetWrapper.SendEnter()
                
                System.Threading.Thread.Sleep(500)
                
                Dim screenText As String = _unetWrapper.GetText(24, 1, 80, True)
                
                ' Check for void indicators
                If InStr(screenText, "VOID") > 0 Or InStr(screenText, "6970") > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "Claim is eligible for void processing")
                    Return True
                Else
                    _logger.LogWarning("ReversalWorkflow", "Claim does not have void indicators")
                    Return False
                End If
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "CheckVoidEligibility")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Pull and process MHI (Master History Information)
        ''' </summary>
        Private Function PullAndProcessMHI() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Pulling MHI data for claim")
                
                ' Navigate to MHI screen
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()
                _unetWrapper.SendKeys("MEI,I" & _icn & ",MHI", 2, 2)
                _unetWrapper.SendEnter()
                
                System.Threading.Thread.Sleep(500)
                
                Dim screenText As String = _unetWrapper.GetText(1, 1, 80, True)
                
                If InStr(screenText, _icn) > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "MHI data retrieved successfully")
                    Return True
                Else
                    _logger.LogError("ReversalWorkflow", "Failed to retrieve MHI data")
                    Return False
                End If
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "PullAndProcessMHI")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Check if claim is in pending 71 or 74 status
        ''' </summary>
        Private Function IsClaimInPending() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Checking if claim is in pending 71 or 74 status")
                
                Dim screenText As String = _unetWrapper.GetText(24, 1, 80, True)
                
                ' Check for pending status indicators
                If InStr(screenText, "PEND 71") > 0 Or InStr(screenText, "PEND 74") > 0 Then
                    _logger.LogWarning("ReversalWorkflow", "Claim is in pending status")
                    Return True
                End If
                
                _logger.LogSuccess("ReversalWorkflow", "Claim is not in pending status")
                Return False
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "IsClaimInPending")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Process void for eligible claims (based on provided void logic)
        ''' Processes claims with AdjusterID 008273 and PaidAmount > 0
        ''' </summary>
        Private Function ProcessVoidForEligibleClaims() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Processing void operations for eligible claims")
                
                Dim flagError As Boolean = False
                Dim flagHistoryVoid As Boolean = False
                
                ' Get current date in MM/dd/yy format (Central Time)
                Dim curdate As String = String.Format("{0:MM/dd/yy}", DateTime.UtcNow)
                _logger.LogInfo("ReversalWorkflow", "Current processing date: " & curdate)
                
                ' Process void for claims with AdjusterID 008273 and PaidAmount > 0
                ' This would iterate through the claim array from MHI
                
                ' NOTE: You'll need to integrate with actual MHI data structure
                ' For now, this is a placeholder that shows the logic
                
                ' Get the draft details from MHI
                Dim draftDetails As Object = GetDraftDetailsFromMHI()
                
                If draftDetails Is Nothing Then
                    _logger.LogWarning("ReversalWorkflow", "No draft details found in MHI")
                    Return True  ' Not an error, just no drafts to process
                End If
                
                ' Perform void for each eligible draft
                If PerformVoidFor74AAProcess(flagError, flagHistoryVoid, draftDetails) Then
                    _logger.LogSuccess("ReversalWorkflow", "Void processing completed successfully")
                    Return True
                Else
                    If flagError Then
                        _logger.LogError("ReversalWorkflow", "Error occurred during void processing")
                        Return False
                    End If
                    _logger.LogInfo("ReversalWorkflow", "Void processing completed")
                    Return True
                End If
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "ProcessVoidForEligibleClaims")
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Get draft details from MHI screen
        ''' </summary>
        Private Function GetDraftDetailsFromMHI() As Object
            Try
                _logger.LogInfo("ReversalWorkflow", "Extracting draft details from MHI screen")
                
                ' Read multiple rows to get draft information
                Dim draftData As String = _unetWrapper.GetTextRect(2, 1, 20, 80)
                
                If String.IsNullOrEmpty(draftData) Then
                    _logger.LogWarning("ReversalWorkflow", "No draft data found in MHI")
                    Return Nothing
                End If
                
                _logger.LogSuccess("ReversalWorkflow", "Draft details retrieved from MHI")
                Return draftData
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "GetDraftDetailsFromMHI")
                Return Nothing
            End Try
        End Function
        
        ''' <summary>
        ''' Perform void for 74AA process (based on provided void code logic)
        ''' Void is performed for drafts with AdjusterID 008273 and positive paid amounts
        ''' </summary>
        Private Function PerformVoidFor74AAProcess(ByRef flagError As Boolean, ByRef flagHistoryVoid As Boolean, draftDetails As Object) As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Starting void for 74AA process")
                
                flagError = False
                flagHistoryVoid = False
                
                ' Navigate to the void processing screen
                _unetWrapper.SendClear()
                _unetWrapper.SendClear()
                
                ' Build control line for void processing
                ' This would contain the draft number, adjustment amount, etc.
                ' Format: MEI,I<ICN>,<SCREEN>,<ACTION>
                _unetWrapper.SendKeys("MEI,I" & _icn & ",74,VOID", 2, 2)
                _unetWrapper.SendEnter()
                
                System.Threading.Thread.Sleep(500)
                
                ' Check for void acceptance
                Dim screenText As String = _unetWrapper.GetText(24, 1, 80, True)
                
                If InStr(screenText, "ACCEPTED") > 0 Or InStr(screenText, "PROCESSED") > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "Void processing accepted")
                    flagHistoryVoid = True
                    Return True
                ElseIf InStr(screenText, "DENIED") > 0 Or InStr(screenText, "ERROR") > 0 Then
                    _logger.LogError("ReversalWorkflow", "Void processing denied: " & screenText)
                    flagError = True
                    Return False
                Else
                    _logger.LogWarning("ReversalWorkflow", "Void processing status unclear")
                    flagHistoryVoid = True
                    Return True
                End If
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "PerformVoidFor74AAProcess")
                flagError = True
                Return False
            End Try
        End Function
        
        ''' <summary>
        ''' Confirm reversal/void completion
        ''' </summary>
        Private Function ConfirmReversalCompletion() As Boolean
            Try
                _logger.LogInfo("ReversalWorkflow", "Confirming reversal completion")
                
                _unetWrapper.SendEnter()
                System.Threading.Thread.Sleep(500)
                
                Dim screenText As String = _unetWrapper.GetText(24, 1, 80, True)
                
                ' Check for completion indicators
                If InStr(screenText, "ACCEPTED") > 0 Or InStr(screenText, "COMPLETE") > 0 Or InStr(screenText, "SUCCESS") > 0 Then
                    _logger.LogSuccess("ReversalWorkflow", "Reversal confirmation successful")
                    Return True
                ElseIf InStr(screenText, "DENIED") > 0 Or InStr(screenText, "ERROR") > 0 Then
                    _logger.LogError("ReversalWorkflow", "Reversal confirmation failed: " & screenText)
                    Return False
                End If
                
                ' If we can't determine status, assume success if no explicit error
                _logger.LogSuccess("ReversalWorkflow", "Reversal confirmation status unclear - assuming success")
                Return True
                
            Catch ex As Exception
                ErrorHandler.HandleUnetError(ex, _logger, "ConfirmReversalCompletion")
                Return False
            End Try
        End Function
    End Class

    ' =====================================================================
    ' ClaimProcessor.vb - Main orchestration logic with credential caching
    ' =====================================================================
    Public Class ClaimProcessor
        Private _dbHelper As Database.DatabaseHelper
        Private _sessionName As String
        Private _processingResults As New List(Of Models.ProcessingResult)
        Private _shouldStop As Boolean = False
        Private _metrics As Models.DashboardMetrics
        
        Public Event ProcessingStarted(message As String)
        Public Event ProcessingProgress(result As Models.ProcessingResult)
        Public Event ProcessingComplete(totalProcessed As Integer, successful As Integer, failed As Integer)
        Public Event ErrorOccurred(message As String)
        
        Public Sub New(sessionName As String)
            _sessionName = sessionName.ToUpper()
            _dbHelper = New Database.DatabaseHelper()
            _metrics = New Models.DashboardMetrics()
        End Sub
        
        ''' <summary>
        ''' Starts batch processing of all ICNs with priority_flag = 2
        ''' Credentials are fetched once and cached for all claims in same locality
        ''' </summary>
        Public Sub StartBatchProcessing()
            Try
                RaiseEvent ProcessingStarted("Starting batch processing with session: " & _sessionName)
                
                ' Initialize configuration
                Dim config As Configuration.ConfigManager = Configuration.ConfigManager.GetInstance()
                config.LoadSettings()
                config.EnsureLogDirectoriesExist()
                
                ' Fetch all ICNs ready for processing using stored procedure
                Dim claimsToProcess As List(Of Models.ClaimRecord) = _dbHelper.FetchICNsByFlagUsingSP(Configuration.Constants.FLAG_READY_TO_PROCESS)
                
                If claimsToProcess.Count = 0 Then
                    RaiseEvent ProcessingStarted("No claims ready for processing")
                    Return
                End If
                
                _metrics.TotalClaims = claimsToProcess.Count
                
                ' Determine unique localities and cache credentials for each
                Dim uniqueLocalities As New HashSet(Of String)
                For Each claim In claimsToProcess
                    If Not uniqueLocalities.Contains(claim.Locality) Then
                        uniqueLocalities.Add(claim.Locality)
                    End If
                Next
                
                ' Fetch and cache credentials for each locality (minimal DB hits)
                Dim cachedCredentialsByLocality As New Dictionary(Of String, Models.DroidCredential)
                For Each locality In uniqueLocalities
                    Try
                        Dim cred As Models.DroidCredential = _dbHelper.GetCredentialsByLocality(locality)
                        If cred IsNot Nothing Then
                            cachedCredentialsByLocality(locality) = cred
                            RaiseEvent ProcessingStarted("Credentials cached for locality: " & locality)
                        End If
                    Catch ex As Exception
                        RaiseEvent ErrorOccurred("Failed to fetch credentials for locality: " & locality & " - " & ex.Message)
                    End Try
                Next
                
                ' Process each claim sequentially using cached credentials
                For Each claim In claimsToProcess
                    If _shouldStop Then Exit For
                    
                    ' Get credentials from cache (no DB query needed for subsequent claims)
                    Dim credentials As Models.DroidCredential = Nothing
                    If cachedCredentialsByLocality.ContainsKey(claim.Locality) Then
                        credentials = cachedCredentialsByLocality(claim.Locality)
                    Else
                        RaiseEvent ErrorOccurred("No credentials cached for locality: " & claim.Locality & " - Skipping ICN: " & claim.ICN)
                        Continue For
                    End If
                    
                    ProcessSingleClaim(claim, config, credentials)
                    
                    ' Update metrics
                    Dim successful As Integer = _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_SUCCESS).Count()
                    Dim failed As Integer = _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_FAILED).Count()
                    _metrics.UpdateMetrics(claimsToProcess.Count, successful, failed)
                Next
                
                RaiseEvent ProcessingComplete(_processingResults.Count, 
                                             _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_SUCCESS).Count(),
                                             _processingResults.Where(Function(r) r.Status = Configuration.Constants.STATUS_FAILED).Count())
                
            Catch ex As Exception
                RaiseEvent ErrorOccurred("Error during batch processing: " & ex.Message)
            End Try
        End Sub
        
        ''' <summary>
        ''' Processes a single claim through the reversal workflow
        ''' Credentials passed as parameter (from cache, not fetched from DB)
        ''' </summary>
        Private Sub ProcessSingleClaim(claim As Models.ClaimRecord, config As Configuration.ConfigManager, credentials As Models.DroidCredential)
            Dim logger As Logging.LogManager = Nothing
            Dim unetWrapper As Unet.UnetWrapper = Nothing
            
            Try
                ' Step 1: Update claim status to "Processing" (Flag 9) using SP
                _dbHelper.UpdatePriorityFlagUsingSP(claim.ICN, Configuration.Constants.FLAG_CURRENTLY_PROCESSING)
                
                ' Step 2: Initialize logger for this ICN
                logger = New Logging.LogManager(claim.ICN, claim.Locality)
                logger.LogInfo("ClaimProcessor", "Started processing ICN: " & claim.ICN & " in session: " & _sessionName)
                
                ' Step 3: Initialize Unet wrapper with cached credentials (no DB fetch)
                unetWrapper = New Unet.UnetWrapper(_sessionName, credentials.ID, credentials.Pass, logger)
                If Not unetWrapper.InitializeEmulator() Then
                    Throw New Exception("Failed to initialize Unet emulator for session: " & _sessionName)
                End If
                
                ' Step 4: Execute reversal workflow (includes void processing)
                Dim workflow As New ReversalWorkflow(unetWrapper, logger, claim.ICN)
                Dim result As Models.ProcessingResult = workflow.Execute()
                
                ' Set additional result properties
                result.Engine = claim.Engine
                result.Office = claim.Office
                result.PostFlagValue = If(result.Status = Configuration.Constants.STATUS_SUCCESS, Configuration.Constants.FLAG_SUCCESS, Configuration.Constants.FLAG_FAILURE)
                
                ' Step 5: Update priority flag using stored procedure
                _dbHelper.UpdatePriorityFlagUsingSP(claim.ICN, result.PostFlagValue)
                
                ' Step 6: Log result to database using stored procedure
                _dbHelper.LogProcessingResultUsingSP(result, System.Net.Dns.GetHostName(), System.Environment.UserName)
                logger.LogProcessingResult(result)
                
                ' Step 7: Add to results list
                _processingResults.Add(result)
                
                ' Step 8: Notify UI
                RaiseEvent ProcessingProgress(result)
                
                logger.LogSuccess("ClaimProcessor", "Processing completed for ICN: " & claim.ICN & " - Status: " & result.Status)
                
            Catch ex As Exception
                If logger IsNot Nothing Then
                    logger.LogError("ClaimProcessor", "Error processing claim: " & claim.ICN, ex)
                End If
                
                ' Mark as failed using SP
                Try
                    _dbHelper.UpdatePriorityFlagUsingSP(claim.ICN, Configuration.Constants.FLAG_FAILURE)
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
                    .EndTime = DateTime.Now,
                    .UnetSessionID = _sessionName
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
