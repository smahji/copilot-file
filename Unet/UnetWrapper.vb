
Namespace Unet
    ' =====================================================================
    ' UnetSession.vb - Manages a Unet session state
    ' =====================================================================
    Public Class UnetSession
        Public Property SessionName As String
        Public Property UnetID As String
        Public Property UnetPassword As String
        Public Property IsActive As Boolean = False
        Public Property StartTime As DateTime
        Public Property LastActivityTime As DateTime

        Public Sub New(sessionName As String, unetID As String, unetPassword As String)
            Me.SessionName = sessionName.ToUpper()
            Me.UnetID = unetID
            Me.UnetPassword = unetPassword
            Me.StartTime = DateTime.Now
            Me.LastActivityTime = DateTime.Now
        End Sub

        Public Function IsSessionAlive() As Boolean
            ' Check if session has been inactive for more than 5 minutes
            Dim inactiveSeconds As Double = (DateTime.Now - LastActivityTime).TotalSeconds
            Return inactiveSeconds < 300
        End Function
    End Class

    ' =====================================================================
    ' ScreenReader.vb - Utilities for reading and parsing screen data
    ' =====================================================================
    Public Class ScreenReader
        ''' <summary>
        ''' Extracts text from a specific row and column position on the Unet screen
        ''' </summary>
        Public Shared Function GetScreenText(row As Integer, col As Integer, length As Integer, Optional isTrim As Boolean = False) As String
            Try
                ' This calls into mdlUnet.GetText function
                Dim text As String = autECLPSObj.GetText(row, col, length)
                If isTrim Then
                    text = Trim(text)
                End If
                Return text
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Extracts a rectangular region of screen text
        ''' </summary>
        Public Shared Function GetScreenRect(startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer) As String
            Try
                ' This calls into mdlUnet.GetTextRect function
                Return GetTextRect(startRow, startCol, endRow, endCol)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Parses screen text to extract ICN from error message
        ''' </summary>
        Public Shared Function ExtractICNFromError(errorText As String) As String
            ' Look for ICN pattern: AA12345678 (two letters + 8 digits)
            Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(errorText, "[A-Z]{2}\d{8}")
            If match.Success Then
                Return match.Value
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Checks if a specific error message appears on screen
        ''' </summary>
        Public Shared Function IsErrorOnScreen(errorKeyword As String) As Boolean
            Try
                Dim screenText As String = GetScreenText(24, 1, 80, True)
                Return InStr(screenText, errorKeyword) > 0
            Catch
                Return False
            End Try
        End Function
    End Class

    ' =====================================================================
    ' UnetWrapper.vb - Wraps UnetHelper.dll interactions
    ' =====================================================================
    Public Class UnetWrapper
        Private _session As UnetSession
        Private _logger As Logging.LogManager

        ''' <summary>
        ''' Public properties to expose session data to ReversalWorkflow
        ''' </summary>
        Public ReadOnly Property SessionName As String
            Get
                Return _session.SessionName
            End Get
        End Property

        Public ReadOnly Property UnetID As String
            Get
                Return _session.UnetID
            End Get
        End Property

        Public ReadOnly Property UnetPassword As String
            Get
                Return _session.UnetPassword
            End Get
        End Property

        Public ReadOnly Property IsSessionActive As Boolean
            Get
                Return _session.IsActive
            End Get
        End Property

        Public Sub New(sessionName As String, unetID As String, unetPassword As String, logger As Logging.LogManager)
            _session = New UnetSession(sessionName, unetID, unetPassword)
            _logger = logger
        End Sub

        ''' <summary>
        ''' Initializes the Unet emulator session
        ''' </summary>
        Public Function InitializeEmulator() As Boolean
            Try
                _logger.LogInfo("UnetWrapper", "Initializing emulator session: " & _session.SessionName)

                ' Call mdlUnet.InitializeEmulator
                mdlUnet.InitializeEmulator(_session.SessionName)

                ' Verify session is active by checking communication
                If Not WaitForCommunication() Then
                    _logger.LogError("UnetWrapper", "Failed to establish communication with emulator")
                    Return False
                End If

                _session.IsActive = True
                _logger.LogSuccess("UnetWrapper", "Emulator initialized successfully - Session: " & _session.SessionName)
                Return True

            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Failed to initialize emulator", ex)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Waits for Unet communication to start
        ''' </summary>
        Private Function WaitForCommunication() As Boolean
            Try
                Dim maxRetries As Integer = 10
                Dim retryCount As Integer = 0

                ' This mirrors the user's code:
                ' Do
                '     System.Threading.Thread.Sleep(50)
                ' Loop Until autECLPSObj.CommStarted

                mdlUnet.apiChk()
                Return True

            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Pulls a claim using EDS screen navigation
        ''' </summary>
        Public Function PullClaimEDS(icn As String, screen As String) As Boolean
            Try
                _logger.LogInfo("UnetWrapper", String.Format("Pulling claim EDS - ICN: {0}, Screen: {1}", icn, screen))
                _session.LastActivityTime = DateTime.Now

                ' Call mdlUnet.PullEDS2
                Dim result As Boolean = mdlUnet.PullEDS2(icn, screen)

                If result Then
                    _logger.LogSuccess("UnetWrapper", "Claim EDS pulled successfully: " & icn)
                    Return True
                Else
                    _logger.LogError("UnetWrapper", "Failed to pull claim EDS for: " & icn)
                    Return False
                End If

            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error pulling claim EDS", ex)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Sends keys to the Unet screen
        ''' </summary>
        Public Sub SendKeys(keys As String, Optional row As Integer = 2, Optional col As Integer = 2)
            Try
                mdlUnet.SendKeyes(keys, row, col)
                _session.LastActivityTime = DateTime.Now
                System.Threading.Thread.Sleep(100) ' Brief pause
                mdlUnet.apiChk()
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error sending keys: " & keys, ex)
            End Try
        End Sub

        ''' <summary>
        ''' Sends ENTER key
        ''' </summary>
        Public Sub SendEnter()
            Try
                mdlUnet.SendEnter()
                _session.LastActivityTime = DateTime.Now
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error sending ENTER key", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Sends CLEAR key
        ''' </summary>
        Public Sub SendClear()
            Try
                mdlUnet.SendClear()
                _session.LastActivityTime = DateTime.Now
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error sending CLEAR key", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Sends F8 key (page down)
        ''' </summary>
        Public Sub SendF8()
            Try
                autECLPSObj.SendKeys("[PF8]")
                mdlUnet.apiChk()
                _session.LastActivityTime = DateTime.Now
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error sending F8 key", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Sends F9 key
        ''' </summary>
        Public Sub SendF9()
            Try
                autECLPSObj.SendKeys("[PF9]")
                mdlUnet.apiChk()
                _session.LastActivityTime = DateTime.Now
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error sending F9 key", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Gets text from the screen at specific position
        ''' </summary>
        Public Function GetText(row As Integer, col As Integer, length As Integer, Optional isTrim As Boolean = False) As String
            Try
                Dim text As String = autECLPSObj.GetText(row, col, length)
                If isTrim Then
                    text = Trim(text)
                End If
                _session.LastActivityTime = DateTime.Now
                Return text
            Catch ex As Exception
                _logger.LogError("UnetWrapper", String.Format("Error reading text at R:{0},C:{1}", row, col), ex)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Gets a rectangle of screen text
        ''' </summary>
        Public Function GetTextRect(startRow As Integer, startCol As Integer, endRow As Integer, endCol As Integer) As String
            Try
                Dim text As String = mdlUnet.GetTextRect(startRow, startCol, endRow, endCol)
                _session.LastActivityTime = DateTime.Now
                Return text
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error reading text rectangle", ex)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Checks if a specific message appears on the screen
        ''' </summary>
        Public Function CheckScreenMessage(messageKeyword As String) As Boolean
            Try
                Dim screenText As String = GetText(24, 1, 80, True)
                Return InStr(screenText, messageKeyword) > 0
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error checking screen message", ex)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Closes the Unet session
        ''' </summary>
        Public Sub CloseSession()
            Try
                SendClear()
                SendClear()
                _session.IsActive = False
                _logger.LogInfo("UnetWrapper", "Session closed: " & _session.SessionName)
            Catch ex As Exception
                _logger.LogError("UnetWrapper", "Error closing session", ex)
            End Try
        End Sub

    End Class

End Namespace
