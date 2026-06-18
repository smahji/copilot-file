Imports System.Data.SqlClient
Imports System.Reflection.MethodBase
Namespace Database
    ' =====================================================================
    ' ConnectionManager.vb - Manages database connections from App.config
    ' =====================================================================
    Public Class ConnectionManager
        Private _configManager As Configuration.ConfigManager

        Public Sub New()
            _configManager = Configuration.ConfigManager.GetInstance()
            _configManager.LoadSettings()
        End Sub

        ''' <summary>
        ''' Get connection to primary database (EIAutomationTeamo2)
        ''' </summary>
        Public Function GetPrimaryConnection() As SqlConnection
            Try
                Dim connStr As String = _configManager.GetPrimaryConnectionString()
                Return New SqlConnection(connStr)
            Catch ex As Exception
                Throw New Exception("Failed to create primary database connection: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Get connection to secondary database (Wipro_Corrected_Droid)
        ''' </summary>
        Public Function GetSecondaryConnection() As SqlConnection
            Try
                Dim connStr As String = _configManager.GetSecondaryConnectionString()
                Return New SqlConnection(connStr)
            Catch ex As Exception
                Throw New Exception("Failed to create secondary database connection: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Tests the connection to the primary database
        ''' </summary>
        Public Function TestPrimaryConnection() As Boolean
            Try
                Using conn As SqlConnection = GetPrimaryConnection()
                    conn.Open()
                    If conn.State = System.Data.ConnectionState.Open Then
                        conn.Close()
                        Return True
                    End If
                End Using
            Catch ex As Exception
                Return False
            End Try
            Return False
        End Function

        ''' <summary>
        ''' Tests the connection to the secondary database
        ''' </summary>
        Public Function TestSecondaryConnection() As Boolean
            Try
                Using conn As SqlConnection = GetSecondaryConnection()
                    conn.Open()
                    If conn.State = System.Data.ConnectionState.Open Then
                        conn.Close()
                        Return True
                    End If
                End Using
            Catch ex As Exception
                Return False
            End Try
            Return False
        End Function
    End Class

    ' =====================================================================
    ' DatabaseHelper.vb - Core database operations using stored procedures
    ' =====================================================================
    Public Class DatabaseHelper
        Private _connManager As ConnectionManager

        Public Sub New()
            _connManager = New ConnectionManager()
        End Sub

        ''' <summary>
        ''' Fetches all ICNs with priority_flag = 2 (Ready to process) using stored procedure
        ''' </summary>
        Public Function FetchICNsByFlag(flag As Integer) As List(Of Models.ClaimRecord)
            Dim records As New List(Of Models.ClaimRecord)

            Try
                Using conn As SqlConnection = _connManager.GetPrimaryConnection()
                    Using cmd As New SqlCommand("sp_GetReadyICNs", conn)
                        cmd.CommandType = System.Data.CommandType.StoredProcedure
                        cmd.Parameters.AddWithValue("@MaxRecords", 100)
                        cmd.CommandTimeout = 30

                        conn.Open()
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            While reader.Read()
                                Dim record As New Models.ClaimRecord With {
                                    .ID = reader("Id"),
                                    .ICN = reader("ICN").ToString().Trim,
                                    .Engine = reader("Engine").ToString().Trim,
                                    .Office = reader("Office").ToString().Trim,
                                    .Locality = reader("Locality").ToString().Trim,
                                    .DateResearch = If(IsDBNull(reader("DateResearch")), DateTime.MinValue, CDate(reader("DateResearch"))),
                                    .ResearchComment = reader("ResearchComment").ToString(),
                                    .PriorityFlag = CInt(reader("priority_flag")),
                                    .CreatedDate = DateTime.Now
                                }
                                records.Add(record)
                            End While
                        End Using
                    End Using
                End Using

            Catch ex As Exception
                Throw New Exception("Error fetching ICNs by flag using stored procedure: " & ex.Message, ex)
            End Try

            Return records
        End Function

        ''' <summary>
        ''' Fetches all credentials from DroidCred table
        ''' Returns Dictionary where key is Locality and value is DroidCredential
        ''' Call this ONCE at startup and cache in ClaimProcessor
        ''' </summary>
        Public Function FetchAllCredentials() As Dictionary(Of String, Models.DroidCredential)
            Dim credentialsDictionary As New Dictionary(Of String, Models.DroidCredential)

            Try
                Dim sql As String = String.Format(
                    "SELECT ID, Pass, Locality FROM {0}",
                    Configuration.Constants.TABLE_DROIDCRED
                )

                Using conn As SqlConnection = _connManager.GetSecondaryConnection()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 30

                        conn.Open()
                        Using reader As SqlDataReader = cmd.ExecuteReader()
                            While reader.Read()
                                Dim credential As New Models.DroidCredential With {
                                    .ID = reader("ID").ToString().Trim(),
                                    .Pass = reader("Pass").ToString().Trim(),
                                    .Locality = reader("Locality").ToString().Trim()
                                }

                                ' Add to dictionary if not already exists (avoid duplicates)
                                If Not credentialsDictionary.ContainsKey(credential.Locality) Then
                                    credentialsDictionary.Add(credential.Locality.Trim(), credential)
                                End If
                            End While
                        End Using
                    End Using
                End Using

            Catch ex As Exception
                Throw New Exception("Error fetching all credentials: " & ex.Message, ex)
            End Try

            Return credentialsDictionary
        End Function

        ''' <summary>
        ''' Get credential from cached dictionary (PREFERRED - no DB call)
        ''' </summary>
        Public Function GetCredentialFromCache(cachedCredentials As Dictionary(Of String, Models.DroidCredential), locality As String) As Models.DroidCredential
            Try
                If cachedCredentials.ContainsKey(locality) Then
                    Return cachedCredentials(locality)
                End If

                Return Nothing
            Catch ex As Exception
                Throw New Exception("Error getting credential from cache for locality " & locality & ": " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Updates the priority_flag for a specific ICN using stored procedure
        ''' </summary>
        Public Function UpdatePriorityFlag(icn As String, newFlag As Integer, id As Integer) As Boolean
            WriteLog($"Entering {GetCurrentMethod.Name}")
            Try
                Using conn As SqlConnection = _connManager.GetPrimaryConnection()
                    Using cmd As New SqlCommand("sp_UpdateClaimReversalStatus", conn)
                        cmd.CommandType = System.Data.CommandType.StoredProcedure
                        cmd.Parameters.AddWithValue("@ICN", icn)
                        cmd.Parameters.AddWithValue("@ID", id)
                        cmd.Parameters.AddWithValue("@NewStatus", newFlag)
                        cmd.CommandTimeout = 30

                        conn.Open()
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                        Return rowsAffected > 0
                    End Using
                End Using

            Catch ex As Exception
                Throw New Exception("Error updating priority flag for ICN " & icn & ": " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Inserts a processing result into ClaimReversalLog table using stored procedure
        ''' </summary>
        Public Function LogProcessingResult(result As Models.ProcessingResult, serverName As String, processedByUser As String) As Boolean
            Try
                ' Get app version from ConfigManager
                Dim config As Configuration.ConfigManager = Configuration.ConfigManager.GetInstance()
                Dim appVersion As String = config.GetAppVersion()

                Using conn As SqlConnection = _connManager.GetPrimaryConnection()
                    Using cmd As New SqlCommand("sp_LogProcessingResult", conn)
                        cmd.CommandType = System.Data.CommandType.StoredProcedure
                        cmd.Parameters.AddWithValue("@icn", result.ICN)
                        cmd.Parameters.AddWithValue("@engine", If(String.IsNullOrEmpty(result.Engine), DBNull.Value, result.Engine))
                        cmd.Parameters.AddWithValue("@office", If(String.IsNullOrEmpty(result.Office), DBNull.Value, result.Office))
                        cmd.Parameters.AddWithValue("@DateResearch", result.DateResearch)
                        cmd.Parameters.AddWithValue("@ResearchComment", result.ResearchComment)
                        cmd.Parameters.AddWithValue("@ServerName", serverName)
                        cmd.Parameters.AddWithValue("@ProcessStartTime", result.StartTime)
                        cmd.Parameters.AddWithValue("@ProcessEndTime", result.EndTime)
                        cmd.Parameters.AddWithValue("@DurationSeconds", result.DurationSeconds)
                        cmd.Parameters.AddWithValue("@ProcessingStatus", result.Status)
                        cmd.Parameters.AddWithValue("@UnetSessionID", If(String.IsNullOrEmpty(result.UnetSessionID), DBNull.Value, result.UnetSessionID))
                        cmd.Parameters.AddWithValue("@ReversalStepReached", If(String.IsNullOrEmpty(result.ReversalStepReached), DBNull.Value, result.ReversalStepReached))
                        cmd.Parameters.AddWithValue("@ErrorMessage", If(String.IsNullOrEmpty(result.ErrorMessage), DBNull.Value, result.ErrorMessage))
                        cmd.Parameters.AddWithValue("@ErrorSource", If(String.IsNullOrEmpty(result.ErrorSource), DBNull.Value, result.ErrorSource))
                        cmd.Parameters.AddWithValue("@PostFlagValue", result.PostFlagValue)
                        cmd.Parameters.AddWithValue("@ProcessedByUser", processedByUser)
                        cmd.Parameters.AddWithValue("@AppVersion", appVersion)
                        cmd.CommandTimeout = 30

                        conn.Open()
                        Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                        Return rowsAffected > 0
                    End Using
                End Using

            Catch ex As Exception
                Throw New Exception("Error logging processing result for ICN " & result.ICN & ": " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Retrieves all logs from ClaimReversalLog table
        ''' </summary>
        Public Function GetAllLogs() As DataTable
            Try
                Dim sql As String = String.Format(
                    "SELECT * FROM {0} ORDER BY CreatedDate DESC",
                    Configuration.Constants.TABLE_CLAIMREVERSALLOG
                )

                Dim dt As New DataTable()
                Using conn As SqlConnection = _connManager.GetPrimaryConnection()
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.CommandTimeout = 30
                        Dim adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using

                Return dt
            Catch ex As Exception
                Throw New Exception("Error retrieving logs: " & ex.Message, ex)
            End Try
        End Function

    End Class

End Namespace
