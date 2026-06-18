Imports System.Configuration
Namespace Configuration
    ' =====================================================================
    ' Constants.vb - Application-wide constants and configuration values
    ' =====================================================================
    Public Module Constants

        ' ========== Table & Column Names ==========
        Public Const TABLE_WCRYPAYMENTLOG As String = "WCRPaymentLog"
        Public Const TABLE_CLAIMREVERSALLOG As String = "ClaimReversalLog"
        Public Const TABLE_DROIDCRED As String = "DroidCred"

        Public Const COL_ICN As String = "ICN"
        Public Const COL_PRIORITY_FLAG As String = "priority_flag"
        Public Const COL_LOCALITY As String = "Locality"
        Public Const COL_ENGINE As String = "Engine"
        Public Const COL_OFFICE As String = "Office"

        ' ========== Priority Flag Values ==========
        Public Const FLAG_READY_TO_PROCESS As Integer = 2
        Public Const FLAG_CURRENTLY_PROCESSING As Integer = 9
        Public Const FLAG_SUCCESS As Integer = 5
        Public Const FLAG_FAILURE As Integer = 6
        Public Const FLAG_PENDING As Integer = 7

        ' ========== Locality Values ==========
        Public Const LOCALITY_GLOBAL As String = "Global"
        Public Const LOCALITY_DOMESTIC As String = "Domestic"

        ' ========== Application Info ==========
        Public Const APP_NAME As String = "Corrected Claim Reversal"
        Public Const APP_TITLE As String = "Corrected Claim Reversal Dashboard"

        ' ========== Status Codes ==========
        Public Const STATUS_PROCESSING As String = "Processing"
        Public Const STATUS_SUCCESS As String = "Success"
        Public Const STATUS_FAILED As String = "Failed"
        Public Const STATUS_PENDING As String = "Pending"

        ' ========== Log Levels ==========
        Public Const LOG_LEVEL_INFO As String = "[INFO ]"
        Public Const LOG_LEVEL_ERROR As String = "[ERROR]"
        Public Const LOG_LEVEL_SUCCESS As String = "[SUCCESS]"
        Public Const LOG_LEVEL_WARNING As String = "[WARN ]"

    End Module

    ' =====================================================================
    ' ConfigManager.vb - Manages application configuration and settings
    ' =====================================================================
    Public Class ConfigManager
        Private Shared _instance As ConfigManager
        Private _primaryConnStr As String
        Private _secondaryConnStr As String

        ' Cache for app settings
        Private _appSettings As New Dictionary(Of String, String)
        Private _isInitialized As Boolean = False

        ''Singleton pattern - Get configManager instance
        Public Shared Function GetInstance() As ConfigManager
            If _instance Is Nothing Then
                _instance = New ConfigManager()
            End If
            Return _instance
        End Function

        ''' <summary>
        ''' Load all settings from App.config
        ''' </summary>
        Public Sub LoadSettings()
            If _isInitialized Then Return
            Try
                ' Load connection strings from App.config 
                _primaryConnStr = ConfigurationManager.ConnectionStrings("PrimaryDB").ConnectionString
                _secondaryConnStr = ConfigurationManager.ConnectionStrings("SecondaryDB").ConnectionString

                If String.IsNullOrEmpty(_primaryConnStr) Then
                    Throw New Exception("PrimaryDB connection string not found in App.config")
                End If
                If String.IsNullOrEmpty(_secondaryConnStr) Then
                    Throw New Exception("SecondaryDB connection string not found in App.config")
                End If

                ' load app settings from App.config
                Dim appSettings = ConfigurationManager.AppSettings

                For Each key In appSettings.AllKeys
                    _appSettings(key) = appSettings(key)
                Next

                _isInitialized = True

            Catch ex As Exception
                Throw New Exception("Failed to load configuration: " & ex.Message)
            End Try
        End Sub

        Public Function GetSetting(key As String, Optional defaultValue As String = "") As String
            If Not _isInitialized Then
                LoadSettings()
            End If

            If _appSettings.ContainsKey(key) Then
                Return _appSettings(key)
            End If

            Return defaultValue
        End Function

        ''' <summary>
        ''' Get primary database connection string
        ''' </summary>
        Public Function GetPrimaryConnectionString() As String
            If Not _isInitialized Then
                LoadSettings()
            End If

            If String.IsNullOrEmpty(_primaryConnStr) Then
                Throw New Exception("Primary connection string is not configured in App.config")
            End If

            Return _primaryConnStr
        End Function

        ''' <summary>
        ''' Get secondary database connection string
        ''' </summary>
        Public Function GetSecondaryConnectionString() As String
            If Not _isInitialized Then
                LoadSettings()
            End If

            If String.IsNullOrEmpty(_secondaryConnStr) Then
                Throw New Exception("Secondary connection string is not configured in App.config")
            End If

            Return _secondaryConnStr
        End Function

        ''' <summary>
        ''' Get database connection string (wrapper for backward compatibility)
        ''' </summary>
        Public Function GetConnectionString(isPrimary As Boolean) As String
            If isPrimary Then
                Return GetPrimaryConnectionString()
            Else
                Return GetSecondaryConnectionString()
            End If
        End Function

        ''' <summary>
        ''' Get log root path from App.config
        ''' </summary>
        Public Function GetLogRootPath() As String
            Return GetSetting("LogRootPath", "\\nas01042pn\Data\WCC_Droid_Drive\Corrected Log\Sourav_Maji\ReversalBackup")
        End Function

        ''' <summary>
        ''' Get log global path
        ''' </summary>
        Public Function GetLogGlobalPath() As String
            Return GetSetting("LogGlobalPath", System.IO.Path.Combine(GetLogRootPath(), "Global"))
        End Function

        ''' <summary>
        ''' Get log domestic path
        ''' </summary>
        Public Function GetLogDomesticPath() As String
            Return GetSetting("LogDomesticPath", System.IO.Path.Combine(GetLogRootPath(), "Domestic"))
        End Function

        ''' <summary>
        ''' Get export path for Excel files
        ''' </summary>
        Public Function GetExportPath() As String
            Return GetSetting("ExportPath", System.IO.Path.Combine(GetLogRootPath(), "Exports"))
        End Function

        ''' <summary>
        ''' Get Unet operation timeout in milliseconds
        ''' </summary>
        Public Function GetUnetTimeoutMs() As Integer
            Dim timeoutStr As String = GetSetting("UnetTimeout", "30000")
            If Integer.TryParse(timeoutStr, timeoutStr) Then
                Return CInt(timeoutStr)
            End If
            Return 30000
        End Function

        ''' <summary>
        ''' Get Unet connection wait time in milliseconds
        ''' </summary>
        Public Function GetUnetConnectionWaitMs() As Integer
            Dim waitStr As String = GetSetting("UnetConnectionWait", "50")
            If Integer.TryParse(waitStr, waitStr) Then
                Return CInt(waitStr)
            End If
            Return 50
        End Function

        ''' <summary>
        ''' Get application version from App.config
        ''' </summary>
        Public Function GetAppVersion() As String
            ' First try to get from App.config, fallback to AssemblyVersion
            Dim version As String = GetSetting("AppVersion", "")
            If Not String.IsNullOrEmpty(version) Then
                Return version
            End If

            ' Fallback to assembly version
            Dim assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Return assemblyVersion.ToString()
        End Function

        ''' <summary>
        ''' Get max batch size for processing
        ''' </summary>
        Public Function GetMaxBatchSize() As Integer
            Dim batchStr As String = GetSetting("MaxBatchSize", "50")
            If Integer.TryParse(batchStr, batchStr) Then
                Return CInt(batchStr)
            End If
            Return 50
        End Function


        Public Function EnsureLogDirectoriesExist() As Boolean
            Try
                Dim globalPath As String = GetLogGlobalPath()
                Dim domesticPath As String = GetLogDomesticPath()
                Dim exportPath As String = GetExportPath()


                If Not System.IO.Directory.Exists(globalPath) Then
                    System.IO.Directory.CreateDirectory(globalPath)
                End If

                If Not System.IO.Directory.Exists(domesticPath) Then
                    System.IO.Directory.CreateDirectory(domesticPath)
                End If

                If Not System.IO.Directory.Exists(exportPath) Then
                    System.IO.Directory.CreateDirectory(exportPath)
                End If

                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error creating log directories: " & ex.Message)
                Return False
            End Try
        End Function

        Public Function GetLogPathForICN(locality As String, icn As String) As String
            ' Returns: C:\ClaimReversalLogs\Global\2025\01\15\AA12345678_20250115_093045.log
            Dim basePath As String = If(locality = Constants.LOCALITY_GLOBAL, GetLogGlobalPath(), GetLogDomesticPath())
            Dim now As DateTime = DateTime.Now
            Dim yearFolder As String = now.Year.ToString()
            Dim monthFolder As String = now.Month.ToString("00")
            Dim dayFolder As String = now.Day.ToString("00")
            Dim timestamp As String = now.ToString("yyyyMMdd_HHmmss")
            Dim filename As String = icn & "_" & timestamp & ".log"

            Dim fullPath As String = System.IO.Path.Combine(basePath, yearFolder, monthFolder, dayFolder)

            ' Ensure directory exists
            If Not System.IO.Directory.Exists(fullPath) Then
                System.IO.Directory.CreateDirectory(fullPath)
            End If

            Return System.IO.Path.Combine(fullPath, filename)
        End Function

    End Class

End Namespace
