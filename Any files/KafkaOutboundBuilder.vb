Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Globalization
Imports Newtonsoft.Json

' Add this file to the VB.NET automation project.
' Reads suffix number (NBR) and suffix code (PC) from mdlVar.mdlVar.CurrentICNEDS9.
' The automation records each suffix's final result as processing happens.

Public Class KafkaSuffixMapItem
    <JsonProperty("sfxCd")>
    Public Property SfxCd As String

    <JsonProperty("sfxNbr")>
    Public Property SfxNbr As String

    <JsonProperty("sfxStatus")>
    Public Property SfxStatus As String
End Class

Public Class KafkaAutomationResponse
    <JsonProperty("automation")>
    Public Property Automation As String

    <JsonProperty("automationStatus")>
    Public Property AutomationStatus As String

    <JsonProperty("claimStatus")>
    Public Property ClaimStatus As String

    <JsonProperty("daysToHold")>
    Public Property DaysToHold As String

    <JsonProperty("inventoryControlNumber")>
    Public Property InventoryControlNumber As String

    <JsonProperty("lastTry")>
    Public Property LastTry As String

    <JsonProperty("mcrEngineResponse")>
    Public Property McrEngineResponse As String

    <JsonProperty("owfId")>
    Public Property OwfId As String

    <JsonProperty("sla")>
    Public Property Sla As String

    <JsonProperty("waitTime")>
    Public Property WaitTime As String

    <JsonProperty("workflowSystem")>
    Public Property WorkflowSystem As String

    <JsonProperty("excludeFurtherOWFProcess")>
    Public Property ExcludeFurtherOWFProcess As String

    <JsonProperty("suffixMap")>
    Public Property SuffixMap As List(Of KafkaSuffixMapItem)
End Class

Public Class KafkaOutboundEnvelope
    <JsonProperty("pxStreamPartition")>
    Public Property PxStreamPartition As String

    <JsonProperty("pxStreamPosition")>
    Public Property PxStreamPosition As String

    <JsonProperty("pxUpdateDateTime")>
    Public Property PxUpdateDateTime As String

    ' The client sample shows this field explicitly as JSON null.
    <JsonProperty("pzDataFlowRecordContext", NullValueHandling:=NullValueHandling.Include)>
    Public Property PzDataFlowRecordContext As Object

    <JsonProperty("UniqueId")>
    Public Property UniqueId As String

    <JsonProperty("AutomationResponse")>
    Public Property AutomationResponse As KafkaAutomationResponse
End Class

Public Module KafkaOutboundBuilder

    ' Key = ICN|suffix number. Safe if more than one ICN is processed concurrently.
    Private ReadOnly _suffixFinalStatus As New ConcurrentDictionary(Of String, String)(
        StringComparer.OrdinalIgnoreCase)

    Private Function NormalizeSuffixNumber(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return String.Empty

        Dim suffixNumber As Integer
        If Integer.TryParse(value.Trim(), suffixNumber) Then
            Return suffixNumber.ToString("00", CultureInfo.InvariantCulture)
        End If

        Return value.Trim().PadLeft(2, "0"c)
    End Function

    Private Function MakeSuffixKey(icn As String, suffixNumber As String) As String
        Return icn.Trim() & "|" & NormalizeSuffixNumber(suffixNumber)
    End Function

    ' Call once when a new ICN starts, before its suffixes are processed.
    Public Sub ResetSuffixStatuses(icn As String)
        If String.IsNullOrWhiteSpace(icn) Then Return

        Dim prefix As String = icn.Trim() & "|"
        For Each existingKey As String In _suffixFinalStatus.Keys
            If existingKey.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                Dim removedValue As String = Nothing
                _suffixFinalStatus.TryRemove(existingKey, removedValue)
            End If
        Next
    End Sub

    ' Record one suffix outcome. Allowed values: FINAL, SKIP, EXCEPTION, HOLD, OPEN.
    Public Sub RecordSuffixStatus(icn As String,
                                  suffixNumber As Integer,
                                  suffixStatus As String)

        If String.IsNullOrWhiteSpace(icn) Then
            Throw New ArgumentException("ICN is required.", NameOf(icn))
        End If

        Dim normalizedStatus As String = If(suffixStatus, String.Empty).Trim().ToUpperInvariant()
        Dim allowedStatuses As String() = {"FINAL", "SKIP", "EXCEPTION", "HOLD", "OPEN"}

        If Array.IndexOf(allowedStatuses, normalizedStatus) < 0 Then
            Throw New ArgumentException("Unsupported suffix status: " & suffixStatus,
                                        NameOf(suffixStatus))
        End If

        Dim key As String = MakeSuffixKey(icn, suffixNumber.ToString(CultureInfo.InvariantCulture))
        _suffixFinalStatus(key) = normalizedStatus
    End Sub

    ' Applies a claim-level status to every suffix currently present on EDS9.
    ' Use for example when the whole ICN is successfully sent to MARS EDITING.
    Public Sub RecordAllSuffixesStatus(icn As String, suffixStatus As String)
        For structureIndex As Integer = mdlVar.CurrentICNEDS9.GetLowerBound(0) To mdlVar.CurrentICNEDS9.GetUpperBound(0)
            Dim suffixNumbers() As String = mdlVar.CurrentICNEDS9(structureIndex).NBR
            If suffixNumbers Is Nothing Then Continue For

            For rowIndex As Integer = suffixNumbers.GetLowerBound(0) To suffixNumbers.GetUpperBound(0)
                Dim suffixText As String = suffixNumbers(rowIndex)
                Dim suffixNumber As Integer

                If Integer.TryParse(If(suffixText, String.Empty).Trim(), suffixNumber) AndAlso suffixNumber > 0 Then
                    RecordSuffixStatus(icn, suffixNumber, suffixStatus)
                End If
            Next
        Next
    End Sub

    ' Builds suffixMap from EDS9: PC = sfxCd, NBR = sfxNbr.
    ' The dictionary supplies the final status; fallbackStatus is used only if
    ' no status was recorded for that suffix.
    Public Function BuildSuffixMap(icn As String,
                                   fallbackStatus As String) As List(Of KafkaSuffixMapItem)

        Dim suffixMap As New List(Of KafkaSuffixMapItem)
        Dim addedSuffixes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For structureIndex As Integer = mdlVar.CurrentICNEDS9.GetLowerBound(0) To mdlVar.CurrentICNEDS9.GetUpperBound(0)

            Dim suffixNumbers() As String = mdlVar.CurrentICNEDS9(structureIndex).NBR
            Dim suffixCodes() As String = mdlVar.CurrentICNEDS9(structureIndex).PC

            If suffixNumbers Is Nothing OrElse suffixCodes Is Nothing Then Continue For

            Dim lastIndex As Integer = Math.Min(suffixNumbers.GetUpperBound(0),
                                                suffixCodes.GetUpperBound(0))

            For rowIndex As Integer = 0 To lastIndex

                Dim suffixNumber As String = NormalizeSuffixNumber(suffixNumbers(rowIndex))
                Dim suffixCode As String = If(suffixCodes(rowIndex), String.Empty).Trim()

                If suffixNumber = String.Empty Then Continue For
                If Not addedSuffixes.Add(suffixNumber) Then Continue For

                Dim recordedStatus As String = Nothing
                Dim key As String = MakeSuffixKey(icn, suffixNumber)

                If Not _suffixFinalStatus.TryGetValue(key, recordedStatus) Then
                    recordedStatus = If(String.IsNullOrWhiteSpace(fallbackStatus),
                                        "OPEN",
                                        fallbackStatus.Trim().ToUpperInvariant())
                End If

                suffixMap.Add(New KafkaSuffixMapItem With {
                    .SfxCd = suffixCode,
                    .SfxNbr = suffixNumber,
                    .SfxStatus = recordedStatus
                })
            Next
        Next

        Return suffixMap
    End Function

    ' Builds the complete outer JSON object containing AutomationResponse.
    ' Empty strings are retained. Optional Nothing-valued fields are omitted.
    Public Function BuildOutboundPayload(
        icn As String,
        owfId As String,
        automationStatus As String,
        claimStatus As String,
        lastTry As Boolean,
        fallbackSuffixStatus As String,
        Optional isMarsEditing As Boolean = False,
        Optional excludeFurtherOWFProcess As String = Nothing,
        Optional pxStreamPartition As String = "",
        Optional pxStreamPosition As String = "",
        Optional uniqueId As String = Nothing,
        Optional automationName As String = "WIPRO",
        Optional workflowSystem As String = "WIPRO",
        Optional daysToHold As String = Nothing,
        Optional mcrEngineResponse As String = Nothing,
        Optional sla As String = Nothing,
        Optional waitTime As String = Nothing) As String

        If String.IsNullOrWhiteSpace(icn) Then
            Throw New ArgumentException("ICN is required.", NameOf(icn))
        End If

        Dim normalizedAutomationStatus As String =
            If(automationStatus, String.Empty).Trim().ToUpperInvariant()
        Dim normalizedClaimStatus As String =
            If(claimStatus, String.Empty).Trim().ToUpperInvariant()

        If String.IsNullOrWhiteSpace(fallbackSuffixStatus) Then
            Throw New ArgumentException(
                "Pass an explicit fallback suffix status (for example FINAL, OPEN, HOLD or SKIP).",
                NameOf(fallbackSuffixStatus))
        End If

        ' Client examples use the string "NA" when daysToHold is not applicable.
        If isMarsEditing Then
            daysToHold = "5"
            fallbackSuffixStatus = "FINAL"
            RecordAllSuffixesStatus(icn, "FINAL")
        ElseIf daysToHold Is Nothing Then
            daysToHold = "NA"
        End If

        ' HOLD and FINAL examples include these fields with "NA".
        ' OPEN/OK and EXCEPTION examples omit them.
        Dim isOpenOk As Boolean =
            normalizedAutomationStatus = "OK" AndAlso normalizedClaimStatus = "OPEN"
        Dim isException As Boolean = normalizedAutomationStatus = "EXCEPTION"

        If Not isOpenOk AndAlso Not isException Then
            If mcrEngineResponse Is Nothing Then mcrEngineResponse = "NA"
            If sla Is Nothing Then sla = "NA"
            If waitTime Is Nothing Then waitTime = "NA"
        End If

        Dim nowUtc As DateTime = DateTime.UtcNow
        Dim timestamp As String =
            nowUtc.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture)

        If String.IsNullOrWhiteSpace(uniqueId) Then
            uniqueId = New DateTimeOffset(nowUtc).ToUnixTimeMilliseconds().ToString(
                CultureInfo.InvariantCulture) & "_" & icn.Trim()
        End If

        Dim response As New KafkaAutomationResponse With {
            .Automation = If(automationName, String.Empty),
            .AutomationStatus = normalizedAutomationStatus,
            .ClaimStatus = normalizedClaimStatus,
            .DaysToHold = daysToHold,
            .InventoryControlNumber = icn.Trim(),
            .LastTry = lastTry.ToString().ToLowerInvariant(),
            .McrEngineResponse = mcrEngineResponse,
            .OwfId = If(owfId, String.Empty),
            .Sla = sla,
            .WaitTime = waitTime,
            .WorkflowSystem = If(workflowSystem, String.Empty),
            .ExcludeFurtherOWFProcess = excludeFurtherOWFProcess,
            .SuffixMap = BuildSuffixMap(icn, fallbackSuffixStatus)
        }

        Dim envelope As New KafkaOutboundEnvelope With {
            .PxStreamPartition = If(pxStreamPartition, String.Empty),
            .PxStreamPosition = If(pxStreamPosition, String.Empty),
            .PxUpdateDateTime = timestamp,
            .PzDataFlowRecordContext = Nothing,
            .UniqueId = uniqueId,
            .AutomationResponse = response
        }

        Dim settings As New JsonSerializerSettings With {
            .NullValueHandling = NullValueHandling.Ignore,
            .Formatting = Formatting.Indented
        }

        Return JsonConvert.SerializeObject(envelope, settings)
    End Function

End Module
