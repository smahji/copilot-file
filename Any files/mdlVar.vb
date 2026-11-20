Public Module mdlVar
    Public Const strConnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\nas01042pn\data\wcc_droid_drive\Corrected DB\NewCorrectedDatabase.accdb;Jet OLEDB:Database Password='meghdut2018'"
    Public Const IngenixRemarkCodes As String = "QH,EB,DN,DO,HU,HV,HW,HS,HT,SD,SS,G1,HK,HP,AR,QI,S3,HQ,AD,AE,UJ,AS,AU,OO,OS,OV,46,47,3I,4A,S0,FA,LV,LW,HX,HY,OI,OK,OA,OF,AF,AG,AW,AX,OW,OX,OY,OZ,8D,2T,TF,UP,EW"
    Public HardCopyText As String, flagExit As Boolean, is30LineToBeused As Boolean, isPEGA As Boolean, FlagTest As Boolean, FlagCARemarkCode As Boolean
    Public FlagScreenNotValid As Boolean = False
    Public Const ReversalModifierList As String = "59,76,91,92,XE,XS,XU,E1,E2,F1,F2,F3,F4,FA,T1,T2,T3,T4,TA,LC,LD,LM,LT,E3,E4,F5,F6,F7,F8,F9,T5,T6,T7,T8,T9,RC,RI,RT"
    Public strPendO2SuffixList As String
    Public Const AllBusinessPartner1 As String = "49,52,53,57,58,71,AH,AJ,AM,AR,AV,B4,B5,B6,B8,B9,BA,BE,BG,BR,BT,BW,BX,BY,BZ,C0,C1,CH"
    Public Const AllBusinessPartner2 As String = "CK,CM,CP,CZ,DJ,DK,DP,DW,DY,DX,DZ,EQ,GM,GS,HG,HJ,HT,HS,HZ,NR,NS,NT,NY,NZ,OP,OZ,PF,PK"
    Public Const AllBusinessPartner3 As String = "PM,CV,PQ,PW,PX,PZ,R2,R1,R4,R6,RA,RV,RX,RZ,S1,S5,SM,SS,SX,SZ,T1,T2,TE,TL,TM,TO,MO,MJ,NK"
    Public Const AllBusinessPartner4 As String = "O1,TH,UO,U0,U3,U4,U6,UA,UB,UC,UE,UH,VH,VX,VZ,WM,Z2,Z5,Z7,Z8,ZE,ZP,ZQ,ZR,ZS,ZT,ZV,ZW,ZX,ZY,ZZ,Z3,OZ,ST,CX,NX,C8,GM"
    Public Const DenialRemarkCodeList As String = "FG,CH,PE,0N,I4,PP,O8,D8,05,13"
    Public SuffixCountOfCurrentICN As Integer
    Public isCurrentClaimIH As Boolean
    Public isCurrentClaimOH As Boolean
    Public icesSuccess As Boolean
    Public flag0762 As Boolean
    Public dt_762Denial As DataTable
    Public allowableFromOCI As String
    Public nonServiceCodePresent As Boolean
    Public billType As String
    Public isC68REV As Boolean
    Public DenialRC() As String
    Public NCopayNBR As String
    Public flagDenial As Boolean
    Public OriginalICN As String
    Public OriginalDraft As String
    Public objConnMgr As Object
    Public FinalHistoryICNList As String
    Public objUNET As Object
    Public historyICN As String
    Public histDraft As String
    Public isNCopayNBR As Boolean
    Public WShell, objShell, objShellWindows, objEDSS As Object
    Public parentedProcess As Process
    Public parentedProcess1 As Process
    Public strEncryptedText As String = ""
    Public strkeyValues As String
    Public emulator1, emulator2
    Public EDSS_URLFound As Boolean = False
    Public OnlyHardCopyOpen As Boolean = False
    Public Const WM_SYSCOMMAND As Integer = 274
    Public Const SC_MAXIMIZE As Integer = 61488
    Public RemarkCode30Line As String
    Public autECLPSObj As Object
    Public autECLConnList As Object
    Public autECLOIAObj As Object
    Public autECLWinMetrics As Object
    Public CurrentClaim(100) As Claim, HistoryClaim(100) As Claim
    Public HistoryClaimCount As Integer, CurrentClaimCount As Integer
    Public strStatus As String
    Public isHistoryICNVoided As Boolean
    Public strComments As String '4
    Public strResearchStatus As String '5
    Public strResearchComments As String '6
    Public strAdjustmentCriteria As String '7
    'Public Const PISLLINE As String = "ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,PLPRE,POD,PRE,REF,SSO,TDPD,BLANK,BLANK-,-----,------,LTR,OI,OIM,OIMEDI,ACPD ,ACPD  ,ACPD--,ACPD-,CFPD ,CFPD  ,CFPD--,COBPD ,COBPD-,COMAC ,COMAC-,EMER ,EMER  ,EMER--,FW   ,FW    ,FW----,FW---,OI   ,OI    ,OI---,OI----,OIM  ,OIM   ,OIM--,OIM---,OIMED"
    Public Const PISLLINE As String = "ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,POD,REF,SSO,TDPD,BLANK,BLANK-,-----,------,LTR,OI,OIM,OIMEDI,ACPD ,ACPD  ,ACPD--,ACPD-,CFPD ,CFPD  ,CFPD--,COBPD ,COBPD-,COMAC ,COMAC-,EMER ,EMER  ,EMER--,FW   ,FW    ,FW----,FW---,OI   ,OI    ,OI---,OI----,OIM  ,OIM   ,OIM--,OIM---,OIMED"
    Public Const InvalidProcCode As String = PISLLINE & ",CXINT,NCOPAY,RX,S9999"
    Public Const ICESList As String = "O8,I4,FE,XS,O5,O6,O7,O9,I3,I5,KM,IA,Q8,1I"
    Public Const ModifierCodes As String = ""
    Public ClaimLevelRemarkCode As String, LineLevelRemarkCode As String, LineLevelRemarkCodePriority2 As String, LineLevel As String
    Public flagProcessed As Boolean = False
    Public FlagMHI As Boolean = False, FlagBillType8 As Boolean = False
    Public StartTime As DateTime = Now, StartDateMonth As String, StartDateLong As String
    Public Const QEDenialCodes As String = "99241,99242,99243,99244,99245,99251,99252,99253,99254,99255,G0406,G0407,G0408,G0425,G0426,G0427,G0508,G0509"
    Public MPRDataSet As DataSet, FlagUndoPerformed As Boolean = False
    Public EDSSText() As String
    Public MPCScreenCount As Integer
    Public IsCOB As Boolean, InvestigateCOBSuffix() As String
    Public EDSSID As String, EDSSPass As String, UnetID As String, UnetPass As String
    Public DXCodeMatchFlag As Boolean
    Public HistoryPaidAmount As String
    Public PaymentComment As String
    Public currentIcnPaymentAmount As Decimal
    Public strDraFT As String
    Public ICESDenialFlag As Boolean
    Public ROVAmount As Decimal
    Public textprevA As String
    Public textprevB As String
    Public IsHDClaim As Boolean
    Public dt_PG_Policyno As DataTable
    Public isModifierPresent As Boolean
    Public LastDatefromHx As String
    Public flag_QC As Boolean
    Public productCode As New Dictionary(Of String, String)

    Public productCodeDict As New Dictionary(Of String, Dictionary(Of String, String))
    Public mrktTypeNDBCodeForPOS As New Dictionary(Of String, String)
    Public mrktTypeNDBCodeForHMO As New Dictionary(Of String, String)
    Public mrktTypeNDBCodeForPPO As New Dictionary(Of String, String)
    Public mrktTypeNDBCodeForEPO As New Dictionary(Of String, String)
    Public mrktTypeNDBCodeForIND As New Dictionary(Of String, String)
    Public mrktTypeNDBCodeForALL As New Dictionary(Of String, String)

    Public Class ServiceLineData
        Public Property Line1 As Dictionary(Of Integer, String)
        Public Property Line2 As Dictionary(Of Integer, String)
    End Class

    Public Sub initializeNDBDict()
        mrktTypeNDBCodeForPOS.Add("50", "C2")
        mrktTypeNDBCodeForPOS.Add("37", "SX")
        mrktTypeNDBCodeForPOS.Add("77", "EL")
        mrktTypeNDBCodeForPOS.Add("63", "A3")
        mrktTypeNDBCodeForPOS.Add("78", "EO")
        mrktTypeNDBCodeForPOS.Add("79", "ER")
        mrktTypeNDBCodeForPOS.Add("42", "V3")
        mrktTypeNDBCodeForPOS.Add("58", "H9")
        mrktTypeNDBCodeForPOS.Add("61", "H2")
        mrktTypeNDBCodeForPOS.Add("30", "ID")
        mrktTypeNDBCodeForPOS.Add("31", "IE")
        mrktTypeNDBCodeForPOS.Add("15", "S7")
        mrktTypeNDBCodeForPOS.Add("70", "I2")
        mrktTypeNDBCodeForPOS.Add("74", "EC")
        mrktTypeNDBCodeForPOS.Add("75", "EF")
        mrktTypeNDBCodeForPOS.Add("90", "L3")
        mrktTypeNDBCodeForPOS.Add("00", "S2")
        mrktTypeNDBCodeForPOS.Add("95", "LI")
        mrktTypeNDBCodeForPOS.Add("76", "EI")

        mrktTypeNDBCodeForHMO.Add("50", "C0")
        mrktTypeNDBCodeForHMO.Add("00", "S0")
        mrktTypeNDBCodeForHMO.Add("69", "SG")
        mrktTypeNDBCodeForHMO.Add("49", "SM")
        mrktTypeNDBCodeForHMO.Add("37", "SZ")
        mrktTypeNDBCodeForHMO.Add("95", "LG")
        mrktTypeNDBCodeForHMO.Add("77", "EJ")
        mrktTypeNDBCodeForHMO.Add("78", "EM")
        mrktTypeNDBCodeForHMO.Add("21", "H7")
        mrktTypeNDBCodeForHMO.Add("22", "H8")
        mrktTypeNDBCodeForHMO.Add("70", "I0")
        mrktTypeNDBCodeForHMO.Add("23", "G0")
        mrktTypeNDBCodeForHMO.Add("15", "S9")
        mrktTypeNDBCodeForHMO.Add("76", "EG")
        mrktTypeNDBCodeForHMO.Add("74", "EA")
        mrktTypeNDBCodeForHMO.Add("75", "ED")
        mrktTypeNDBCodeForHMO.Add("90", "L1")
        mrktTypeNDBCodeForHMO.Add("41", "V1")
        mrktTypeNDBCodeForHMO.Add("16", "SA")
        mrktTypeNDBCodeForHMO.Add("18", "SC")
        mrktTypeNDBCodeForHMO.Add("24", "SE")
        mrktTypeNDBCodeForHMO.Add("48", "SL")
        mrktTypeNDBCodeForHMO.Add("63", "A1")

        mrktTypeNDBCodeForPPO.Add("61", "HB")
        mrktTypeNDBCodeForPPO.Add("57", "H6")
        mrktTypeNDBCodeForPPO.Add("05", "P4")
        mrktTypeNDBCodeForPPO.Add("03", "P2")
        mrktTypeNDBCodeForPPO.Add("70", "I3")
        mrktTypeNDBCodeForPPO.Add("00", "P3")
        mrktTypeNDBCodeForPPO.Add("88", "XN")

        mrktTypeNDBCodeForEPO.Add("37", "SY")
        mrktTypeNDBCodeForEPO.Add("90", "L2")
        mrktTypeNDBCodeForEPO.Add("95", "LH")
        mrktTypeNDBCodeForEPO.Add("77", "EK")
        mrktTypeNDBCodeForEPO.Add("74", "EB")
        mrktTypeNDBCodeForEPO.Add("58", "HD")
        mrktTypeNDBCodeForEPO.Add("32", "IC")
        mrktTypeNDBCodeForEPO.Add("15", "S8")
        mrktTypeNDBCodeForEPO.Add("75", "EE")
        mrktTypeNDBCodeForEPO.Add("63", "A2")
        mrktTypeNDBCodeForEPO.Add("00", "S1")
        mrktTypeNDBCodeForEPO.Add("50", "C1")
        mrktTypeNDBCodeForEPO.Add("76", "EH")

        mrktTypeNDBCodeForIND.Add("00", "N4")

        mrktTypeNDBCodeForALL.Add("50", "AC")
        mrktTypeNDBCodeForALL.Add("00", "AA")

        productCodeDict.Add("POS", mrktTypeNDBCodeForPOS)
        productCodeDict.Add("EPO", mrktTypeNDBCodeForEPO)
        productCodeDict.Add("HMO", mrktTypeNDBCodeForHMO)
        productCodeDict.Add("IND", mrktTypeNDBCodeForIND)
        productCodeDict.Add("PPO", mrktTypeNDBCodeForPPO)
        productCodeDict.Add("ALL", mrktTypeNDBCodeForALL)
    End Sub

    Public Sub initializeProductDictionary()
        productCode.Add("POS", "C2")
        productCode.Add("PS1", "C2")
        productCode.Add("EPO", "C1")
    End Sub


    Public Structure ReconCode
        Dim RemarkCode As String
        Dim ReconCd As String
    End Structure

    Public StructReconCode(100) As ReconCode
    Public Const OnlyPISLLINE As String = "ACPD,CFPD,COBPD,COMAC,EMER,FW,MATDC,PARS,PLPRE,POD,PRE,REF,SSO,TDPD"
    Public isOI As Boolean, isOIM As Boolean
    Public Structure ISIMap
        Dim CurrentClaimProcCode As String
        Dim CurrentClaimBillCharge As String
        Dim CurrentClaimModifier As String
        Dim CurrentClaimDispositionCode As String
        Dim CurrentClaimLineNumber As String
        Dim CurrentClaimflag As String
        Dim CurrentClaimClaimIndex As Integer
        Dim CurrentClaimLineIndex As Integer
        Dim CurrentClaimCommentICN As String
        Dim CurrentClaimCommentProc As String
        Dim CurrentClaimFullComment As String
        Dim HistoryClaimProcCode As String
        Dim HistoryClaimBillCharge As String
        Dim HistoryClaimModifier As String
        Dim HistoryClaimDispositionCode As String
        Dim HistoryClaimLineNumber As String
        Dim HistoryClaimflag As String
        Dim HistoryClaimClaimIndex As Integer
        Dim HistoryClaimLineIndex As Integer
        Dim HistoryClaimCommentICN As String
        Dim HistoryClaimCommentProc As String
        Dim HistoryClaimFullComment As String
    End Structure
    Public CreateMapinISI(100) As ISIMap
    Public Structure ISI
        Dim ProcCode As String
        Dim BillCharge As String
        Dim Modifier As String
        Dim DispositionCode As String
        Dim LineNumber As String
        Dim flag As String
        Dim ClaimIndex As Integer
        Dim LineIndex As Integer
        Dim CommentICN As String
        Dim CommentProc As String
        Dim FullComment As String
    End Structure
    Public ISIScreen(100) As ISI
    Public ISICurrent(100) As ISI
    Public ISICurrentActual(100) As ISI
    Public Structure ClaimFromPaymentScreen
        Dim POS() As String
        Dim ProcCode() As String
        Dim FDOS() As String
        Dim LDOS() As String
        Dim NBR() As String
        Dim Override() As String
        Dim PayeeCode() As String
        Dim RemarkCode() As String
        Dim Charge() As String
        Dim NotCovered() As String
        Dim Allowable() As String
        Dim Deductible() As String
        Dim Coins() As String
        Dim Paid() As String
        Dim Sanction() As String
        Dim Modifier() As String
        Dim EditWarning As String
        Dim CauseCode As String
        Dim Suffix As Integer
        Dim Draft As String
        Dim DateOfProcessing As String
        Dim TotalPaid As String
        Dim HxCindex() As Integer
        Dim HxLIndex() As Integer
    End Structure
    Public CurrentClaimPaymentScreen(100) As ClaimFromPaymentScreen
    Public CurrentClaimPaymentScreenNew(100) As ClaimFromPaymentScreen
    Public HistoryClaimPaymentScreen(100) As ClaimFromPaymentScreen
    Public CurrentPaymentScreenForLog(100) As ClaimFromPaymentScreen


    Public Structure MapCurrentProcDetailsWithHistory
        Dim CurrentClaimProcCode As String
        Dim CurrentClaimFdos As String
        Dim CurrentClaimLdos As String
        Dim CurrentClaimUnit As String
        Dim CurrentClaimCharge As Double
        Dim CurrentClaimNotCovered As Double
        Dim CurrentClaimSuffix As Integer
        Dim CurrentClaimLineNumber As Integer
        Dim CurrentClaimAllowable As String
        Dim CurrentClaimDeductible As String
        Dim CurrentClaimCoins As String
        Dim HistoryClaimProcCode As String
        Dim HistoryClaimFdos As String
        Dim HistoryClaimLdos As String
        Dim HistoryClaimUnit As String
        Dim HistoryClaimCharge As Double
        Dim HistoryClaimNotCovered As Double
        Dim HistoryClaimAllowable As String
        Dim HistoryClaimDeductible As String
        Dim HistoryClaimCoins As String
        Dim HistoryClaimCIndex As Integer
        Dim HistoryClaimLIndex As Integer
        Dim HXICN As String
        Dim IsCodeChange As Boolean
    End Structure
    Public Map(100) As MapCurrentProcDetailsWithHistory
    Public Structure DraftDetails
        Dim Draft As String
        Dim Suffix As String
        Dim Reversal As Boolean
        Dim ReversalType As String
        Dim DateOfProcessing As String
        Dim TotalPaid As Double
        Dim TotalAllowable As Double
        Dim TotalNotCovered As Double
        Dim isClosed As Boolean
        Dim ClosureCode As String
        Dim TotalBillCharge As Double
        Dim DateOfReversal As String
    End Structure
    Public Structure MatchingLines
        Dim ServiceLines As String
    End Structure
    Public Lines(100) As MatchingLines
    Public HistoryClaimDraftDetails(100) As DraftDetails
    Public CurrentClaimDraftDetails(100) As DraftDetails
    Public HistoryClaimDraftDetailsCount As Integer
    Public CurrentClaimDraftDetailsCount As Integer
    Public Structure ModifierList
        Dim POS As String
        Dim ProcedureCode As String
        Dim Modifier As String
        Dim FDOS As String
        Dim LDOS As String
        Dim Suffix As String
        Dim LineNumber As Integer
        Dim Charge As String
        Dim NBR As String
        Dim ServicingPhysician As String
        Dim ReferringPhysician As String
    End Structure
    Public CurrentClaimModifier() As ModifierList
    Public Structure ClaimDetails
        Dim Policy As String
        Dim SSN As String
        Dim PatientName As String
        Dim PatientRelationship As String
        Dim ICN As String
        Dim Tin As String
        Dim FLN As String
        Dim DCC As String
        Dim Fdate As String
        Dim Ldate As String
        Dim Prifix_Tin_Suffix As String
        Dim PatDOB As String
        Dim TotalCharge As String
        Dim Prefix_Tin As String
    End Structure
    Public CurrentClaimDetails As ClaimDetails
    Public HistoryClaimDetails As ClaimDetails
    Public Structure CAStruct
        Dim LineIndex As Integer
        Dim ClaimIndex As Integer
    End Structure
    Public CAline(100) As CAStruct
    Public Structure MRIPat
        Dim Name As String
        Dim Relationship As String
        Dim BirthDate As String
        Dim MDEff As String
        Dim MdCanc As String
        Dim S As String
        Dim C As String
        Dim M As String
        Dim EffDate As String
        Dim CancDate As String
        Dim T As String
        Dim R As String
        Dim Ret_ST As String
        Dim CobDT As String
    End Structure

    Public MyMRI As MRIPat
    'Public Structure MRI
    '    Dim MriRec() As MRIPat
    'End Structure
    Public Structure MedicareDetails
        Dim Indicator As String
        Dim EffectiveDate As String
        Dim CancellationDate As String
    End Structure
    Public Structure OIDetails
        Dim Indicator As String
        Dim EffectiveDate As String
        Dim CancellationDate As String
    End Structure
    Public Structure OIIDetails
        Dim Medicare() As MedicareDetails
        Dim OI() As OIDetails
        Public Sub GetOIIDetails()
            Dim Count As Integer = -1
            Dim Count2 As Integer = -1
            SendF12()
            SendClear()
            SendClear()
            SendClear()
            SendKeyes("OII," + CurrentClaimDetails.Policy + "," + CurrentClaimDetails.SSN + "," + CurrentClaimDetails.PatientName + "," + CurrentClaimDetails.PatientRelationship, 2, 2)
            SendEnter()
            For I As Integer = 19 To 23
                If GetText(I, 2, 27, True).Replace("-", "").Trim <> "" Then
                    If Count = -1 Then
                        Count += 1
                        ReDim Medicare(Count)
                    Else
                        Count += 1
                        ReDim Preserve Medicare(Count)
                    End If
                    Medicare(Count).Indicator = GetText(I, 2, 1)
                    Medicare(Count).EffectiveDate = IIf(GetText(I, 11, 8) = "99999999", "12/31/2099", IIf(GetText(I, 11, 8) = "--------", "01/01/1901", GetText(I, 11, 2) + "/" + GetText(I, 13, 2) + "/" + GetText(I, 15, 4)))
                    Medicare(Count).CancellationDate = IIf(GetText(I, 20, 8) = "99999999", "12/31/2099", IIf(GetText(I, 20, 8) = "--------", "01/01/1901", GetText(I, 20, 2) + "/" + GetText(I, 22, 2) + "/" + GetText(I, 24, 4)))
                End If
                If GetText(I, 29, 27, True).Replace("-", "").Trim <> "" Then
                    If Count2 = -1 Then
                        Count2 += 1
                        ReDim OI(Count2)
                    Else
                        Count2 += 1
                        ReDim Preserve OI(Count2)
                    End If
                    OI(Count2).Indicator = GetText(I, 29, 1)
                    OI(Count2).EffectiveDate = IIf(GetText(I, 34, 8) = "99999999", "12/31/2099", IIf(GetText(I, 34, 8) = "--------", "01/01/1901", GetText(I, 34, 2) + "/" + GetText(I, 36, 2) + "/" + GetText(I, 38, 4)))
                    OI(Count2).CancellationDate = IIf(GetText(I, 46, 8) = "99999999", "12/31/2099", IIf(GetText(I, 46, 8) = "--------", "01/01/1901", GetText(I, 46, 2) + "/" + GetText(I, 48, 2) + "/" + GetText(I, 50, 4)))
                End If
            Next
            SendF12()
        End Sub

    End Structure
    Public Structure ProcAndDos
        Dim proc As String
        Dim Fdos As String
        Dim Ldos As String
    End Structure

    Public TempProcAndDos(1000) As ProcAndDos
    Public CountTempProcAndDos As Integer
    Public Structure Claim
        'Dim MyISI() As ISI
        'Dim ModFr() As ModifierList
        Dim ServiceLine() As String
        Dim LineCount As Integer
        Dim ICN As String
        Dim Suffix As String
        Dim FLN As String
        Dim DCC As String
        Dim ClaimNumber As String
        Dim CauseCode As String
        Dim Draft As String
        Dim Tin As String
        Dim BillCharge As String
        Dim PaidAmount As String
        Dim DateOfProcessing As String
        Dim AdjusterID As String
        Dim Voided As Boolean
        Dim Line30 As Boolean
        Dim Policy As String
        Dim SSN As String
        Dim EmployeeName As String
        Dim Relationship As String
        Dim Closed() As String
        Dim RemarkCode() As String
        Dim RHI() As String
        Dim DxCode As String
        Dim PrimaryDxCode As String
        Dim ProcCode() As String
        Dim POS() As String
        Dim FDOS() As String
        Dim LDOS() As String
        Dim NBR() As String
        Dim OV() As String
        Dim Charge() As String
        Dim NotCovered() As String
        Dim Allowable() As String
        Dim Deductible() As String
        Dim CoIns() As String
        Dim Paid() As String
        Dim TotalAllowable As String
        Dim TotalDeductible As String
        Dim TotalNotCovered As String
        Dim Copay() As String
        Dim Modifier() As String
    End Structure
    'Public HistoryMHI1(100) As Claim
    Public Structure AHI
        Dim FDOS As String
        Dim LDOS As String
        Dim Paid As String
        Dim Ded As String
        Dim NC As String
        Dim RemarkCode As String
        Dim OV As String
        Dim ICN As String
        Dim ProviderName As String
        Dim MyClaim() As Claim
        Dim MyDraftDetails() As DraftDetails
        Dim MyClaimFromPaymentScreen() As ClaimFromPaymentScreen
        Dim MyClaimDetails As ClaimDetails
        Dim MyISI() As ISI
        Dim MyMap() As MapCurrentProcDetailsWithHistory
        Dim MaxSuffix As Integer
        Dim ProcDT As String
    End Structure
    Public CurrentAHI(1000) As AHI
    Public CountAHI As Integer
    Public Structure ProcDetails
        Dim Prc As String
        Dim FDOS As String
        Dim LDOS As String
        Dim Charge As String
        Dim Modifier As String
    End Structure
    Public HxProcDetails(1000) As ProcDetails

    Public Structure HistoryClaimAllDetails

    End Structure
    Public Structure PreviousValueCheck
        Dim RevCode() As String
        Dim DOS() As String
        Dim Units() As String
        Dim OV() As String
        Dim Charge() As String
    End Structure
    Public PreviosValue(100) As PreviousValueCheck
    Public CurrentValue(100) As PreviousValueCheck
    Public Structure CurrentICNPaymentScreen
        Dim POS() As String
        Dim SVC() As String
        Dim FDate() As String
        Dim LDAte() As String
        Dim NBR() As String
        Dim OV() As String
        Dim RC() As String
        Dim Charge() As String
        Dim LineCount() As Integer
    End Structure
    Public PaymentScreenData(100) As CurrentICNPaymentScreen
    Public Structure PossibleICN
        Dim ICN As String
        Dim ProcCodeMatchCount As Integer
        Dim Position As Integer
        Dim isPaid As Boolean
        Dim IsClosed As Boolean
        Dim E9or12Closed As Boolean
        Dim TotalBillCharge As String
        Dim TotalPaid As String
        Dim IsDXMatch As Boolean
        Dim isVoided As Boolean
        Dim isPartiallyVoided As Boolean
        Dim FLN As String
        Dim TempClaim() As Claim
        Dim TempDraftDetails() As DraftDetails
        Dim TempClaimFromPaymentScreen() As ClaimFromPaymentScreen
        Dim TempClaimDetails As ClaimDetails
        Dim TempISI() As ISI
        Dim Map() As MapCurrentProcDetailsWithHistory
        Dim MaxSuffix As Integer
    End Structure
    Public ListOfPossibleICNActual(1000) As PossibleICN
    Public GlobalMemoryDump() As PossibleICN
    Public ROVline() As String
    Public CurrICN As String, Engine1 As String, Office As String, Session As String, ID As String, Pass As String, Policy As String, DBORGID As String
    Public flagZeroPayment As Boolean = False, flagUnderpayment As Boolean = False, flagOverPayment As Boolean = False
    Public Structure MPRStruct
        Dim ProcCode As String
        Dim Fdos As String
        Dim Ldos As String
        Dim Modifier As String
        Dim RVUValue As Double
        Dim POS As String
        Dim Rank As Integer
        Dim CIndex As Integer
        Dim LIndex As Integer
    End Structure
    Public MyMPRStruct(0) As MPRStruct
    Public FlagRPException As Boolean = False, flagHistoryVoid As Boolean = False
    Public Structure EDS1Value
        Dim ServiceLine As Translation
        Dim ServiceLineCount As String
        Dim FDate As String
        Dim LDate As String
        Dim ServiceWiseTotalCharge As String
    End Structure
    Public Structure Translation
        Dim LineTranslation As String
        Dim RevCode() As String
        Dim RevCodeCount() As String
        Dim CPTCode() As String
        Dim M1() As String
        Dim M2() As String
        Dim DateofService() As String
        Dim BillCharge() As String
        Dim Unit() As String
        Dim PS() As String
        Dim ServiceCode() As String
    End Structure

    Public Structure EDS9Value
        Dim WQ() As String
        Dim PC() As String
        Dim TopsID() As String
        Dim NBR() As String
        Dim Status() As String
        Dim LineColor() As Integer
        Dim LineCount() As Integer
        Dim JulianDate As String
        Dim Comment() As String
    End Structure
    Public Structure HistoryDetails
        Dim POS() As String
        Dim OV() As String
        Dim RC() As String
        Dim FDate() As String
        Dim LDate() As String
        Dim Draft() As String
        Dim Line() As String

    End Structure

    Public HistoryAllDetail(100) As HistoryDetails

    Public Structure STDXCODE
        Dim DXCode() As String
    End Structure
    Public CurrentDXCode(100) As STDXCODE
    Public HistoryDxCode(100) As STDXCODE
    Public Structure PaymentScreen
        Dim Allowable() As String
        Dim CoInsurance() As String
        Dim Deductiable() As String
        Dim Payment() As String
        Dim LineCount() As Integer
    End Structure
    Public HistoryICNDetails(100) As PaymentScreen
    Public BasicHistoryICNDetails(100) As PaymentScreen
    Public Structure HistoryClaimDeils

        Dim FLN As String
        Dim DCC As String
        Dim TIN As String

    End Structure

    Public HistoryDetailsICN(100) As HistoryClaimDeils
    Public Structure MHI
        Dim MAllowable() As String
        Dim MCoInsurance() As String
        Dim MDeductiable() As String
        Dim MPayment() As String
        Dim MLineCount() As Integer
    End Structure
    Public HistoryMHI(100) As MHI
    Public CurrentICNEDS1(100) As EDS1Value
    Public CurrentICNEDS9(100) As EDS9Value
    Public HistoryICNEDS1(100) As EDS1Value
    Public CurreentICNPaymentScreen(100) As PaymentScreen
    Public CurrentIcnMPCScreenIfICES(100) As PaymentScreen
    Public Structure EDS6Point5Screen
        Dim EDS6ServiceLineCount As EDS6LineTranslation
        Dim EDS6LineCount As Integer
    End Structure
    Public CurrentEDS6Point5Screen(100) As EDS6Point5Screen
    Public Structure OriginalEDSValue
        Dim EDSLineCount As OriginalEDS6Point5Value
        Dim RemarkCode As String
        Dim OV As String
        Dim NoteCover As String
    End Structure
    Public Structure OriginalEDS6Point5Value
        Dim LineTranslation As String
        Dim RevCode() As String
        Dim NoteCover() As String
        Dim RemarkCode() As String
    End Structure
    Public Original(100) As OriginalEDSValue
    Public Structure EDS6LineTranslation
        Dim EDS6ServiceTranslation As String
        Dim LineCount() As String
        Dim RevCode() As String
        Dim Charge() As String
        Dim DateOfService() As String
        Dim ProcCode() As String
        Dim M1() As String
        Dim M2() As String
    End Structure
    Public Structure HRHIScreen
        Dim RHIServiceLineCount As HRHI
        Dim RHILineCount As Integer
    End Structure
    Public Structure HRHI
        Dim RHITranslation As String
        Dim LineCount() As String
        Dim RevCode() As String
        Dim Charge() As String
        Dim DateOfService() As String
        Dim ProcCode() As String
        Dim M1() As String
        Dim M2() As String
        Dim Unit() As String
        Dim ServiceCode() As String
    End Structure
    Public HistoryRHI(100) As HRHIScreen
    Public Sub redimAllVariables()
        ReDim CurrentClaim(100)
        ReDim HistoryClaim(100)
        ReDim ISIScreen(100)
        ReDim ISICurrent(100)
        ReDim TempProcAndDos(1000)
        ReDim CurrentAHI(1000)
        ReDim HxProcDetails(1000)
        ReDim ListOfPossibleICNActual(1000)
        ReDim CurrentClaimPaymentScreen(0)
        ReDimension(StructReconCode, 99)
        HardCopyText = ""
        flagExit = False
        is30LineToBeused = False
        FlagScreenNotValid = False
        strPendO2SuffixList = ""
        strEncryptedText = ""
        strkeyValues = ""
        ReDim CurrentClaim(100)
        ReDim HistoryClaim(100)
        HistoryClaimCount = 0
        CurrentClaimCount = 0
        ClaimLevelRemarkCode = ""
        LineLevelRemarkCode = ""
        LineLevelRemarkCodePriority2 = ""
        LineLevel = ""
        flagProcessed = False
        FlagMHI = False

        ReDim CreateMapinISI(100)
        ReDim ISIScreen(100)
        ReDim ISICurrent(100)
        ReDim ISICurrentActual(100)
        ReDim CurrentClaimPaymentScreen(100)
        ReDim HistoryClaimPaymentScreen(100)
        ReDim Map(100)
        ReDim CurrentClaimModifier(100)
        CurrentClaimDetails.Policy = ""
        CurrentClaimDetails.SSN = ""
        CurrentClaimDetails.PatientName = ""
        CurrentClaimDetails.PatientRelationship = ""
        CurrentClaimDetails.ICN = ""
        CurrentClaimDetails.Tin = ""
        CurrentClaimDetails.FLN = ""
        CurrentClaimDetails.DCC = ""
        HistoryClaimDetails.Policy = ""
        HistoryClaimDetails.SSN = ""
        HistoryClaimDetails.PatientName = ""
        HistoryClaimDetails.PatientRelationship = ""
        HistoryClaimDetails.ICN = ""
        HistoryClaimDetails.Tin = ""
        HistoryClaimDetails.FLN = ""
        HistoryClaimDetails.DCC = ""
        ReDim CAline(100)
        ReDim TempProcAndDos(100)
        CountTempProcAndDos = 0
        ReDim CurrentAHI(100)
        CountAHI = 0
        ReDim HxProcDetails(100)
        ReDim ListOfPossibleICNActual(100)
        ReDim EDSSText(0)

        flagZeroPayment = False
        flagUnderpayment = False
        flagOverPayment = False
        ReDimension(MyMPRStruct, 0)
        FlagCARemarkCode = False
        FlagRPException = False
        FlagUndoPerformed = False
        flagHistoryVoid = False
        IsCOB = False
        ReDim CurrentICNEDS1(0)
    End Sub
    Public Sub ReDimension(ByRef myclaim() As PossibleICN, ByVal index As Integer)
        ReDim myclaim(index)
        'test comment
        For i As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(i).ICN = ""
            myclaim(i).ProcCodeMatchCount = 0
            myclaim(i).Position = 0
            myclaim(i).isPaid = False
            myclaim(i).IsClosed = False
            myclaim(i).E9or12Closed = False
            myclaim(i).TotalBillCharge = ""
            myclaim(i).TotalPaid = ""
            myclaim(i).IsDXMatch = False
            myclaim(i).isVoided = False
            myclaim(i).isPartiallyVoided = False
            myclaim(i).FLN = ""
            ReDim myclaim(i).TempClaim(0)
            ReDimension(myclaim(i).TempClaimFromPaymentScreen, 0)
            ReDimension(myclaim(i).TempDraftDetails, 0)
            ReDimension(myclaim(i).TempISI, 0)
            myclaim(i).MaxSuffix = 0
            ReDimension(myclaim(i).Map, 0)
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As ProcDetails, ByVal index As Integer)
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).Prc = ""
            myclaim(I).FDOS = ""
            myclaim(I).LDOS = ""
            myclaim(I).Charge = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As AHI, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).FDOS = ""
            myclaim(I).LDOS = ""
            myclaim(I).Paid = ""
            myclaim(I).Ded = ""
            myclaim(I).NC = ""
            myclaim(I).RemarkCode = ""
            myclaim(I).OV = ""
            myclaim(I).ICN = ""
            myclaim(I).ProviderName = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As ProcAndDos, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).proc = ""
            myclaim(I).Fdos = ""
            myclaim(I).Ldos = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As ISI, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).ProcCode = ""
            myclaim(I).BillCharge = ""
            myclaim(I).Modifier = ""
            myclaim(I).DispositionCode = ""
            myclaim(I).LineNumber = ""
            myclaim(I).flag = ""
            myclaim(I).ClaimIndex = -1
            myclaim(I).LineIndex = -1
            myclaim(I).CommentICN = ""
            myclaim(I).CommentProc = ""
            myclaim(I).FullComment = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As MRIPat, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).Name = ""
            myclaim(I).Relationship = ""
            myclaim(I).BirthDate = ""
            myclaim(I).MDEff = ""
            myclaim(I).MdCanc = ""
            myclaim(I).S = ""
            myclaim(I).C = ""
            myclaim(I).M = ""
            myclaim(I).EffDate = ""
            myclaim(I).CancDate = ""
            myclaim(I).T = ""
            myclaim(I).R = ""
            myclaim(I).Ret_ST = ""
            myclaim(I).CobDT = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myOBJ() As ReconCode, ByVal Index As Integer)
        If Index < 0 Then
            ReDimension(myOBJ, 0)
            Exit Sub
        End If
        ReDim myOBJ(Index)
        For I As Integer = LBound(myOBJ) To UBound(myOBJ)
            myOBJ(I).ReconCd = ""
            myOBJ(I).RemarkCode = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As CAStruct, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).LineIndex = -1
            myclaim(I).ClaimIndex = -1
        Next
    End Sub
    Public Sub ReDimension(ByRef MyStruct() As MPRStruct, ByVal index As Integer)
        If index < 0 Then
            ReDimension(MyStruct, 0)
            Exit Sub
        End If
        ReDim MyStruct(index)
        For I As Integer = LBound(MyStruct) To UBound(MyStruct)
            MyStruct(I).CIndex = -1
            MyStruct(I).Fdos = ""
            MyStruct(I).Ldos = ""
            MyStruct(I).LIndex = -1
            MyStruct(I).Modifier = ""
            MyStruct(I).POS = ""
            MyStruct(I).ProcCode = ""
            MyStruct(I).Rank = -1
            MyStruct(I).RVUValue = -999
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As ModifierList, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).POS = ""
            myclaim(I).ProcedureCode = ""
            myclaim(I).Modifier = ""
            myclaim(I).FDOS = ""
            myclaim(I).LDOS = ""
            myclaim(I).Suffix = ""
            myclaim(I).LineNumber = 0
            myclaim(I).Charge = ""
            myclaim(I).NBR = ""
            myclaim(I).ServicingPhysician = ""
            myclaim(I).ReferringPhysician = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As MatchingLines, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).ServiceLines = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As DraftDetails, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).Draft = ""
            myclaim(I).Suffix = ""
            myclaim(I).Reversal = False
            myclaim(I).ReversalType = ""
            myclaim(I).DateOfProcessing = ""
            myclaim(I).TotalPaid = 0
            myclaim(I).TotalAllowable = 0
            myclaim(I).TotalNotCovered = 0
            myclaim(I).isClosed = False
            myclaim(I).ClosureCode = ""
            myclaim(I).TotalBillCharge = 0
            myclaim(I).DateOfReversal = ""
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As MapCurrentProcDetailsWithHistory, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            myclaim(I).CurrentClaimProcCode = ""
            myclaim(I).CurrentClaimFdos = ""
            myclaim(I).CurrentClaimLdos = ""
            myclaim(I).CurrentClaimUnit = ""
            myclaim(I).CurrentClaimCharge = 0
            myclaim(I).CurrentClaimNotCovered = 0
            myclaim(I).CurrentClaimSuffix = -1
            myclaim(I).CurrentClaimLineNumber = -1
            myclaim(I).CurrentClaimAllowable = "0"
            myclaim(I).CurrentClaimDeductible = "0"
            myclaim(I).CurrentClaimCoins = "0"
            myclaim(I).HistoryClaimProcCode = ""
            myclaim(I).HistoryClaimFdos = ""
            myclaim(I).HistoryClaimLdos = ""
            myclaim(I).HistoryClaimUnit = ""
            myclaim(I).HistoryClaimCharge = 0
            myclaim(I).HistoryClaimNotCovered = 0
            myclaim(I).HistoryClaimAllowable = "0"
            myclaim(I).HistoryClaimDeductible = "0"
            myclaim(I).HistoryClaimCoins = "0"
            myclaim(I).HistoryClaimCIndex = -1
            myclaim(I).HistoryClaimLIndex = -1
            myclaim(I).HXICN = ""
            myclaim(I).IsCodeChange = False
        Next
    End Sub
    Public Sub ReDimension(ByRef myclaim() As ClaimFromPaymentScreen, ByVal index As Integer)
        If index < 0 Then
            ReDimension(myclaim, 0)
            Exit Sub
        End If
        ReDim myclaim(index)
        For I As Integer = LBound(myclaim) To UBound(myclaim)
            ReDim myclaim(I).POS(0)
            myclaim(I).POS(0) = ""
            ReDim myclaim(I).ProcCode(0)
            myclaim(I).ProcCode(0) = ""
            ReDim myclaim(I).FDOS(0)
            myclaim(I).FDOS(0) = ""
            ReDim myclaim(I).LDOS(0)
            myclaim(I).LDOS(0) = ""
            ReDim myclaim(I).NBR(0)
            myclaim(I).NBR(0) = "0"
            ReDim myclaim(I).Override(0)
            myclaim(I).Override(0) = ""
            ReDim myclaim(I).PayeeCode(0)
            myclaim(I).PayeeCode(0) = ""
            ReDim myclaim(I).RemarkCode(0)
            myclaim(I).RemarkCode(0) = ""
            ReDim myclaim(I).Charge(0)
            myclaim(I).Charge(0) = "0"
            ReDim myclaim(I).NotCovered(0)
            myclaim(I).NotCovered(0) = "0"
            ReDim myclaim(I).Allowable(0)
            myclaim(I).Allowable(0) = "0"
            ReDim myclaim(I).Deductible(0)
            myclaim(I).Deductible(0) = "0"
            ReDim myclaim(I).Coins(0)
            myclaim(I).Coins(0) = "0"
            ReDim myclaim(I).Paid(0)
            myclaim(I).Paid(0) = "0"
            ReDim myclaim(I).Sanction(0)
            myclaim(I).Sanction(0) = ""
            myclaim(I).EditWarning = ""
            myclaim(I).CauseCode = ""
            myclaim(I).Draft = "0000000000"
            myclaim(I).DateOfProcessing = "99/99/99"
            myclaim(I).TotalPaid = "0"
        Next
    End Sub
    Public Sub ReDimension(ByRef MyClaim() As Claim, ByVal ICN As String)
        Dim Count As Integer = -1
        Dim CurrentScreen As String = "", findObject
        PullMHI(ICN)

        For Row As Integer = 10 To 18 Step 4
            autECLPSObj.auteclfieldlist.Refresh
            apiChk()
            findObject = autECLPSObj.auteclfieldlist.FindFieldByRowCol(Row, 1)
            If autECLPSObj.GetText(Row + 2, 1, 3) = "ICN" And findObject.Display Then
                Count += 1
            End If
            If Row = 18 Then
                If InStr(autECLPSObj.GetText(24, 1, 80), "E065NO MORE CLMS ON FILE") OrElse InStr(autECLPSObj.GetText(24, 1, 80), "W062NO CLAIMS ON FILE") Then Exit For
                CurrentScreen = GetTextRect(10, 1, 20, 79)
                Do
                    SendF8()
                Loop While CurrentScreen = GetTextRect(10, 1, 20, 79)
                Row = 6
            End If
        Next
        ReDim MyClaim(Count + 1)
    End Sub
End Module

