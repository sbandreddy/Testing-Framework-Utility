import pandas as pd
import xlsxwriter

input_ConfigData = pd.read_excel('CCP_InputFile.xlsx')

data_dict = {}

for index, row in input_ConfigData.iterrows():

    key = row['Field']
    value = row['Value']

    if key in data_dict:
        data_dict[key].append(value)
    else:
        data_dict[key] = [value]

key = 'SourceSystem'
if key in data_dict:
    SourceSystem = data_dict[key]
    SourceSystem = ', '.join(map(str, SourceSystem))

key = 'Module'
if key in data_dict:
    Module = data_dict[key]
    Module = ', '.join(map(str, Module))

key = 'Instance'
if key in data_dict:
    Instance = data_dict[key]
    Instance = ', '.join(map(str, Instance))

key = 'CaseName'
if key in data_dict:
    CaseName = data_dict[key]
    CaseNameList = list(CaseName)

key = 'KeyColumnForStgMarCall'
if key in data_dict:
    KeyColumnForStgMarCall = data_dict[key]

key = 'KeyColumnForStgColl'
if key in data_dict:
    KeyColumnForStgColl = data_dict[key]

key = 'KeyColumnForStgAlloc'
if key in data_dict:
    KeyColumnForStgAlloc = data_dict[key]

key = 'KeyColumnFor_MarCall'
if key in data_dict:
    KeyColumnFor_MarCall = data_dict[key]

key = 'KeyColumnFor_ColBal'
if key in data_dict:
    KeyColumnFor_ColBal = data_dict[key]

key = 'KeyColumnFor_MCAlloc'
if key in data_dict:
    KeyColumnFor_MCAlloc = data_dict[key]



def suiteName(Module, SourceSystem):
    SuiteNameSheet1 = Module + '_' + SourceSystem
    return SuiteNameSheet1


def suiteDescription(SourceSystem):
    suiteDescriptionSheet1 = 'It Contain ' + SourceSystem + ' related gateway cases'
    return suiteDescriptionSheet1


def moduleName(Module):
    moduleNameSheet1 = Module
    return moduleNameSheet1


def caseName(CaseName):
    caseNameSheet1 = CaseName
    CaseName = ', '.join(map(str, CaseName))
    return caseNameSheet1


configData = 'ConfigData'

workbook = xlsxwriter.Workbook(Module + '_' + configData + '_' + SourceSystem + '.xlsx')

worksheet1 = workbook.add_worksheet("TestSuiteHeaderDetail")
worksheet2 = workbook.add_worksheet("TestCaseHeaderDetail")
worksheet3 = workbook.add_worksheet("ActionHeader")
worksheet4 = workbook.add_worksheet("SnapshotCaptureConfiguration")
worksheet5 = workbook.add_worksheet("SnapshotCompareConfiguration")
worksheet6 = workbook.add_worksheet("ImportDataConfiguration")
worksheet7 = workbook.add_worksheet("AlterTableConfiguration")
worksheet8 = workbook.add_worksheet("ExecuteScriptConfiguration")
worksheet9 = workbook.add_worksheet("WaitForProcessConfiguration")

Header1 = ["SuiteName", "SuiteDescription", "Modules", "CaseName", "ExecutionOrder", "ProceedOnFail"]
Header2 = ["Modules", "CaseName", "CaseDescription", "ActionName", "ExecutionOrder", "ExternalReference"]
Header3 = ["ActionName", "ActionType", "ActionDescription", "ProceedOnFail"]
Header4 = ["TableName", "ActionName", "WhereCondition", "SelectColumns", "GroupByColumns", "SqlStatement", "KeyColumn"]
Header5 = ["ActionName", "PreActionName", "PostActionName", "CompareCondition", "KeyColumn", "CompareColumns"]
Header6 = ["ActionName", "TableName", "InputFilePath", "TabName", "WhereCondition", "KeyColumn"]
Header7 = ["ActionName", "SqlStatement"]
Header8 = ["ActionName", "Directory", "ScriptName", "InputArguments"]
Header9 = ["ActionName", "Time"]

Worksheets = [worksheet1, worksheet2, worksheet3, worksheet4, worksheet5, worksheet6, worksheet7,
              worksheet8, worksheet9]
Headers = [Header1, Header2, Header3, Header4, Header5, Header6, Header7, Header8, Header9]

for i in range(0, len(Worksheets), 1):
    row = 0
    column = 0
    for header_column_name in Headers[i]:
        Worksheets[i].write(row, column, header_column_name)
        column += 1

w30Sec = 'WaitFor30Seconds'
w1Min = 'WaitForOneMin'
w2Min = 'WaitForTwoMins'
w3Min = 'WaitForThreeMins'
w5Min = 'WaitForFiveMins'
w10Min = 'WaitForTenMins'

Sec30 = int(30000)
Min1 = int(60000)
Min2 = int(120000)
Min3 = int(180000)
Min5 = int(300000)
Min10 = int(600000)

wList = [w30Sec, w1Min, w2Min, w3Min, w5Min, w10Min]
mList = [Sec30, Min1, Min2, Min3, Min5, Min10]

allTables = ['StagingRawPayLoad', 'StagingMarginCall', 'StagingCollateralBalance', 'StagingMarginCollateralAllocation',
             'MarginCall', 'CollateralBalance', 'MarginCollateralAllocation']

deleteQueries = ['delete from StagingRawPayLoad where SourceSystem = :SourceSystem '
                 'and cast(LastUpdated as Date) = :Businessdate',
                 'delete from StagingMarginCall where SourceSystem = :SourceSystem '
                 'and Businessdate in (:Yesterdaydate, :Businessdate)',
                 'delete from StagingCollateralBalance where SourceSystem = :SourceSystem '
                 'and  Businessdate in (:Yesterdaydate, :Businessdate)',
                 'delete from StagingMarginCollateralAllocation where SourceSystem = :SourceSystem '
                 'and Businessdate in (:Yesterdaydate, :Businessdate)',
                 'delete from MarginCall where SourceSystem = :SourceSystem '
                 'and Businessdate in (:Yesterdaydate, :Businessdate)',
                 'delete from CollateralBalance where SourceSystem = :SourceSystem '
                 'and Businessdate in (:Yesterdaydate, :Businessdate)',
                 'delete from MarginCollateralAllocation where SourceSystem = :SourceSystem '
                 'and Businessdate in (:Yesterdaydate, :Businessdate)']

svcDirectory = '/data/transcend/coreqaauto/current/scripts/services'
startScript = 'Start_Connectivityscript_' + SourceSystem
scriptName = 'start-connectivity-engine.sh'

stgMarCallColumns = 'SourceSystem,AsOfBusinessDate,BusinessDate,CallDate,AgreementID,AgreementType,CounterpartyID,CounterpartyName,' \
                    'CallStatus,Currency,CallType,MTM,IMRequirement,CallAmount,AgreedAmountCash,AgreedAmountNonCash,' \
                    'CounterpartyCallAmount,CallAmountDifference,Region,FXrate,CurrentBalanceCash,' \
                    'CurrentBalanceNonCash,LegalEntityName,AgreementSourceSystem,AccountId,FeedName,' \
                    'LegalEntityType,ExternalAccountId,ExternalAccountSourceSystem,CounterPartySource,' \
                    'CounterpartyIdType,RequirementType,BalanceType,ExcessOrDeficit,TotalCollateralBalance,' \
                    'NetOptionValue,CoreIMRequirement,ConcentrationIMRequirement,DeliveryIMRequirement,' \
                    'VariationRequirement,TotalRequirement,CashSettlement,LegalEntityId,LegalEntityIdType,' \
                    'FirmAccountId,FirmAccountIdType,CounterpartySourceSystem,ExternalAccountIdType,' \
                    'CallCurrency,CallNotificationType,MarginType,MinimumCashBalance,MinimumCashTreasuryBalance,' \
                    'FirmAccountName,MinimumCashRequirement,IneligibleCollateralValue,' \
                    'SettlementToBalance,TotalRequirementCurrency,Version,FeedType,FeedFlowId,riskCurrency'

stgCollColumns = 'SourceSystem,AgreementID,AgreementType,AsOfBusinessDate,BusinessDate,LegalEntity,CounterpartyID,CounterpartyName,' \
                 'CollateralBalanceType,CollateralType,CollateralStatus,ProductId,ProductAltId,ProductIdType,' \
                 'CollateralCurrency,TransactionType,MarketValue,MarginValue,AgreementSourceSystem,BalanceType,' \
                 'FirmAcntId,FirmAcntIdSourceSystem,CounterPartyIdType,LegalEntityIdType,AccountId,' \
                 'AccountIdSourceSystem,Notional,FeedName,AccountType,MaturityDate,AvailableMarginValue,' \
                 'AppliedMarginValue,LegalEntityId,LegalEntityName,FirmAccountId,FirmAccountIdType,' \
                 'CounterpartySourceSystem,CounterpartyAccountId,CounterpartyAccountSourceSystem,' \
                 'ExternalAccountId,ExternalAccountIdType,CollateralQuantity,CleanPrice,CleanMarketValue,' \
                 'FirmAccountName,Version,FeedType,FeedFlowId,MarginType'

stgAllocColumns = 'AssetPBAmountRemaining,AsOfBusinessDate,BusinessDate,CollateralCurrency,CollateralType,FXrate,MarginValue,' \
                  'RequirementAmount,RequirementCurrency,RequirementType,SourceSystem,FeedName,ErrorCode,' \
                  'ErrorDescription,CounterpartyId,CounterpartyIdType,CounterpartyName,LegalEntityId,' \
                  'LegalEntityIdType,LegalEntityName,FirmAccountId,CrossCurrencyHaircut,' \
                  'AppliedValueRequirementCurrency,AvailableMarginValue,FirmAccountName,CounterpartySourceSystem,' \
                  'CounterpartyAccountId,CounterpartyAccountSourceSystem,Region,TenantId,Version,FeedType,FeedFlowId,' \
                  'MarginType,ExternalAccountId,ExternalAccountIdType,FirmAccountIdType'


MarCallQuery = 'Select SourceVer,SourceSystem,BusinessDate,CallDate,CounterpartyName,Currency,BillingFees,' \
               'BillingInterest,CallAmount,AgreedAmountNonCash,CallAmountDifference,CallAmountBase,' \
               'CONVERT(decimal(19,2), AgreedAmountCashBase) as AgreedAmountCashBase,AgreedAmountNonCashBase,' \
               'CONVERT(decimal(19,2),CurrentBalanceCash) as CurrentBalanceCash,CallAmountDifferenceBase,' \
               'Region,FXrate,CONVERT(decimal(19,2),CurrentBalanceNonCash) as CurrentBalanceNonCash,' \
               'InTransitAmountCash,InTransitAmountNonCash,TradeCount,Threshold,ProcessingStatus,LegalEntityName,' \
               'CounterpartySourceSystem,CounterpartyIdType,SourceCounterPartyId,SourceAccountId,' \
               'AccountIdSourceSystem,FeedName,SrcLegalEntityId,SourceLegalEntityName,LegalEntityType,' \
               'AsOfBusinessDate,FirmAcntAltId,FirmAcntAltIdType,CONVERT(decimal(19,2), MTMBASE) as MTMBASE,' \
               'THRESHOLDBASE,EXPOSUREDATE,CONVERT(decimal(38,2),TotalCollateralBalance) as TotalCollateralBalance,' \
               'InTransitAmountCashBase,InTransitAmountNonCashBase,CONVERT(decimal(19,2),' \
               'CurrentBalanceCashHeldBase) as CurrentBalanceCashHeldBase,CONVERT(decimal(19,2),' \
               'CurrentBalanceNonCashHeldBase) as CurrentBalanceNonCashHeldBase,BalanceType,CONVERT(decimal(38,2),' \
               'ExcessOrDeficit) as ExcessOrDeficit,' \
               'CONVERT(decimal(38,2), ExcessOrDeficitBase) as ExcessOrDeficitBase,' \
               'CONVERT(decimal(38,2),TotalCollateralBalanceBase) as TotalCollateralBalanceBase,' \
               'CurrentDayActivityCash,CurrentDayActivityNonCash,Status,NetOptionValue,' \
               'Floor(cast(DeliveryIMRequirement as float)) as DeliveryIMRequirement,' \
               'ConcentrationIMRequirement,CONVERT(decimal(38,2),VariationRequirement) as VariationRequirement,' \
               'CONVERT(decimal(38,2), CoreIMRequirementBase) as CoreIMRequirementBase,' \
               'CONVERT(decimal(38,2),ConcentrationIMRequirementBase) as ConcentrationIMRequirementBase ,' \
               'CONVERT(decimal(38,2), DeliveryIMRequirementBase) as DeliveryIMRequirementBase,RiskCounterpartyName,' \
               'CONVERT(decimal(38,2),TotalRequirement) as TotalRequirement,CONVERT(decimal(38,2),' \
               'TotalRequirementBase) as TotalRequirementBase,ExternalAccountId,ReserveAdditionalRisk,' \
               'ReserveAdditionalRiskBase,CashSettlement,MarginType,MinimumCashBalance,MinimumCashTreasuryBalance,' \
               'eMIRAddOnRequirement,eMIRAddOnRequirementBase,CallType,RequirementType,CallStatus,CoreIMRequirement,' \
               'BusinessLine,MinimumExcess,IMRequirement,IneligibleCollateralValue,' \
               'CONVERT(decimal(19,2),AgreedAmountCash) as AgreedAmountCash,SourceMarginCallId ' \
               'from MarginCall where SOURCESYSTEM=:SourceSystem and BusinessDate = :ExpectedDate and' \
               ' FeedName = :FeedName_Exp and FeedType = :FeedType_Exp and AsOfBusinessDate = :AsOfBusinessDate'

ColBalQuery = 'Select SourceVer,SourceSystem,LegalEntityName,SourceLegalEntityName,CounterpartyName,' \
              'CounterpartySourceSystem,SourceCounterpartyName,CollateralBalanceType,CollateralType,CollateralStatus,' \
              'ProductAltId,ProductIdType,CollateralCurrency,TransactionType,' \
              'CONVERT(decimal(38,2), CollateralQty) as CollateralQty,PriceClean,MarketValueClean,' \
              'CptySrcSafeKeepId,Price,CONVERT(decimal(38,2), MarketValue) as MarketValue,Haircut,' \
              'CONVERT(decimal(19,2), MarginValue) as MarginValue,' \
              'CONVERT(decimal(19,2), MarginValueBase) as MarginValueBase,FXrate,CounterpartyNegativeInterest,' \
              'PrincipalInterestIndex,PrincipalInterestSpread,PrincipalNegativeInterest,ProcessingStatus,Region,' \
              'CONVERT(decimal(38,2), MarketValueBase) as MarketValueBase,SourceLegalEntityId,LegalEntityIdType,' \
              'SourceCounterPartyId,ExternalAccountId,CounterPartyIdType,CustAcntId,BalanceType,AccountId,' \
              'AccountName,SourceAccountId,AccountIdSourceSystem,Notional,AsOfBusinessDate,ReportingCounterPartyName,' \
              'FirmAcntAltId,FirmAcntAltIdType,convert(varchar, MaturityDate, 105) as MaturityDate,Status,CouponRate,' \
              'tenantId,isnull(CollateralJurisdiction, 0) as CollateralJurisdiction,BusinessDate,MarginOffset,' \
              'AppliedMarginValue,AvailableMarginValue,IneligibleCollateralValue,prodDesc,HaircutType,BusinessLine,' \
              'SegStatus,CollateralLocation,CounterPartySafeKeepingIdType from CollateralBalance ' \
              'where SOURCESYSTEM=:SourceSystem and BusinessDate = :ExpectedDate and  FeedName = :FeedName_Exp and ' \
              'FeedType = :FeedType_Exp and AsOfBusinessDate = :AsOfBusinessDate'

MCAllocQuery = 'Select Ver,BusinessDate,AsOfBusinessDate,SourceSystem,SrcLegalEntityId,FirmAcntAltId,' \
               'SourceCounterPartyId,CollateralType,CollateralCurrency,RequirementCurrency,RequirementType,' \
               'CONVERT(decimal(38,2), RequirementAmount) as RequirementAmount,ExternalAccountId,' \
               'CONVERT(decimal(38,2), MarginValue) as MarginValue,FXrate,ProcessingStatus,Status,CounterpartyIdType,' \
               'CounterpartyName,LegalEntityIdType,LegalEntityName,FirmAccountId,' \
               'CONVERT(decimal(25,2), AppliedValueRequirementCurrency) as AppliedValueRequirementCurrency,' \
               'AvailableMarginValue,CounterpartySourceSystem,FirmAccountName,CounterpartyAccountId,' \
               'CounterpartyAccountSourceSystem from MarginCollateralAllocation where SOURCESYSTEM=:SourceSystem and ' \
               'BusinessDate = :ExpectedDate and  FeedName = :FeedName_Exp and ' \
               'FeedType = :FeedType_Exp and AsOfBusinessDate = :AsOfBusinessDate'

marginCallColumns = 'SourceVer,SourceSystem,BusinessDate,CallDate,CounterpartyName,Currency,BillingFees,' \
                    'BillingInterest,CallAmount,AgreedAmountNonCash,CallAmountDifference,CallAmountBase,' \
                    'AgreedAmountCashBase,AgreedAmountNonCashBase,CurrentBalanceCash,CallAmountDifferenceBase,' \
                    'Region,FXrate,CurrentBalanceNonCash,InTransitAmountCash,InTransitAmountNonCash,TradeCount,' \
                    'Threshold,ProcessingStatus,LegalEntityName,CounterpartySourceSystem,CounterpartyIdType,' \
                    'SourceCounterPartyId,SourceAccountId,AccountIdSourceSystem,FeedName,SrcLegalEntityId,' \
                    'SourceLegalEntityName,LegalEntityType,AsOfBusinessDate,FirmAcntAltId,FirmAcntAltIdType,' \
                    'MTMBASE,THRESHOLDBASE,EXPOSUREDATE,,TotalCollateralBalance,InTransitAmountCashBase,' \
                    'InTransitAmountNonCashBase,,CurrentBalanceCashHeldBase,CurrentBalanceNonCashHeldBase,' \
                    'BalanceType,ExcessOrDeficit,ExcessOrDeficitBase,TotalCollateralBalanceBase,' \
                    'CurrentDayActivityCash,CurrentDayActivityNonCash,Status,NetOptionValue,DeliveryIMRequirement,' \
                    'ConcentrationIMRequirement,VariationRequirement,CoreIMRequirementBase,' \
                    'ConcentrationIMRequirementBase,DeliveryIMRequirementBase,RiskCounterpartyName,' \
                    'TotalRequirement,TotalRequirementBase,ExternalAccountId,ReserveAdditionalRisk,' \
                    'ReserveAdditionalRiskBase,CashSettlement,MarginType,MinimumCashBalance,' \
                    'MinimumCashTreasuryBalance,eMIRAddOnRequirement,eMIRAddOnRequirementBase,CallType,' \
                    'RequirementType,CallStatus,CoreIMRequirement,BusinessLine,MinimumExcess,IMRequirement,' \
                    'IneligibleCollateralValue,AgreedAmountCash'

colBalColumns = 'SourceVer,SourceSystem,LegalEntityName,SourceLegalEntityName,CounterpartyName,' \
                'CounterpartySourceSystem,SourceCounterpartyName,CollateralBalanceType,CollateralType,' \
                'CollateralStatus,ProductAltId,ProductIdType,CollateralCurrency,TransactionType,CollateralQty,' \
                'PriceClean,MarketValueClean,CptySrcSafeKeepId,Price,MarketValue,Haircut,MarginValue,MarginValueBase,' \
                'FXrate,CounterpartyNegativeInterest,PrincipalInterestIndex,PrincipalInterestSpread,' \
                'PrincipalNegativeInterest,ProcessingStatus,Region,MarketValueBase,SourceLegalEntityId,' \
                'LegalEntityIdType,SourceCounterPartyId,ExternalAccountId,CounterPartyIdType,CustAcntId,BalanceType,' \
                'AccountId,AccountName,SourceAccountId,AccountIdSourceSystem,Notional,AsOfBusinessDate,' \
                'ReportingCounterPartyName,FirmAcntAltId,FirmAcntAltIdType,MaturityDate,Status,CouponRate,tenantId,' \
                'CollateralJurisdiction,BusinessDate,MarginOffset,AppliedMarginValue,AvailableMarginValue,' \
                'IneligibleCollateralValue,prodDesc,HaircutType,BusinessLine,SegStatus,CollateralLocation,' \
                'CounterPartySafeKeepingIdType'

mcAllocColumns = 'Ver,BusinessDate,AsOfBusinessDate,SourceSystem,SrcLegalEntityId,FirmAcntAltId,' \
                 'SourceCounterPartyId,CollateralType,CollateralCurrency,RequirementCurrency,RequirementType,' \
                 'RequirementAmount,ExternalAccountId,MarginValue,AppliedValueReqCcy,FXrate,ProcessingStatus,Status,' \
                 'CounterpartyIdType,CounterpartyName,LegalEntityIdType,LegalEntityName,FirmAccountId,' \
                 'AppliedValueRequirementCurrency,AvailableMarginValue,CounterpartySourceSystem,FirmAccountName,' \
                 'CounterpartyAccountId,CounterpartyAccountSourceSystem'

# TestSuiteHeaderDetail

row = 1

if Module == 'Connectivity':

    for cases in CaseNameList:

        column = 0
        worksheet1.write(row, column, suiteName(Module, SourceSystem))
        column = 1
        worksheet1.write(row, column, suiteDescription(SourceSystem))
        column = 2
        worksheet1.write(row, column, moduleName(Module))
        column = 3
        worksheet1.write(row, column, cases)
        column = 5
        worksheet1.write(row, column, 'TRUE')
        row += 1

if Module == 'Margin':

    for cases in CaseNameList:

        if 'Stg' not in cases:

            column = 0
            worksheet1.write(row, column, suiteName(Module, SourceSystem))
            column = 1
            worksheet1.write(row, column, suiteDescription(SourceSystem))
            column = 2
            worksheet1.write(row, column, moduleName(Module))
            column = 3
            worksheet1.write(row, column, cases)
            column = 5
            worksheet1.write(row, column, 'TRUE')
            row += 1


row = 1

for i in range(len(CaseNameList)):
    column = 4
    worksheet1.write(row, column, i+1)
    row += 1

# TestCaseHeaderDetail

stagingTables = ['StagingMarginCall', 'StagingCollateralBalance', 'StagingMarginCollateralAllocation']
mainTables = ['MarginCall', 'CollateralBalance', 'MarginCollateralAllocation']
alterTables = ['StagingRawPayload', 'StagingMarginCall',
               'StagingCollateralBalance', 'StagingMarginCollateralAllocation']

connectivity = 'Connectivity'
cleanUp = 'CleanUp'
imPort = 'Import'
captureOn = 'CaptureOn'
compare = 'Compare'
stagingMarginCall = 'StagingMarginCall'
stagingCollateralBalance = 'StagingCollateralBalance'
stagingMarginCollateralAllocation = 'StagingMarginCollateralAllocation'
marginCall = 'MarginCall'
collateralBalance = 'CollateralBalance'
margincollateralAllocation = 'MarginCollateralAllocation'
caseDescription = ' Testcase verification for ' + SourceSystem
initializeCaseDesc = 'Initializing the CCP Load'

initializeCaptureList = [captureOn + '_' + sub for sub in allTables]
initializewaitandstart = [startScript, w5Min]

initializeCaptureList.extend(initializewaitandstart)

row = 1

if Module == 'Connectivity':

    for m in CaseNameList:
        if 'Initialize' in m:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 2
            worksheet2.write(row, column, initializeCaseDesc)
            for tables in allTables:
                column = 1
                worksheet2.write(row, column, m)
                column = 3
                worksheet2.write(row, column, cleanUp + '_' + tables)
                row += 1

    row = 1

    for r in range(len(allTables)):
        column = 4
        worksheet2.write(row, column, r + 1)
        row += 1

if Module == 'Connectivity':

    for cases in CaseNameList:

        if '_StgMarCall_' in cases:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, stagingMarginCall + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + stagingMarginCall)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
        if '_StgColl_' in cases:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, stagingCollateralBalance + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + stagingCollateralBalance)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
        if '_StgAlloc_' in cases:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, stagingMarginCollateralAllocation + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + stagingMarginCollateralAllocation)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
    ######################################################
if Module == 'Margin':

    for cases in CaseNameList:

        if '_MarCall_' in cases:

            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, marginCall + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + marginCall)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
        if '_ColBal_' in cases:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, collateralBalance + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + collateralBalance)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
        if '_MCAlloc_' in cases:
            column = 0
            worksheet2.write(row, column, moduleName(Module))
            column = 1
            worksheet2.write(row, column, cases)
            column = 2
            worksheet2.write(row, column, margincollateralAllocation + caseDescription)
            column = 3
            worksheet2.write(row, column, imPort + '_' + margincollateralAllocation)
            column = 4
            worksheet2.write(row, column, int(1))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, captureOn + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(2))
            row += 1
            column = 1
            worksheet2.write(row, column, cases)
            column = 3
            worksheet2.write(row, column, compare + '_' + cases)
            column = 4
            worksheet2.write(row, column, int(3))
            row += 1
            ###########################################

# ActionHeader

importData = 'ImportData'
snapshotCapture = 'SnapshotCapture'
snapshotCompare = 'SnapshotCompare'
waitTime = 'WaitForProcess'
alterTable = 'AlterTable'
executeScript = 'ExecuteScript'
actionDescCapture = 'Capture on Staging Table for ' + SourceSystem
actionDescMainTableCapture = 'Capture on Main Table for ' + SourceSystem
actionDescCompare = 'Compare Database and Input Data for ' + SourceSystem
actionDescImport = 'Importing Gateway Feedload data'
actionDescAlter = 'Cleanup the Staging tables for current BusinessDate'
executeDesc = 'Execute the connectivity script to load feeds'

row = 1

if Module == 'Connectivity':

    column = 0
    worksheet3.write(row, column, startScript)
    column = 1
    worksheet3.write(row, column, executeScript)
    column = 2
    worksheet3.write(row, column, executeDesc)
    column = 3
    worksheet3.write(row, column, 'FALSE')
    row += 1

    for n in alterTables:

        column = 0
        worksheet3.write(row, column, cleanUp + '_' + n)
        column = 1
        worksheet3.write(row, column, alterTable)
        column = 2
        worksheet3.write(row, column, actionDescAlter)
        column = 3
        worksheet3.write(row, column, 'TRUE')
        row += 1

    for i in stagingTables:
        column = 0
        worksheet3.write(row, column, imPort + '_' + i)
        column = 1
        worksheet3.write(row, column, importData)
        column = 2
        worksheet3.write(row, column, actionDescImport)
        column = 3
        worksheet3.write(row, column, 'TRUE')
        row += 1

if Module == 'Margin':

    for i in mainTables:
        column = 0
        worksheet3.write(row, column, imPort + '_' + i)
        column = 1
        worksheet3.write(row, column, importData)
        column = 2
        worksheet3.write(row, column, actionDescImport)
        column = 3
        worksheet3.write(row, column, 'TRUE')
        row += 1


for w in wList:
    column = 0
    worksheet3.write(row, column, w)
    column = 1
    worksheet3.write(row, column, 'WaitForProcess')
    column = 2
    worksheet3.write(row, column, 'Wait for some Time')
    column = 3
    worksheet3.write(row, column, 'TRUE')
    row += 1

if Module == 'Connectivity':

    for cases in CaseNameList:

        if 'Stg' in cases:

            column = 0
            worksheet3.write(row, column, captureOn + '_' + cases)
            column = 1
            worksheet3.write(row, column, snapshotCapture)
            column = 2
            worksheet3.write(row, column, actionDescCapture)
            column = 3
            worksheet3.write(row, column, 'TRUE')
            row += 1

    for cases in CaseNameList:

        if 'Stg' in cases:

            column = 0
            worksheet3.write(row, column, compare + '_' + m)
            column = 1
            worksheet3.write(row, column, snapshotCompare)
            column = 2
            worksheet3.write(row, column, actionDescCompare)
            column = 3
            worksheet3.write(row, column, 'TRUE')
            row += 1

if Module == 'Margin':

    for cases in CaseNameList:

        if 'Stg' not in cases:
            column = 0
            worksheet3.write(row, column, captureOn + '_' + cases)
            column = 1
            worksheet3.write(row, column, snapshotCapture)
            column = 2
            worksheet3.write(row, column, actionDescMainTableCapture)
            column = 3
            worksheet3.write(row, column, 'TRUE')
            row += 1

    for cases in CaseNameList:

        if 'Stg' not in cases:
            column = 0
            worksheet3.write(row, column, compare + '_' + cases)
            column = 1
            worksheet3.write(row, column, snapshotCompare)
            column = 2
            worksheet3.write(row, column, actionDescCompare)
            column = 3
            worksheet3.write(row, column, 'TRUE')
            row += 1

# SnapshotCaptureConfiguration

eodWhereCondition = 'SOURCESYSTEM=:SourceSystem and Businessdate = :Yesterdaydate ' \
                    'and  FeedName = :FeedName_Exp and FeedType = :FeedType_Exp'

itdWhereCondition = 'SOURCESYSTEM=:SourceSystem and Businessdate = :BusinessDate ' \
                    'and  FeedName = :FeedName_Exp and FeedType = :FeedType_Exp'

row = 1

if Module == 'Connectivity':

    for m in CaseNameList:
        if '_EOD' in m:
            column = 1
            worksheet4.write(row, column, captureOn + '_' + m)
            column = 2
            worksheet4.write(row, column, eodWhereCondition)
            row += 1
        if '_ITD' in m:
            column = 1
            worksheet4.write(row, column, captureOn + '_' + m)
            column = 2
            worksheet4.write(row, column, itdWhereCondition)
            row += 1

    row = 1

    for m in CaseNameList:
        if '_StgMarCall_' in m:
            column = 0
            worksheet4.write(row, column, stagingMarginCall)
            column = 3
            worksheet4.write(row, column, stgMarCallColumns)
            row += 1
        if '_StgColl_' in m:
            column = 0
            worksheet4.write(row, column, stagingCollateralBalance)
            column = 3
            worksheet4.write(row, column, stgCollColumns)
            row += 1
        if '_StgAlloc_' in m:
            column = 0
            worksheet4.write(row, column, stagingMarginCollateralAllocation)
            column = 3
            worksheet4.write(row, column, stgAllocColumns)
            row += 1

    row = 1

    for m in CaseNameList:
        column = 6
        for i in KeyColumnForStgMarCall:
            if '_StgMarCall_' in m:
                worksheet4.write(row, column, i)
                row += 1
        for i in KeyColumnForStgColl:
            if '_StgColl_' in m:
                worksheet4.write(row, column, i)
                row += 1
        for i in KeyColumnForStgAlloc:
            if '_StgAlloc_' in m:
                worksheet4.write(row, column, i)
                row += 1

if Module == 'Margin':

    for cases in CaseNameList:

        if '_Stg' not in cases:

            if '_EOD' in cases:

                column = 1
                worksheet4.write(row, column, captureOn + '_' + cases)
                row += 1

            if '_ITD' in cases:

                column = 1
                worksheet4.write(row, column, captureOn + '_' + cases)
                row += 1

    row = 1

    for cases in CaseNameList:

        if '_MarCall_' in cases:

            column = 0
            worksheet4.write(row, column, marginCall)
            column = 5
            worksheet4.write(row, column, MarCallQuery)
            row += 1

        if '_ColBal_' in cases:

            column = 0
            worksheet4.write(row, column, collateralBalance)
            column = 5
            worksheet4.write(row, column, ColBalQuery)
            row += 1

        if '_MCAlloc_' in cases:

            column = 0
            worksheet4.write(row, column, margincollateralAllocation)
            column = 5
            worksheet4.write(row, column, MCAllocQuery)
            row += 1

    row = 1

    for cases in CaseNameList:

        column = 6

        for i in KeyColumnFor_MarCall:

            if '_MarCall_' in cases:

                worksheet4.write(row, column, i)
                row += 1

        for i in KeyColumnFor_ColBal:

            if '_ColBal_' in cases:

                worksheet4.write(row, column, i)
                row += 1

        for i in KeyColumnFor_MCAlloc:

            if '_MCAlloc_' in cases:

                worksheet4.write(row, column, i)
                row += 1

# SnapshotCompareConfiguration

row = 1

if Module == 'Connectivity':

    for m in CaseNameList:

        if '_Stg' in m:

            if SourceSystem + '_' in m:
                column = 0
                worksheet5.write(row, column, compare + '_' + m)
                column = 2
                worksheet5.write(row, column, captureOn + '_' + m)
                row += 1

    row = 1

    for m in CaseNameList:
        if '_StgMarCall_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + stagingMarginCall)
            column = 5
            worksheet5.write(row, column, stgMarCallColumns)
            row += 1
        if '_StgColl_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + stagingCollateralBalance)
            column = 5
            worksheet5.write(row, column, stgCollColumns)
            row += 1
        if '_StgAlloc_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + stagingMarginCollateralAllocation)
            column = 5
            worksheet5.write(row, column, stgAllocColumns)
            row += 1


    row = 1

    for m in CaseNameList:
        column = 4
        for i in KeyColumnForStgMarCall:
            if 'StgMarCall' in m:
                worksheet5.write(row, column, i)
                row += 1
        for i in KeyColumnForStgColl:
            if 'StgColl' in m:
                worksheet5.write(row, column, i)
                row += 1
        for i in KeyColumnForStgAlloc:
            if 'StgAlloc' in m:
                worksheet5.write(row, column, i)
                row += 1



if Module == 'Margin':

    for m in CaseNameList:

        if '_Stg' not in m:

            if SourceSystem + '_' in m:
                column = 0
                worksheet5.write(row, column, compare + '_' + m)
                column = 2
                worksheet5.write(row, column, captureOn + '_' + m)
                row += 1

    row = 1

    for m in CaseNameList:
        if '_MarCall_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + marginCall)
            column = 5
            worksheet5.write(row, column, marginCallColumns)
            row += 1
        if '_ColBal_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + collateralBalance)
            column = 5
            worksheet5.write(row, column, colBalColumns)
            row += 1
        if '_MCAlloc_' in m:
            column = 1
            worksheet5.write(row, column, imPort + '_' + margincollateralAllocation)
            column = 5
            worksheet5.write(row, column, mcAllocColumns)
            row += 1


    row = 1

    for m in CaseNameList:
        column = 4
        for i in KeyColumnFor_MarCall:
            if '_MarCall_' in m:
                worksheet5.write(row, column, i)
                row += 1
        for i in KeyColumnFor_ColBal:
            if '_ColBal_' in m:
                worksheet5.write(row, column, i)
                row += 1
        for i in KeyColumnFor_MCAlloc:
            if '_MCAlloc_' in m:
                worksheet5.write(row, column, i)
                row += 1


# ImportDataConfiguration

row = 1

expectedData = 'ExpectedData'
inputFilePath = '$inputFilePath'
mTabName = 'MarginCall_TabName'
cTabName = 'CollateralBalance_TabName'
mcaTabName = 'MarginCollateralAllocation_TabName'

if Module == 'Connectivity':

    for table in stagingTables:
        column = 0
        worksheet6.write(row, column, imPort + '_' + table)
        column = 1
        worksheet6.write(row, column, expectedData)
        column = 2
        worksheet6.write(row, column, inputFilePath)
        if 'StagingMarginCall' in table:
            column = 3
            worksheet6.write(row, column, mTabName)
            row += 1
        if 'StagingCollateralBalance' in table:
            column = 3
            worksheet6.write(row, column, cTabName)
            row += 1
        if 'StagingMarginCollateralAllocation' in table:
            column = 3
            worksheet6.write(row, column, mcaTabName)
            row += 1

    row = 1

    for i in stagingTables:
        column = 5
        for m in KeyColumnForStgMarCall:
            if 'StagingMarginCall' in i:
                worksheet6.write(row, column, m)
                row += 1
        for m in KeyColumnForStgColl:
            if 'StagingCollateralBalance' in i:
                worksheet6.write(row, column, m)
                row += 1
        for m in KeyColumnForStgAlloc:
            if 'StagingMarginCollateralAllocation' in i:
                worksheet6.write(row, column, m)
                row += 1

if Module == 'Margin':

    for table in mainTables:
        column = 0
        worksheet6.write(row, column, imPort + '_' + table)
        column = 1
        worksheet6.write(row, column, expectedData)
        column = 2
        worksheet6.write(row, column, inputFilePath)
        if 'MarginCall' in table:
            column = 3
            worksheet6.write(row, column, mTabName)
            row += 1
        if 'CollateralBalance' in table:
            column = 3
            worksheet6.write(row, column, cTabName)
            row += 1
        if 'MarginCollateralAllocation' in table:
            column = 3
            worksheet6.write(row, column, mcaTabName)
            row += 1

    row = 1

    for i in mainTables:
        column = 5
        for m in KeyColumnFor_MarCall:
            if 'MarginCall' in i:
                worksheet6.write(row, column, m)
                row += 1
        for m in KeyColumnFor_ColBal:
            if 'CollateralBalance' in i:
                worksheet6.write(row, column, m)
                row += 1
        for m in KeyColumnFor_MCAlloc:
            if 'MarginCollateralAllocation' in i:
                worksheet6.write(row, column, m)
                row += 1




# AlterTableConfiguration

row = 1

for i in allTables:
    column = 0
    worksheet7.write(row, column, cleanUp + '_' + i)
    row += 1

row = 1

for i in deleteQueries:
    column = 1
    worksheet7.write(row, column, i)
    row += 1

# ExecuteScriptConfiguration

row = 1
column = 0
worksheet8.write(row, column, startScript)
column = 1
worksheet8.write(row, column, svcDirectory)
column = 2
worksheet8.write(row, column, scriptName)
column = 3
worksheet8.write(row, column, int(Instance))


# WaitForProcessConfiguration

row = 1

for w in wList:
    column = 0
    worksheet9.write(row, column, w)
    row += 1

row = 1

for x in mList:
    column = 1
    worksheet9.write(row, column, x)
    row += 1


OutputMsg = (Module + '_' + configData + '_' + SourceSystem + '.xlsx' + ' ' + 'has been created')

print(OutputMsg)

workbook.close()

# Workbook Completed
