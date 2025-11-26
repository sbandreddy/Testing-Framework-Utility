import pandas as pd
import xlsxwriter

input_SuiteData = pd.read_excel('CCP_InputFile.xlsx')

data_dict = {}

for index, row in input_SuiteData.iterrows():

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

key = 'CaseName'
if key in data_dict:
    CaseName = data_dict[key]

CaseNameList = list(CaseName)


def testSuiteName(Module, SourceSystem):
    testSuiteName1 = Module + '_' + SourceSystem
    return(testSuiteName1)


def caseName(CaseName):
    caseNameSheet1 = CaseName
    CaseName = ', '.join(map(str, CaseName))
    return caseNameSheet1


strDate = 'STR_BD'
bd = 'BD'
ybd = 'BD_1'
eod = 'EOD'
itd = 'ITD'
stgmargincall = 'StgMarginCall'
stgcollbal = 'StgCollBal'
stgalloc = 'StgAlloc'
margincall = 'MarginCall'
collateralbalance = 'CollateralBalance'
mcallocation = 'MarginCollateralAllocation'
inputfilepath = '/data/transcend/coreqaauto/COREQAAUTO_Automation/Connectivity_Margin/'
stgmcTabName = SourceSystem + '_' + stgmargincall
stgcbTabName = SourceSystem + '_' + stgcollbal
stgmcaTabName = SourceSystem + '_' + stgalloc
mcTabName = SourceSystem + '_' + margincall
cbTabName = SourceSystem + '_' + collateralbalance
mcaTabName = SourceSystem + '_' + mcallocation

suiteAndCaseData = 'SuiteAndCaseData'

workbook = xlsxwriter.Workbook(Module + '_' + suiteAndCaseData + '_' + SourceSystem + '.xlsx')

workbookName = Module + '_' + suiteAndCaseData + '_' + SourceSystem + '.xlsx'
inputWorkbookName = Module + '_' + 'inputData' + '_' + SourceSystem + '.xlsx'

suiteWorksheet1 = workbook.add_worksheet("TestSuiteInputData")
suiteWorksheet2 = workbook.add_worksheet("TestCaseInputData")

suiteHeader1 = ["TestSuiteExecutionId", "TestSuiteName", "stringdate", "Businessdate", "SourceSystem", "Yesterdaydate",
                "$inputFilePath", "ExpectedDate"]
suiteHeader2 = ["TestSuiteExecutionId", "TestSuiteName", "TestCaseName", "MarginCall_TabName",
                "CollateralBalance_TabName", "MarginCollateralAllocation_TabName", "SourceSystem",
                "FeedName_Exp", "AsOfBusinessDate", "FeedType_Exp"]

suiteWorksheets = [suiteWorksheet1, suiteWorksheet2]
suiteHeaders = [suiteHeader1, suiteHeader2]

for i in range(0, len(suiteWorksheets), 1):
    row = 0
    column = 0
    for header_column_name in suiteHeaders[i]:
        suiteWorksheets[i].write(row, column, header_column_name)
        column += 1

#                                         Worksheet 1

row = 1

if 'Connectivity' in Module:
    column = 0
    suiteWorksheet1.write(row, column, int(1))
    column = 1
    suiteWorksheet1.write(row, column, testSuiteName(Module, SourceSystem))
    column = 2
    suiteWorksheet1.write(row, column, strDate)
    column = 3
    suiteWorksheet1.write(row, column, bd)
    column = 4
    suiteWorksheet1.write(row, column, SourceSystem)
    column = 5
    suiteWorksheet1.write(row, column, ybd)
    column = 6
    suiteWorksheet1.write(row, column, inputfilepath + inputWorkbookName)
    column = 7
    suiteWorksheet1.write(row, column, bd)

if 'Margin' in Module:
    column = 0
    suiteWorksheet1.write(row, column, int(1))
    column = 1
    suiteWorksheet1.write(row, column, testSuiteName(Module, SourceSystem))
    column = 2
    suiteWorksheet1.write(row, column, strDate)
    column = 3
    suiteWorksheet1.write(row, column, bd)
    column = 4
    suiteWorksheet1.write(row, column, SourceSystem)
    column = 5
    suiteWorksheet1.write(row, column, ybd)
    column = 6
    suiteWorksheet1.write(row, column, inputfilepath + inputWorkbookName)
    column = 7
    suiteWorksheet1.write(row, column, bd)

# WorkSheet 2

row = 1

if 'Connectivity' in Module:

    for m in CaseNameList:

        column = 0
        suiteWorksheet2.write(row, column, int(1))
        column = 1
        suiteWorksheet2.write(row, column, testSuiteName(Module, SourceSystem))
        column = 2
        suiteWorksheet2.write(row, column, m)
        column = 3
        if '_StgMarCall_' in m:
            suiteWorksheet2.write(row, column, stgmcTabName)
        column = 4
        if '_StgColl_' in m:
            suiteWorksheet2.write(row, column, stgcbTabName)
        column = 5
        if '_StgAlloc_' in m:
            suiteWorksheet2.write(row, column, stgmcaTabName)
        column = 6
        suiteWorksheet2.write(row, column, SourceSystem)
        column = 9
        if '_Stg' in m:

            if '_EOD' in m:
                suiteWorksheet2.write(row, column, eod)
            if '_ITD' in m:
                suiteWorksheet2.write(row, column, itd)

            row += 1

if 'Margin' in Module:

    for m in CaseNameList:

        if '_Stg' not in m:

            column = 0
            suiteWorksheet2.write(row, column, int(1))
            column = 1
            suiteWorksheet2.write(row, column, testSuiteName(Module, SourceSystem))
            column = 2
            suiteWorksheet2.write(row, column, m)
            column = 3
            if '_MarCall_' in m:
                suiteWorksheet2.write(row, column, mcTabName)
            column = 4
            if '_ColBal_' in m:
                suiteWorksheet2.write(row, column, cbTabName)
            column = 5
            if '_MCAlloc_' in m:
                suiteWorksheet2.write(row, column, mcaTabName)
            column = 6
            suiteWorksheet2.write(row, column, SourceSystem)
            column = 9

            if '_EOD' in m:
                suiteWorksheet2.write(row, column, eod)
            if '_ITD' in m:
                suiteWorksheet2.write(row, column, itd)

            row += 1

    row = 1

    column = 8

    for cases in CaseNameList:

        if '_Stg' not in cases:

            if '_EOD' in cases:

                suiteWorksheet2.write(row, column, ybd)
                row += 1
            if '_ITD' in cases:

                suiteWorksheet2.write(row, column, bd)
                row += 1

OutputMsg = (Module + '_' + suiteAndCaseData + '_' + SourceSystem + '.xlsx' + ' ' + 'has been created')

print(OutputMsg)

workbook.close()
