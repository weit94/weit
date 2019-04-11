# pip install XlsxWriter python-levenshtein pandas xlwt xlrd openpyxl Levenshtein azure-storage-file rope
import os
import pandas
from collections import defaultdict
import Levenshtein
import time
import datetime
import multiprocessing as mp
import xlsxwriter
import numpy


def nameToNameList(name, ignoreList):
    nameList = []
    leftP = name.find("（")
    rightP = name.find("）")
    if leftP > 0 and rightP > leftP:
        altName = name[leftP+1:rightP]
        if altName not in ignoreAltName:
            nameList.extend(altName.split("，"))
        nameList.append(name[:leftP] + name[rightP+1:len(name)])
    else:
        nameList.append(name)
    return nameList


#approvalFilename = r'C:/Users/pengf/OneDrive/Python/Order/testImported.xls'
# approvalFolder = 'C:/Users/pengf/OneDrive/Python/Order/'
approvalFolder = 'D:/pycharm_pro/'
xlsxSuffix = ".xlsx"
xlsSuffix = ".xls"
approvalFilename = '采购申请办公室'
outputName1 = approvalFolder + approvalFilename + '-updated1' + xlsxSuffix
outputName2 = approvalFolder + approvalFilename + '-updated2' + xlsxSuffix

# materialFilename = 'C:/Users/pengf/OneDrive/Python/Order/物料信息.xls'
materialFilename = 'D:/pycharm_pro/物料信息.xls'
#materialFilename = 'C:/Users/pengf/OneDrive/Python/Order/materialAllTiny.xls'

approvalFilenameFull = approvalFolder + approvalFilename + xlsSuffix
approval = pandas.read_excel(approvalFilenameFull, converters={
                             '审批编号': str, '期望交付日期': str})
material = pandas.read_excel(materialFilename)
ignoreCatalogName = ["水管配件"]
ignoreAltName = ["非泵送", "泵送", "86型", "25A", "63A，380V"]
detailNames = ["物料类别", "物料名称", "规格", "型号", "数量", "单位", "匹配度"]
replacements = {"*":"x"}
importantWords = ["45度", "KBG", "325", "425", "线槽", "硅胶"]
includeOldItem = True
t0 = time.time()

approvalValues = approval.values
materialValues = material.values
approvalColumns = approval.columns
shortColumns = (approvalColumns[0] == '审批编号')
detailStartIndex = 10
detailEndIndex = 16
detailEndIndexOld = 15
finalColumnCount = 26
materialNameCol = 11
SpecNameCol = 12
CountColOld = 14
if(shortColumns):
    detailStartIndex = 7
    detailEndIndex = 13
    detailEndIndexOld = 10
    finalColumnCount = 21
    materialNameCol = 7
    SpecNameCol = 8
    CountColOld = 9

# build the cell value matrix without column names
approvalValuesSorted = []
for row in range(1, len(approvalValues)):
    rowValue = []
    for col in range(0, len(approvalValues[row])):
        rowValue.append(approvalValues[row][col])
    approvalValuesSorted.append(rowValue)

# Collect the seperating rows before sorting
nonEmptyRow = []
for row in range(0, len(approvalValuesSorted)):
    if(approvalValuesSorted[row][0] == approvalValuesSorted[row][0]):
        nonEmptyRow.append(row)
nonEmptyRow.append(len(approvalValuesSorted))


def sortBlocks(approvalValuesToSort, startRow, startCol, endRow, endCol, sortIndexCol, secondSortIndexCol):
    blockToSort = []
    for i in range(startRow, endRow):
        blockToSort.append(approvalValuesSorted[i][startCol:endCol+1])
    sortedBlock = sorted(blockToSort, key=lambda x: (
        x[sortIndexCol-startCol], x[secondSortIndexCol-startCol]))
    for i in range(0, len(blockToSort)):
        for j in range(0, len(blockToSort[i])):
            approvalValuesToSort[startRow +
                                 i][startCol + j] = sortedBlock[i][j]
    return sortedBlock


# sorting the material in each block
for i in range(0, len(nonEmptyRow)-1):
    sortBlocks(approvalValuesSorted, nonEmptyRow[i], detailStartIndex,
               nonEmptyRow[i+1], detailEndIndexOld, materialNameCol, SpecNameCol)

# get the complete meterial names for performance opetimization
completNameAll = []
for i in range(0, len(materialValues)):
    name = materialValues[i][3]
    spec = materialValues[i][4]
    model = materialValues[i][5]
    if "办公桌" in name:
        print(name)
    completNameAll.append(name + spec + model)
t1 = time.time()

def CalcBestMatched(i):
    toCheck = ""
    alternative = []
    length = 4
    if(shortColumns):
        length = 2
    for j in range(detailStartIndex, detailStartIndex+length):
        print(approvalValuesSorted[i][j])
        if(approvalValuesSorted[i][j] == approvalValuesSorted[i][j] and approvalValuesSorted[i][j] != "/" and approvalValuesSorted[i][j] != "无"):
            toCheck += approvalValuesSorted[i][j]
    print("to Check {} {}".format(i, toCheck))
    mostMatched = defaultdict(list)

    compareCount = 0
    t1 = time.time()
    for ii in range(0, len(materialValues)):
        catagoryList = []
        nameList = []
        catagory = materialValues[ii][3]
        name = materialValues[ii][4]
        spec = materialValues[ii][5]
        if len(set(completNameAll[ii]) & set(toCheck)) == 0:
            continue
        AllAltNames = []
        if catagory not in name and catagory not in ignoreCatalogName:
            catagoryList = nameToNameList(catagory, [])
        else:
            catagoryList.append("")
        if name != "/":
            nameList = nameToNameList(name, ignoreAltName)
        else:
            nameList.append("")
        if spec == "/" or spec == '无':
            spec = ""
        for c in catagoryList:
            for n in nameList:
                compareCount += 1
                fullName = "{}{}{}".format(c, n, spec)
                distanceRatio = Levenshtein.ratio(toCheck.lower(), fullName.lower())
                for w in importantWords:
                    if w in fullName and w in toCheck :
                        distanceRatio = distanceRatio + 0.14 #if important words appear in both words, distance ratio will be increased.
                mostMatched[distanceRatio].append(ii)
    t2 = time.time()
    print("comparing count {}, time {} seconds".format(
        compareCount, int(t2 - t1)))
    if not mostMatched:
        return [-1, 0]
    sortedMostMatched = sorted(mostMatched, reverse=True)
    key = sortedMostMatched[0]
    inde = mostMatched[key][0]
    print("{}: {} {}".format(completNameAll[inde], toCheck, key))
    for j in range(0, min(10, len(sortedMostMatched))):
        keyJ = sortedMostMatched[j]
        print("{}: {} {}".format(completNameAll[mostMatched[keyJ][0]], toCheck, keyJ))
    return [inde, key]


# Find the most matched item
bestIndexList = []
for i in range(0, len(approvalValuesSorted)):
    inde = CalcBestMatched(i)
    bestIndexList.append(inde)

# Calculate the column name
finalColumn = []
for col in range(0, detailStartIndex):
    finalColumn.append(approvalColumns[col])
for col in range(0, len(detailNames)):
    finalColumn.append(detailNames[col])
for col in range(detailEndIndexOld+1, len(approvalColumns)):
    finalColumn.append(approvalColumns[col])


def writeToExcel(outputName, includeOldItem):
    workbook = xlsxwriter.Workbook(outputName)
    worksheet = workbook.add_worksheet()

    # Fill the column name
    for col in range(0, len(finalColumn)):
        val = finalColumn[col]
        worksheet.write_string(0, col, val)
        worksheet.write_string(1, col, val)

    # Calculate the cell values
    finalValue = []
    for row in range(0, len(approvalValuesSorted)):
        # appending the original row
        if(includeOldItem):
            originalRowValues = []
            for col in range(0, detailStartIndex):
                val = approvalValuesSorted[row][col]
                originalRowValues.append(val)
            originalRowValues.append(numpy.NaN)  # 物料类别
            originalRowValues.append(
                approvalValuesSorted[row][materialNameCol])  # 物料名称
            originalRowValues.append(
                approvalValuesSorted[row][SpecNameCol])  # 规格
            originalRowValues.append(numpy.NaN)  # 型号
            originalRowValues.append(
                approvalValuesSorted[row][CountColOld])  # 数量
            originalRowValues.append(
                approvalValuesSorted[row][CountColOld+1])  # 单位
            originalRowValues.append(numpy.NaN)  # 匹配度
            for col in range(detailEndIndexOld+1, len(approvalColumns)):
                originalRowValues.append(approvalValuesSorted[row][col])
            finalValue.append(originalRowValues)
        # appending the updated row
        rowValues = []
        for col in range(0, detailStartIndex):
            val = numpy.NaN if includeOldItem else approvalValuesSorted[row][col]
            rowValues.append(val)
        bestI = bestIndexList[row][0]
        if (bestI >= 0):
            rowValues.append(materialValues[bestI][2])  # 物料类别
            rowValues.append(materialValues[bestI][3])  # 物料名称
            rowValues.append(materialValues[bestI][4])  # 规格
            rowValues.append(materialValues[bestI][5])  # 型号
        else:
            rowValues.append("")
            rowValues.append("")
            rowValues.append("")
            rowValues.append("")
            
        rowValues.append(approvalValuesSorted[row][CountColOld])  # 数量
        rowValues.append(approvalValuesSorted[row][CountColOld+1])  # 单位
        rowValues.append(bestIndexList[row][1])  # 匹配度
        for col in range(detailEndIndexOld+1, len(approvalColumns)):
            val = numpy.NaN if includeOldItem else approvalValuesSorted[row][col]
            rowValues.append(val)
        finalValue.append(rowValues)

    # Filling the cell values
    for row in range(0, len(finalValue)):
        for col in range(0, len(finalValue[0])):
            val = finalValue[row][col]
            if(val == val and val != 'nan'):
                worksheet.write_string(row+2, col, str(val))
            else:
                worksheet.write_string(row+2, col, '')

    # Collect the rows before merging cells
    nonEmptyRow = []
    for row in range(0, len(finalValue)):
        if(finalValue[row][0] == finalValue[row][0]):
            nonEmptyRow.append(row)
    nonEmptyRow.append(len(finalValue))

    # Merge the cells
    worksheet.merge_range(0, detailStartIndex, 0, detailEndIndex, "采购明细")
    for col in range(0, len(finalColumn)):
        if(col < detailStartIndex or col > detailEndIndex):
            worksheet.merge_range(0, col, 1, col, finalColumn[col])
            for i in range(0, len(nonEmptyRow)-1):
                val = str(finalValue[nonEmptyRow[i]][col])
                if(val == 'nan'):
                    val = ''
                if(nonEmptyRow[i] + 1 != nonEmptyRow[i+1]):  # not a single cell
                    worksheet.merge_range(
                        nonEmptyRow[i]+2, col, nonEmptyRow[i+1]+1, col, val)

    workbook.close()

writeToExcel(outputName1, False)
writeToExcel(outputName2, True)
