import xlrd, xlwt,openpyxl
from openpyxl.styles import PatternFill
from xlutils.styles import Styles
import math,os

ORIGINAL_READROAD = "C:\\Users\\Administrator\\Desktop\\内部進捗（源文件）.xlsx";
CONTRAST_READROAD = "C:\\Users\\Administrator\\Desktop\\内部進捗（目标文件）.xlsx";
OUTPUTROAD = "";
START_COL = ""
END_COL = "DQ";


loadWorkBook = openpyxl.load_workbook(ORIGINAL_READROAD);
loadSheet = loadWorkBook["進捗明細"];

#输入第几列，返回其对应的英文序列号，col起始位为1
def numToEn(col):
    col = int(col);
    enList = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
              "V", "W", "X", "Y", "Z"];
    if col <= 26:
        return enList[col - 1];
    elif col > 26 and col <= 702:
        if(col == 702):
            return "ZZ";
        ex = enList[math.floor(col / 26) - 1];
        remain = enList[col % 26 - 1];
        return ex +remain;
    elif col > 702:
        ex_first = enList[math.floor(math.floor(col / 26) / 26) - 1];
        ex_second = enList[math.floor(col / 26) % 26 - 1];
        remain = enList[col % 26 - 1];
        return ex_first + ex_second + remain;


#遍历第三行，将单元格地址与背景色按键值对存储
def cellColor():
    colList = [];
    for col in range(len(tuple(loadSheet.columns))):
        cellIndex = numToEn(col + 1) + "3";
        colList.append(cellIndex);
    colDict = dict.fromkeys(colList);
    for index in range(len(colList)):
        colorIndex = loadSheet[colList[index]].fill.start_color.index;
        if type(colorIndex) == type(1):
            colDict[colList[index]] = openpyxl.styles.colors.COLOR_INDEX[colorIndex];
        elif type(colorIndex) == type(""):
            colDict[colList[index]] = colorIndex;
    return colDict;

print(cellColor());













# originalBook = xlrd.open_workbook(ORIGINAL_READROAD);
# originalSheet = originalBook.sheets()[0];#第一个sheet
# o_rows = originalSheet.nrows;
# o_cols = originalSheet.ncols;

# book = xlrd.open_workbook(ORIGINAL_READROAD,formatting_info=1);
# sheet = book.sheets()[0];
# s = Styles(book);
# print(s[sheet.cell(0,0)])


# for row in range(o_cols):
#     print(originalSheet.col_values(row));
#
# contrastBook = xlrd.open_workbook(CONTRAST_READROAD);
# contrastSheet = contrastBook.sheets()[0];
# for row in range(o_rows):
#     print("第" + str(row + 1) + "行", end = "");
#     for col in range(o_cols):
#         if originalSheet.cell_type(row,col) == 0:
#             print("\t空的",end = "");
#         else:
#             print("\t", end = "");
#             print(originalSheet.cell_value(row,col),end = "");
#             print("\t",end = "");
#     print()

# list = [];
# for row in range(o_rows):
#     print("第" + str(row + 1) + "行", end = "");
#     for col in range(o_cols):
#         if originalSheet.cell_type(row,col) == 0:
#             print("\t空的",end = "");
#         else:
#             print("\t", end = "");
#             print(originalSheet.cell_value(row,col) == contrastSheet.cell_value(row,col),end = "");
#             if(originalSheet.cell_value(row ,col ) != contrastSheet.cell_value(row,col)):
#                 temp = (row + 1,col + 1);
#                 list.append(temp);
#             # print("\t",end = "");
#     print()

# loadWorkBook = openpyxl.load_workbook(ORIGINAL_READROAD);
# loadSheet = loadWorkBook["進捗明細"];
# fill = PatternFill(fill_type='solid',fgColor='FFFF0000');
# print(list);
# for i in range(len(list)):
#     # print(EnList[((list[i])[1] - 1)] + str((list[i])[0]),end = "");
#     cell = loadSheet[EnList[((list[i])[1] - 1)] + str((list[i])[0])];
#     cell.fill = fill;
#
#
# loadWorkBook.save(CONTRAST_READROAD);

# print(openpyxl.styles.colors.COLOR_INDEX[loadSheet["AC3"].fill.start_color.index])
# print(loadSheet["P3"].fill.start_color.index)
# enList = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
#           "V", "W", "X", "Y", "Z"];
# for index in range(49):
#     print("第" + str(index + 1) + "列    ",end = "");
#     c = numToEn(index+1)+"2";
#     cx = loadSheet[c].fill.start_color.index;
#     if type(cx) == type(1):
#         print(openpyxl.styles.colors.COLOR_INDEX[cx]);
#     elif type(cx) == type(""):
#         print(cx);


