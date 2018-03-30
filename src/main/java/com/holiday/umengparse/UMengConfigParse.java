package com.holiday.umengparse;

import com.holiday.umengparse.model.MergeCell;
import com.holiday.umengparse.utils.FileAppend;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuan on 2018/3/29.
 */
public class UMengConfigParse {

    public static void main(String args[]) throws IOException, InvalidFormatException {
        System.out.println("hello1");

//        testMergedRegion();
//        testReadCell();
//        testRowNum();
//        testSplit();
        String filePath = "/Users/yuan/dev/product/00需求资料/新单词友盟统计V1.0.xlsx";
        String sheetName = "新单词友盟统计";

        new UMengConfigParse().generateUmengInfo(filePath, sheetName);

        System.out.println("hello2");
    }

    public void generateUmengInfo(String filePath, String sheetName) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(
                new FileInputStream(filePath));
        Sheet sheet = wb.getSheet(sheetName);
        List<MergeCell> moduleMergeCellList = getMergeCellList(sheet, 1, 0);

        p("------------------------模块包含以下部分--------------------------");
        //print log
        logMergeList(moduleMergeCellList);

        List<MergeCell> pageMergeCellList = getMergeCellList(sheet, 1, 1);

        p("------------------------页面路径包含以下部分--------------------------");
        //print log
        logMergeList(pageMergeCellList);

        childBindToParent(pageMergeCellList, moduleMergeCellList);

        p("------------------------页面绑定到模块后页面路径为--------------------------");
        //print log
        List<String> pagePathList = TraverseAllEventPath(moduleMergeCellList, "zn");
        logPath(pagePathList);

        List<MergeCell> eventMergeCellList = getEventCellList(sheet, 1, 2);

        p("------------------------具体事件包含以下部分--------------------------");
        logMergeList(eventMergeCellList);

        childBindToParent(eventMergeCellList, pageMergeCellList);
        p("------------------------事件绑定到页面后，事件路径为--------------------------");
        //print log
        List<String> eventPathList = TraverseAllEventPath(moduleMergeCellList, "zn");
        logPath(eventPathList);

        String pre = "ZN";
        String zone = "000000";
        generateAndroidFile(eventPathList, pre, zone);
        generateUmengFile(eventPathList, pre, zone);
    }

    private void generateUmengFile(List<String> eventPathList, String pre, String zone) {
        String androidFilePath = "/Users/yuan/Desktop/tmp/umeng.txt";
        for (int index = 0;index < eventPathList.size();index++) {
            String varName = getVarName(pre, zone, index);
            String annotation = eventPathList.get(index);
            String sourceLineTemplate = "%s,%s,0";
            String sourceLine = String.format(sourceLineTemplate, varName, annotation) + "\n";
            FileAppend.append1(androidFilePath, sourceLine);
        }
    }

    private void generateAndroidFile(List<String> eventPathList, String pre, String zone) {
        String androidFilePath = "/Users/yuan/Desktop/tmp/android.txt";
        for (int index = 0;index < eventPathList.size();index++) {
            String annotation = "   //" + eventPathList.get(index) + "\n";
            FileAppend.append1(androidFilePath, annotation);

            String sourceLineTemplate = "   public static final %s = \"%s\";";
            String varName = getVarName(pre, zone, index);
            String sourceLine = String.format(sourceLineTemplate, varName, varName) + "\n";
            FileAppend.append1(androidFilePath, sourceLine);
        }
    }

    public String getVarName(String pre, String zone, int index) {
        return pre + zone + index;
    }

    private void logPath(List<String> eventPathList) {
        for (String path : eventPathList) {
            System.out.println(path);
        }
    }

    private void logMergeList(List<MergeCell> pageMergeCellList) {
        for (MergeCell mergeCell : pageMergeCellList) {
            p(mergeCell.content + "," + mergeCell.mergeRegion.formatAsString());
        }
    }

    private List<MergeCell> getEventCellList(Sheet sheet, int startRow, int column) {
        List<MergeCell> mixedEventCellList = getMergeCellList(sheet, startRow, column);
        List<MergeCell> eventCellList = new ArrayList<>();
        for (MergeCell mixedCell : mixedEventCellList) {
            splitMixedCell(eventCellList, mixedCell);
        }
        return eventCellList;
    }

    private void splitMixedCell(List<MergeCell> eventCellList, MergeCell mixedCell) {
        String events[] = mixedCell.content.split("、|/|,|，");
        for (String event : events) {
            MergeCell mergeCell = new MergeCell();
            eventCellList.add(mergeCell);

            mergeCell.content = event;
            mergeCell.mergeRegion = mixedCell.mergeRegion;
        }
    }

    // TODO: 2018/3/30 改进遍历算法
    private List<String> TraverseAllEventPath(List<MergeCell> moduleMergeCellList, String parentPath) {
        List<String> eventPathList = new ArrayList<>();
        String separator = "_";

        for (MergeCell moduleCell : moduleMergeCellList) {
            if (moduleCell.children.size() > 0) {
                addPathBehindModule(eventPathList, separator, getPath(separator, parentPath, moduleCell.content), moduleCell);
            } else {
                eventPathList.add(getPath(separator, parentPath, moduleCell.content));
            }
        }

        return eventPathList;
    }

    private void addPathBehindModule(List<String> eventPathList, String separator, String parentPath, MergeCell moduleCell) {
        for (MergeCell pageCell : moduleCell.children) {
            if (pageCell.children.size() > 0) {
                addPathBehindPage(eventPathList, separator, getPath(separator, parentPath, pageCell.content), pageCell);
            } else {
                eventPathList.add(getPath(separator, parentPath, pageCell.content));
            }
        }
    }

    private String getPath(String separator, String parentPath, String name) {
        if (isEmptyStr(parentPath)) {
            return name;
        } else {
            return parentPath + separator + name;
        }
    }

    private void addPathBehindPage(List<String> eventPathList, String separator, String parentPath, MergeCell pageCell) {
        for (MergeCell eventCell : pageCell.children) {
            eventPathList.add(getPath(separator, parentPath, eventCell.content));
        }
    }

    private void childBindToParent(List<MergeCell> childMergeCellList, List<MergeCell> parentMergeCellList) {
        for (MergeCell pageMerge : childMergeCellList) {
            int frontRowIndex = pageMerge.mergeRegion.getFirstRow();
            int frontColumnIndex = pageMerge.mergeRegion.getFirstColumn() - 1;
            MergeCell parentMergeCell = findMergeCellContains(parentMergeCellList, frontRowIndex, frontColumnIndex);
            parentMergeCell.children.add(pageMerge);
        }
    }

    private MergeCell findMergeCellContains(List<MergeCell> moduleMergeCellList, int frontRowIndex, int frontColumnIndex) {
        for (MergeCell mergeCell : moduleMergeCellList) {
            if (mergeCell.mergeRegion.isInRange(frontRowIndex, frontColumnIndex)) {
                return mergeCell;
            }
        }

        return null;
    }

    private List<MergeCell> getMergeCellList(Sheet sheet, int startRow, int column) {
        List<MergeCell> mergeCellList = new ArrayList<>();
        int rowIndex = startRow;

        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();
        while (rowIndex <= sheet.getLastRowNum()) {
            Cell cell = sheet.getRow(rowIndex).getCell(column);
            addCellToMergeCell(mergeCellList, cellRangeAddressList, cell);
            rowIndex++;
        }
        return mergeCellList;
    }

    private void addCellToMergeCell(List<MergeCell> mergeCellList, List<CellRangeAddress> cellRangeAddressList, Cell cell) {
        CellRangeAddress mergeRegionAssociateWithCell = checkCellInMergeRgion(cellRangeAddressList, cell);
        boolean isCellAtMergeRegion = mergeRegionAssociateWithCell != null;
        MergeCell mergeCell = null;
        if (isCellAtMergeRegion) {
            p("cell." + cell.getAddress().toString() + "is in merge Area." + mergeRegionAssociateWithCell.toString());
            mergeCell = findMergeCell(mergeCellList, mergeRegionAssociateWithCell);
            if (mergeCell == null) {
                mergeCell = generateMergeCell(mergeRegionAssociateWithCell);
                mergeCellList.add(mergeCell);
            }
        } else {
            p("cell." + cell.getAddress().toString() + "is a single cell");
            mergeCell = generateMergeCell(cell);
            mergeCellList.add(mergeCell);
        }
        addContentToMergeCellIfExist(cell, mergeCell);
    }

    private void p(String s) {
        System.out.println(s);
    }

    private void addContentToMergeCellIfExist(Cell cell, MergeCell mergeCell) {
        String content = cell.getStringCellValue();
        if (!isEmptyStr(content)) {
            mergeCell.content = content;
        }
    }

    private boolean isEmptyStr(String content) {
        return (content == "") || (content == null);
    }

    private MergeCell generateMergeCell(CellRangeAddress mergeRegionAssociateWithCell) {
        MergeCell mergeCell = new MergeCell();
        mergeCell.mergeRegion = mergeRegionAssociateWithCell;
        return mergeCell;
    }

    private MergeCell generateMergeCell(Cell cell) {
        MergeCell mergeCell = new MergeCell();
        int firstRow = cell.getRowIndex();
        int lastRow = cell.getRowIndex();
        int firstColumn = cell.getColumnIndex();
        int lastColumn = cell.getColumnIndex();
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn);
        mergeCell.mergeRegion = cellRangeAddress;
        return mergeCell;
    }

    private MergeCell findMergeCell(List<MergeCell> mergeCellList, CellRangeAddress mergeRegionAssociateWithCell) {
        for (MergeCell mergeCell : mergeCellList) {
            if (mergeRegionAssociateWithCell.equals(mergeCell.mergeRegion)) {
                return mergeCell;
            }
        }

        return null;
    }

    private CellRangeAddress checkCellInMergeRgion(List<CellRangeAddress> cellRangeAddressList, Cell cell) {
        for (CellRangeAddress cellRangeAddress : cellRangeAddressList) {
            if (cellRangeAddress.isInRange(cell)) {
                return cellRangeAddress;
            }
        }

        return null;
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //                                                    test
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    private static void testMergedRegion(String filePath, String sheetName) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(
                new FileInputStream(filePath));
        Sheet sheet = wb.getSheet(sheetName);
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i); //Region of merged cells

            int colIndex = region.getFirstColumn(); //number of columns merged
            int rowNum = region.getFirstRow();      //number of rows merged
            System.out.println(region.formatAsString());
        }
    }

    private static void testReadCell() throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(
                new FileInputStream("/Users/yuan/dev/product/00需求资料/新单词友盟统计V1.0.xlsx"));
        Sheet sheet = wb.getSheet("新单词友盟统计");
        System.out.println("firstRowNum = " + sheet.getFirstRowNum() + ", lastRowNum = " + sheet.getLastRowNum());
        for (Row row : sheet) {
            Cell cell = row.getCell(0);
            if (cell != null) {
                String content = cell.getStringCellValue();
                System.out.println(content + "，row = " + cell.getRowIndex() + ", columnIndex = " + cell.getColumnIndex());
            }
        }
    }

    private static void testRowNum() throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(
                new FileInputStream("/Users/yuan/dev/product/00需求资料/新单词友盟统计V1.0.xlsx"));
        Sheet sheet = wb.getSheet("新单词友盟统计");
        System.out.println("firstRowNum = " + sheet.getFirstRowNum() + ", lastRowNum = " + sheet.getLastRowNum());
    }

    private static void testSplit() {
        String str = "页面进入次数、返回按钮/音频icon,给点提示点击次数";
        String strs[] = str.split("、|/|,");
        for (String s : strs) {
            System.out.println(s);
        }
    }
}
