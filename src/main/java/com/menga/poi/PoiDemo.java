package com.menga.poi;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author Marvel Cheng
 * @date 2020年04月15日
 */
public class PoiDemo {

    public static void main(String[] args) throws FileNotFoundException {
        PoiUtil.WorkbookData workbookData = createWorkbookData();
        FileOutputStream stream = new FileOutputStream("test.xls");
        PoiUtil.writeHSSFWorkbook(workbookData, stream);

        PoiUtil.WorkbookData workbookData2 = createWorkbookData2();
        FileOutputStream stream2 = new FileOutputStream("test2.xlsx");
        PoiUtil.writeSXSSFWorkbook(workbookData2, stream2);
        System.out.println("yes");
    }

    public static PoiUtil.WorkbookData createWorkbookData2() {
        PoiUtil.WorkbookData workbookData = new PoiUtil.WorkbookData();
        PoiUtil.SheetData sheetData = new PoiUtil.SheetData();

        workbookData.setSheetDataList(Collections.singletonList(sheetData));

        List<PoiUtil.TitleData> titleDataList = Lists.newArrayList();

        PoiUtil.TitleData titleData1 = new PoiUtil.TitleData().setValue("练一练").setXPos(1).setYPos(0).setWidth(10).setStyle(2);
        PoiUtil.TitleData titleData2 = new PoiUtil.TitleData().setValue("拍一拍").setXPos(11).setYPos(0).setWidth(10).setStyle(2);
        titleDataList.add(titleData1);
        titleDataList.add(titleData2);

        for (int i = 1; i <= 10; i++) {
            PoiUtil.TitleData titleData11 = new PoiUtil.TitleData().setValue(String.valueOf(i)).setXPos(i).setYPos(1);
            PoiUtil.TitleData titleData12 = new PoiUtil.TitleData().setValue(String.valueOf(i)).setXPos(i + 10).setYPos(1);
            titleDataList.add(titleData11);
            titleDataList.add(titleData12);
        }
        PoiUtil.TitleData titleData3 = new PoiUtil.TitleData().setValue("优良率").setXPos(21).setYPos(1);
        titleDataList.add(titleData3);
        sheetData.setTitleDataList(titleDataList);

        List<PoiUtil.RowData> rowDataList = IntStream.range(0, 10).mapToObj(ri -> {
            PoiUtil.RowData rowData = new PoiUtil.RowData();

            List<PoiUtil.CellData> cellDataList = IntStream.range(0, 22).mapToObj(ci -> {
                PoiUtil.CellData cellData = new PoiUtil.CellData();

                if (ci == 0) {
                    cellData.setValue("某某人");
                } else if (ci == 21) {
                    cellData.setValue("100%");
                } else if (ci == 2) {
                    cellData.setValue("A");
                } else if (ci == 5) {
                    cellData.setValue("错误");
                    cellData.setStyle(PoiUtil.RED_FONT_STYLE);
                } else {
                    cellData.setValue("C");
                }
                return cellData;
            }).collect(Collectors.toList());

            rowData.setCellDataList(cellDataList);
            return rowData;
        }).collect(Collectors.toList());

        PoiUtil.RowData rowData2 = new PoiUtil.RowData();
        List<PoiUtil.CellData> cellDataList2 = IntStream.range(0, 22).mapToObj(ci -> {
            PoiUtil.CellData cellData = new PoiUtil.CellData();
            if (ci == 0) {
                cellData.setValue("王锦锐");
            } else if (ci == 21) {
                cellData.setValue("0%");
            } else {
                cellData.setValue("未答");
            }
            return cellData;
        }).collect(Collectors.toList());
        rowData2.setCellDataList(cellDataList2);

        PoiUtil.RowData rowData3 = new PoiUtil.RowData();
        List<PoiUtil.CellData> cellDataList3 = IntStream.range(0, 21).mapToObj(ci -> {
            PoiUtil.CellData cellData = new PoiUtil.CellData();
            if (ci == 0) {
                cellData.setValue("优良率");
            } else {
                cellData.setValue("77%");
            }
            return cellData;
        }).collect(Collectors.toList());
        rowData3.setCellDataList(cellDataList3);

        rowDataList.add(rowData2);
        rowDataList.add(rowData3);

        sheetData.setRowDataList(rowDataList);

        return workbookData;
    }

    public static PoiUtil.WorkbookData createWorkbookData() {
        int sheetCount = 2;
        int rowCount = 5;
        int cellCount = 8;

        PoiUtil.WorkbookData workbookData = new PoiUtil.WorkbookData();

        List<PoiUtil.SheetData> sheetDataList = IntStream.range(0, sheetCount).mapToObj(si -> {
            PoiUtil.SheetData sheetData = new PoiUtil.SheetData();
            sheetData.setName("表" + si);

            List<PoiUtil.RowData> rowDataList = IntStream.range(0, rowCount).mapToObj(ri -> {
                PoiUtil.RowData rowData = new PoiUtil.RowData();

                List<PoiUtil.CellData> cellDataList = IntStream.range(0, cellCount).mapToObj(ci -> {
                    PoiUtil.CellData cellData = new PoiUtil.CellData();
                    cellData.setValue("数据哈（" + ri + "," + ci + ")");
                    if (2 == ci) {
                        cellData.setStyle(PoiUtil.RED_FONT_STYLE);
                    }
                    return cellData;
                }).collect(Collectors.toList());

                rowData.setCellDataList(cellDataList);
                return rowData;
            }).collect(Collectors.toList());

            sheetData.setRowDataList(rowDataList);
            return sheetData;
        }).collect(Collectors.toList());

        workbookData.setSheetDataList(sheetDataList);
        return workbookData;
    }

    /**
     * 测试写入 Excel 表
     *
     * @param stream
     */
    public static void testWriteFile(OutputStream stream) {
        // 1 创建Excel工作文件对象
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            // 2 根据文件对象创建表格对象
            HSSFSheet sheet = workbook.createSheet("wei");
            // 3 根据表格对象创建表格的行对象
            HSSFRow row = sheet.createRow(0);
            // 4 根据行对象创建表格的单元格对象
            HSSFCell cell = row.createCell(0);
            // 5 往指定的位置插入数据
            cell.setCellValue("哈2");
            // 6 将数据以流的方式存储到文件中
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
