package com.menga.poi;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Apache POI 工具类
 *
 * @author Marvel Cheng
 * @date 2020年04月15日
 */
public class PoiUtil {
    /**
     * 默认样式
     */
    public static final Integer DEFAULT_FONT_STYLE = 0;
    /**
     * 红色字体
     */
    public static final Integer RED_FONT_STYLE = 1;
    /**
     * 居中文本
     */
    public static final Integer CENTER_FONT_STYLE = 2;

    /**
     * 工作簿
     */
    private Workbook workbook;
    /**
     * 单元格样式
     */
    private Map<Integer, CellStyle> cellStyleMap = new HashMap<>();

    public PoiUtil(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 工作簿数据类
     */
    @Data
    public static class WorkbookData {
        private List<SheetData> sheetDataList;
    }

    /**
     * 工作表数据类
     */
    @Data
    public static class SheetData {
        private String name;
        // 表头数据
        private List<TitleData> titleDataList;
        // 行数据
        private List<RowData> rowDataList;
    }

    /**
     * 表头数据类
     */
    @Data
    @Accessors(chain = true)
    public static class TitleData {
        /**
         * 单元格字符串值
         */
        private String value;
        /**
         * 样式, 0:默认; 1:红色字体; 2:居中文本;
         */
        private Integer style = 0;
        /**
         * 单元格列坐标
         */
        private Integer xPos;
        /**
         * 单元格行坐标
         */
        private Integer yPos;
        /**
         * 单元格占用的列数
         */
        private Integer width = 1;
        /**
         * 单元格占用的行数
         */
        private Integer height = 1;
    }

    /**
     * 行数据类
     */
    @Data
    public static class RowData {
        private List<CellData> cellDataList;
    }

    /**
     * 单元格数据类
     */
    @Data
    public static class CellData {
        /**
         * 单元格字符串值
         */
        private String value;
        /**
         * 样式, 0:默认; 1:红色字体; 2:居中文本;
         */
        private Integer style = 0;
    }

    /**
     * 写入 Excel 表工作簿，Excel2003 以前（包括2003）的版本，扩展名是 .xls
     *
     * @param workbookData 工作簿数据
     * @param stream       输出流
     */
    public static void writeHSSFWorkbook(WorkbookData workbookData, OutputStream stream) {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            // 写入工作表数据
            new PoiUtil(workbook).writeSheet(workbookData.getSheetDataList());
            // 写入流
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 写入 Excel 表工作簿，Excel2007 的版本，扩展名是 .xlsx，且是一种基于 XSSF 的低内存占用的 API
     *
     * @param workbookData 工作簿数据
     * @param stream       输出流
     */
    public static void writeSXSSFWorkbook(WorkbookData workbookData, OutputStream stream) {
        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            // 写入工作表数据
            new PoiUtil(workbook).writeSheet(workbookData.getSheetDataList());
            // 写入流
            workbook.write(stream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 写入工作表数据
     *
     * @param sheetDataList 工作表数据列表
     */
    private void writeSheet(List<SheetData> sheetDataList) {
        for (SheetData sheetData : sheetDataList) {
            Sheet sheet;
            if (sheetData.getName() == null) {
                sheet = workbook.createSheet();
            } else {
                sheet = workbook.createSheet(sheetData.getName());
            }
            List<TitleData> titleDataList = sheetData.getTitleDataList();
            // 写入表头数据
            writeTitle(sheet, titleDataList);
            // 获取行数据的起始行数，不存在表头数据时直接在第0行开始写，否则从新的一行开始写
            int rowNum = (titleDataList == null || titleDataList.size() == 0) ? 0 : sheet.getLastRowNum() + 1;
            // 写入行数据
            writeRow(sheet, sheetData.getRowDataList(), rowNum);
        }
    }

    /**
     * 写入表头数据
     *
     * @param sheet         工作表
     * @param titleDataList 表头数据列表
     */
    private void writeTitle(Sheet sheet, List<TitleData> titleDataList) {
        if (titleDataList == null) {
            return;
        }
        for (TitleData titleData : titleDataList) {
            Row row = getRow(sheet, titleData.getYPos());
            Cell cell = getCell(row, titleData.getXPos());
            CellStyle cellStyle = getCellStyle(titleData.getStyle());

            if (titleData.getHeight() > 1 || titleData.getWidth() > 1) {
                int firstRow = row.getRowNum();
                int lastRow = firstRow + titleData.getHeight() - 1;
                int firstCol = cell.getColumnIndex();
                int lastCol = firstCol + titleData.getWidth() - 1;

                CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
                // 合并单元格
                sheet.addMergedRegion(region);
                // 设置区域样式
                setRegionStyle(sheet, region, cellStyle);
            } else {
                cell.setCellStyle(cellStyle);
            }

            cell.setCellValue(titleData.getValue());
        }
    }

    /**
     * 写入行数据
     *
     * @param sheet       工作表
     * @param rowDataList 行数据列表
     * @param rowNum      起始行号
     */
    private void writeRow(Sheet sheet, List<RowData> rowDataList, int rowNum) {
        for (RowData rowData : rowDataList) {
            Row row = sheet.createRow(rowNum++);
            writeCell(row, rowData.getCellDataList());
        }
    }

    /**
     * 写入单元格数据
     *
     * @param row          行
     * @param cellDataList 单元格数据列表
     */
    private void writeCell(Row row, List<CellData> cellDataList) {
        for (CellData cellData : cellDataList) {
            // 注意，getLastCellNum 是以 1 为起始值，并且没有单元格时返回 -1
            int cellNum = Math.max(row.getLastCellNum(), 0);
            Cell cell = row.createCell(cellNum);

            CellStyle cellStyle = getCellStyle(cellData.getStyle());
            cell.setCellStyle(cellStyle);
            // 设置单元格值
            cell.setCellValue(cellData.getValue());
        }
    }

    /**
     * 设置区域样式
     *
     * @param sheet  工作表
     * @param region 区域
     * @param style  单元格样式
     */
    private void setRegionStyle(Sheet sheet, CellRangeAddress region, CellStyle style) {
        for (int rowIndex = region.getFirstRow(); rowIndex <= region.getLastRow(); rowIndex++) {
            Row row = getRow(sheet, rowIndex);

            for (int colIndex = region.getFirstColumn(); colIndex <= region.getLastColumn(); colIndex++) {
                Cell cell = getCell(row, colIndex);
                cell.setCellStyle(style);
            }
        }
    }

    private Row getRow(Sheet sheet, Integer index) {
        int rowIndex = index == null ? sheet.getLastRowNum() : index;
        Row row = sheet.getRow(rowIndex);

        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        return row;
    }

    private Cell getCell(Row row, Integer index) {
        int colIndex = index == null ? row.getLastCellNum() : index;
        colIndex = Math.max(colIndex, 0);
        Cell cell = row.getCell(colIndex);

        if (cell == null) {
            cell = row.createCell(colIndex);
            cell.setCellValue(" ");
        }

        return cell;
    }

    /**
     * 获取单元格样式
     *
     * @param style 样式枚举
     */
    private CellStyle getCellStyle(Integer style) {
        if (RED_FONT_STYLE.equals(style)) {
            return getRedFontStyle();
        } else if (CENTER_FONT_STYLE.equals(style)) {
            return getCenterStyle();
        } else {
            return getDefaultStyle();
        }
    }

    /**
     * 创建新样式
     */
    private CellStyle createDefaultCellStyle() {
        return workbook.createCellStyle();
    }

    /**
     * 获取默认样式
     */
    private CellStyle getDefaultStyle() {
        if (!cellStyleMap.containsKey(DEFAULT_FONT_STYLE)) {
            cellStyleMap.put(DEFAULT_FONT_STYLE, createDefaultCellStyle());
        }
        return cellStyleMap.get(DEFAULT_FONT_STYLE);
    }

    /**
     * 获取红色字体样式
     */
    private CellStyle getRedFontStyle() {
        if (!cellStyleMap.containsKey(RED_FONT_STYLE)) {
            // 创建字体
            Font font = workbook.createFont();
            font.setColor(IndexedColors.RED.getIndex());

            // 创建样式对象
            CellStyle style = createDefaultCellStyle();
            style.setFont(font);
            cellStyleMap.put(RED_FONT_STYLE, style);
        }
        return cellStyleMap.get(RED_FONT_STYLE);
    }

    /**
     * 获取居中样式
     */
    private CellStyle getCenterStyle() {
        if (!cellStyleMap.containsKey(CENTER_FONT_STYLE)) {
            CellStyle style = createDefaultCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            cellStyleMap.put(CENTER_FONT_STYLE, style);
        }
        return cellStyleMap.get(CENTER_FONT_STYLE);
    }
}
