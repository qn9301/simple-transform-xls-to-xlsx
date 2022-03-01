package xxx.common.utils;
import xxx.common.entity.ExcelCell;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.*;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelUtil {

    /**
     * 获取合并单元格的值
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static ExcelCell getMergedRegionEntity(Sheet sheet , int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    ExcelCell cell = new ExcelCell();
                    cell.setMergedRegion(true);
                    cell.setFirstColumn(firstColumn);
                    cell.setLastColumn(lastColumn);
                    cell.setFirstRow(firstRow);
                    cell.setLastRow(lastRow);
                    cell.setValue(getCellValue(fCell));
                    return cell;
                }
            }
        }
        return null ;
    }

    /**
     * 判断指定的单元格是否是合并单元格
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet , int row , int column){
        int sheetMergeCount = sheet.getNumMergedRegions();

        for(int i = 0 ; i < sheetMergeCount ; i++ ){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){

                    return true ;
                }
            }
        }

        return false ;
    }

    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell){

        if(cell == null) return "";

        if(cell.getCellType() == STRING){

            return cell.getStringCellValue();

        }else if(cell.getCellType() == BOOLEAN){

            return String.valueOf(cell.getBooleanCellValue());

        }else if(cell.getCellType() == FORMULA){

            return cell.getCellFormula() ;

        }else if(cell.getCellType() == NUMERIC){

            return String.valueOf(cell.getNumericCellValue());

        }

        return "";
    }

    /**
     * 转换 AA => 27
     * @param rowName
     * @return
     * @throws Exception
     */
    public static Integer getCellRowIndex(String rowName) throws Exception {
        Integer A_charAt = Integer.valueOf('A');
        Integer Z_charAt = Integer.valueOf('Z');
        Integer full = Z_charAt - A_charAt;
        char[] codeList = rowName.toUpperCase().toCharArray();
        Integer offset = A_charAt - 1;
        int num = 0;
        for (int i=0; i<codeList.length; i++) {
            char chr = codeList[i];
            Integer chr_charAt = Integer.valueOf(chr);
            if (chr_charAt < A_charAt
                || chr_charAt > Z_charAt) {
                throw new Exception("非法的单元格位置！");
            }
            num += (chr_charAt - offset) + full * i;
        }
        return num;
    }

    /**
     * 分割cell地址
     * AAA111 => { x: AAA, y:111 }
     * @param cellAddr
     * @return
     */
    public static Map<String, String> splitSingleCellAddr(String cellAddr) {
        String x = new String(cellAddr);
        String y = new String(cellAddr);
        x = x.replaceAll("\\d+", "");
        y = y.replaceAll("[a-zA-Z]+", "");
        Map<String, String> res = new HashMap<>();
        res.put("x", x);
        res.put("y", y);
        return res;
    }

    /**
     * 将xls转换为xlsx
     * @param xls
     * @return
     */
    public static XSSFWorkbook xls2xlsx(HSSFWorkbook xls) {
        int sheetNum = xls.getNumberOfSheets();
        XSSFWorkbook newXlsx = new XSSFWorkbook();
        for (int i=0; i<sheetNum; i++) {
            HSSFSheet oldSheet = xls.getSheetAt(i);
            XSSFSheet newSheet = newXlsx.createSheet(oldSheet.getSheetName());
            oldSheet.getActiveCell();
            // sheet页隐藏
            if (xls.isSheetHidden(i)) {
                newXlsx.setSheetHidden(i, true);
            }

            // 冻结的单元格 ok
            PaneInformation paneInformation = oldSheet.getPaneInformation();
            if (!Objects.isNull(paneInformation)) {
                newSheet.createFreezePane(paneInformation.getVerticalSplitPosition(), paneInformation.getHorizontalSplitPosition(),
                    paneInformation.getVerticalSplitLeftColumn(), paneInformation.getHorizontalSplitTopRow());
            }
            // 这个不知道是设置什么 ？
            int[] oldColumnBreaks = oldSheet.getColumnBreaks();
            for (int j : oldColumnBreaks) {
                newSheet.setColumnBreak(j);
            }

            // 设置默认行高
            newSheet.setDefaultRowHeightInPoints(oldSheet.getDefaultRowHeightInPoints());
            newSheet.setDefaultRowHeight(oldSheet.getDefaultRowHeight());
            newSheet.setDefaultColumnWidth(oldSheet.getDefaultColumnWidth());
            // 合并单元格 OK
            List<CellRangeAddress> mergeRegins = oldSheet.getMergedRegions();
            for (int j = 0; j < mergeRegins.size(); j++) {
                newSheet.addMergedRegion(mergeRegins.get(j));
            }
            // 保护单元格 @TODO
            // 下拉列表（单元格校验） OK
            XSSFDataValidationHelper helper = new XSSFDataValidationHelper(newSheet);
            List<HSSFDataValidation> validationList = oldSheet.getDataValidations();
            for (HSSFDataValidation item : validationList) {
                CellRangeAddressList cellRangeAddressList = item.getRegions();
                DataValidationConstraint dvConstraint =  item.getValidationConstraint();
                XSSFDataValidationConstraint dv = copyDataValidationConstraint(dvConstraint);
                DataValidation dataValidation = helper.createValidation(dv, cellRangeAddressList);
                newSheet.addValidationData(dataValidation);
            }
            // 单元格的值
            for (int j = 0; j <= oldSheet.getLastRowNum(); j++) {
                HSSFRow row = oldSheet.getRow(j);
                if(row != null){
                    XSSFRow newRow = newSheet.getRow(j);
                    if (Objects.isNull(newRow)) newRow = newSheet.createRow(j);
                    // 设置行高 OK
                    newRow.setHeight(row.getHeight());
                    newRow.setHeightInPoints(row.getHeightInPoints());
                    // 隐藏列 —— 无效
                    newRow.setZeroHeight(row.getZeroHeight());
                    // 设置列样式 OK
                    XSSFCellStyle newStyle = newXlsx.createCellStyle();
                    if (!Objects.isNull(row.getRowStyle())) {
                        HSSFFont oldFont = row.getRowStyle().getFont(xls);
                        if (!Objects.isNull(oldFont)) {
                            XSSFFont newFont = newXlsx.createFont();
                            transXSSFFont(oldFont, newFont);
                            newStyle.setFont(newFont);
                        }
                    }
                    transXSSFStyle(row.getRowStyle(), newStyle);
                    newRow.setRowStyle(newStyle);
                    for (int k = oldSheet.getRow(j).getFirstCellNum(); k < oldSheet.getRow(j).getLastCellNum(); k++) {
                        // 设置列宽 OK
                        newSheet.setColumnWidth(k, oldSheet.getColumnWidth(k));
                        // 获取每个单元格
                        HSSFCell cell = row.getCell(k);
                        if (cell == null) {
                            continue;
                        }
                        XSSFCell newCell = newRow.getCell(k);
                        // 隐藏行 —— 无效
                        if (oldSheet.isColumnHidden(k)) {
                            newSheet.setColumnHidden(k, true);
                        }
                        if (Objects.isNull(newCell)) newCell = newRow.createCell(k);
                        // 设置单元格样式 OK
                        XSSFCellStyle newCellStyle = newXlsx.createCellStyle();
                        transXSSFStyle(cell.getCellStyle(), newCellStyle);
                        // 设置单元格字体
                        if (!Objects.isNull(cell.getCellStyle())) {
                            HSSFFont oldFont = cell.getCellStyle().getFont(xls);
                            if (!Objects.isNull(oldFont)) {
                                XSSFFont newFont = newXlsx.createFont();
                                transXSSFFont(oldFont, newFont);
                                newCellStyle.setFont(newFont);
                            }
                        }
                        newCell.setCellStyle(newCellStyle);
                        // 设置单元格批注 OK
                        // 批注值 yes
                        XSSFComment newComment = copyComment(cell, newSheet.createDrawingPatriarch());
                        // 批注样式 no
                        // 将批注添加到单元格对象中
                        newCell.setCellComment(newComment);
                        // 设置单元格公式 OK
                        newCell.setHyperlink(cell.getHyperlink());
                        // 设置单元格值 OK
                        newCell.setCellType(cell.getCellType());
                        switch (cell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(cell.getRichStringCellValue().getString());
                                break;
                            case NUMERIC:
                                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                                    newCell.setCellValue(cell.getDateCellValue());
                                } else {
                                    newCell.setCellValue(cell.getNumericCellValue());
                                }
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(cell.getCellFormula());
                                break;
                            case ERROR:
                                newCell.setCellErrorValue(cell.getErrorCellValue());
                            default:
                                newCell.setCellValue("");
                                break;
                        }
                    }
                }
            }
        }
        return newXlsx;
    }

    /**
     * 迁移字体
     * @param oldFont
     * @param newFont
     */
    private static void transXSSFFont(HSSFFont oldFont, XSSFFont newFont) {
        newFont.setBold(oldFont.getBold());
        newFont.setFontName(oldFont.getFontName());
        newFont.setColor(oldFont.getColor());
        newFont.setFontHeightInPoints(oldFont.getFontHeightInPoints());
        newFont.setItalic(oldFont.getItalic());
        newFont.setCharSet(oldFont.getCharSet());
        newFont.setFontHeight(oldFont.getFontHeight());
        newFont.setStrikeout(oldFont.getStrikeout());
        newFont.setTypeOffset(oldFont.getTypeOffset());
        newFont.setUnderline(oldFont.getUnderline());
    }

    /**
     * 复制注解
     * @param cell
     * @param newDrawing
     * @return
     */
    private static XSSFComment copyComment(HSSFCell cell, XSSFDrawing newDrawing) {
        HSSFComment oldComment = cell.getCellComment();
        if (Objects.isNull(oldComment)) return null;
        ClientAnchor anchor = new XSSFClientAnchor();
        // 关键修改
        anchor.setDx1(0);
        anchor.setDx2(0);
        anchor.setDy1(0);
        anchor.setDy2(0);
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        anchor.setCol2(cell.getColumnIndex() + 5);
        anchor.setRow2(cell.getRowIndex() + 6);
        // 结束
        XSSFComment comment = newDrawing.createCellComment(anchor);
        // 输入批注信息
        HSSFRichTextString str = oldComment.getString();
        XSSFRichTextString newStr = copyRichTextString(str);
        comment.setString(newStr);
        return comment;
    }

    /**
     * 复制富文本
     * @param str
     * @return
     */
    private static XSSFRichTextString copyRichTextString(HSSFRichTextString str) {
        XSSFRichTextString newStr = new XSSFRichTextString();
        newStr.setString(str.getString());
        return newStr;
    }

    /**
     * 复制下拉框
     * @param dvConstraint
     * @return
     */
    private static XSSFDataValidationConstraint copyDataValidationConstraint(DataValidationConstraint dvConstraint) {
        System.out.println(dvConstraint.getExplicitListValues());
        XSSFDataValidationConstraint dv = new XSSFDataValidationConstraint(dvConstraint.getExplicitListValues());
        dv.setOperator(dvConstraint.getOperator());
        if (!Objects.isNull(dvConstraint.getFormula1())) {
            dv.setFormula1(dvConstraint.getFormula1());
        }
        if (!Objects.isNull(dvConstraint.getFormula2())) {
            dv.setFormula2(dvConstraint.getFormula2());
        }
        return dv;
    }

    /**
     * 转换单元格属性
     * @param oldStyle
     * @param newStyle
     */
    public static void transXSSFStyle(HSSFCellStyle oldStyle, XSSFCellStyle newStyle) {
        if (Objects.isNull(oldStyle)) return ;
        newStyle.setAlignment(oldStyle.getAlignment());
        newStyle.setVerticalAlignment(oldStyle.getVerticalAlignment());
        newStyle.setBorderBottom(oldStyle.getBorderBottom());
        newStyle.setBorderLeft(oldStyle.getBorderLeft());
        newStyle.setBorderRight(oldStyle.getBorderRight());
        newStyle.setBorderTop(oldStyle.getBorderTop());
        newStyle.setBottomBorderColor(oldStyle.getBottomBorderColor());
        newStyle.setTopBorderColor(oldStyle.getTopBorderColor());
        newStyle.setLeftBorderColor(oldStyle.getLeftBorderColor());
        newStyle.setRightBorderColor(oldStyle.getRightBorderColor());
        newStyle.setFillForegroundColor(oldStyle.getFillForegroundColor());
        newStyle.setFillPattern(oldStyle.getFillPattern());
//        newStyle.setFont(oldStyle.getFont());
        newStyle.setWrapText(oldStyle.getWrapText());
        newStyle.setDataFormat(oldStyle.getDataFormat());
        newStyle.setFillBackgroundColor(oldStyle.getFillForegroundColor());
        newStyle.setHidden(oldStyle.getHidden());
        newStyle.setIndention(oldStyle.getIndention());
        newStyle.setLocked(oldStyle.getLocked());
        newStyle.setQuotePrefixed(oldStyle.getQuotePrefixed());
        newStyle.setRotation(oldStyle.getRotation());
        newStyle.setShrinkToFit(oldStyle.getShrinkToFit());
    }

}
