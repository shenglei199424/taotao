import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

public class ExcelUtil {

public static SXSSFWorkbook  createWorkbook(){
    return  new SXSSFWorkbook();
}
    /**
     * 设置下拉框
     *
     * @param sheet sheet 对象
     * @param list     下拉枚举值string[]
     * @param firstRow 开始行数
     * @param lastRow  结束行数
     * @param firstCol 第一列
     * @param lastCol  最后一列
     */
    public static void setDataValidation(SXSSFSheet sheet, String[] list, Integer firstRow, Integer lastRow, Integer firstCol, Integer lastCol) {
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = dataValidationHelper.createExplicitListConstraint(list);
        DataValidation dataValidation = dataValidationHelper.createValidation(constraint, cellRangeAddressList);
        dataValidation.createErrorBox("本系统提醒您", "数据不规范，请选择表格列中的数据");
        //处理excel兼容性问题
        if (dataValidation instanceof XSSFDataValidation) {
            dataValidation.setSuppressDropDownArrow(true);
            dataValidation.setShowErrorBox(true);
        } else {
            dataValidation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(dataValidation);
    }

    /**
     * @param sheet sheet对象
     * @param columnNum   第几列
     * @param columnWidth 宽度
     */
    public static void setColumnWidth(Sheet sheet, int columnNum, int columnWidth) {
        sheet.setColumnWidth(columnNum, columnWidth * 256);
    }

    /**
     * 该方法用于返回带边框显示并且左对齐的cellstyle
     *
     * @param workbook 工作簿
     * @return cellStyle 表格样式
     */
    public static CellStyle createAlignLeftCellStyle(SXSSFWorkbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 创建没有边框左对齐的cellstyle
     *
     * @param workbook 工作簿
     * @return cellstyle 表格样式
     */
    public static CellStyle createAlignLeftCellStyleWithoutBorder(SXSSFWorkbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        return cellStyle;
    }

    /**
     * 返回指定字体，字高，是否加粗的字体
     * @param workbook 工作表
     * @param fontName 字体名称
     * @param fontHeight 字高
     * @param bold 字体是否加粗
     * @return Font
     */
    public static Font createFont(SXSSFWorkbook workbook, String fontName, int fontHeight, Boolean bold) {
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setFontHeight((short) fontHeight);
        font.setBold(bold);
        return  font;
    }

    /**
     * 返回文本格式的样式
     * @param workbook 工作本
     * @return CellStyle 表格样式
     */
    public CellStyle getDataTextFormatCellSyle(SXSSFWorkbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));
        return cellStyle;
    }

    /**
     * 返回保留小数位的cellstyle
     * @param workbook  工作表
     * @param decimalPlace 保留小数位数
     * @return cellStyle 表格样式
     */
    public CellStyle getDecimalFormatCellSyle(SXSSFWorkbook workbook,int decimalPlace){
        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        if (2==decimalPlace){
            cellStyle.setDataFormat(dataFormat.getFormat("0.00"));
        }else if (3==decimalPlace){
            cellStyle.setDataFormat(dataFormat.getFormat("0.000"));
        }else if (4==decimalPlace){
            cellStyle.setDataFormat(dataFormat.getFormat("0.0000"));
        }
        return cellStyle;
}

    /**
     * 根据RGB来设置背景颜色并返回cellstyle
     * @param workbook 工作本
     * @param r 颜色Red
     * @param g 颜色Green
     * @param b 颜色blue
     * @return cellStyle
     */
    public static XSSFCellStyle getBackgroundColorCellStyle(SXSSFWorkbook workbook,int r,int g,int b) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(r,g,b)));
        return cellStyle;
    }
}