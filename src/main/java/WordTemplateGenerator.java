import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

public class WordTemplateGenerator {
    private static final int TITLE_FONT_SIZE = 14;
    private static final int INFO_FONT_SIZE = 11;
    private static final int Forget_WORD_FONT_SIZE = 12;
    private static final int WORD_FONT_SIZE = 14;  // X单元格字体大小
    private static final int MEANING_FONT_SIZE = 11;  // Y单元格字体大小
    private static final int CELL_WIDTH = 1700;
    private static final int CELL_HEIGHT = 650;

    public static void generateDocument(String outputPath, String name, String date1,
                                        String date2, String wordCount,
                                        List<String> reviewDates,
                                        List<Pair<String, Pair<String, String>>> wordPairs) throws Exception {

        XWPFDocument doc = new XWPFDocument();

        // 设置页面边距为1.27厘米
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(720L);   // 1.27厘米 = 720 twips
        pageMar.setRight(720L);
        pageMar.setTop(720L);
        pageMar.setBottom(720L);

        // 添加标题
        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("家庭CEO 伴学课后单词打印");
        titleRun.setBold(true);
        titleRun.setFontSize(TITLE_FONT_SIZE);
        titleRun.setFontFamily("宋体");

        // 在一行中添加所有信息
        XWPFParagraph info = doc.createParagraph();
        info.setAlignment(ParagraphAlignment.CENTER);

        // 添加姓名
        XWPFRun infoRun = info.createRun();
        infoRun.setText("姓名：");
        infoRun.setFontSize(INFO_FONT_SIZE);
        infoRun.setFontFamily("宋体");

        XWPFRun nameValueRun = info.createRun();
        nameValueRun.setText(name);
        nameValueRun.setUnderline(UnderlinePatterns.SINGLE);
        nameValueRun.setFontSize(INFO_FONT_SIZE);
        nameValueRun.setFontFamily("宋体");

        // 添加日期
        XWPFRun dateRun = info.createRun();
        dateRun.setText(" 日期：");
        dateRun.setFontSize(INFO_FONT_SIZE);
        dateRun.setFontFamily("宋体");

        XWPFRun dateValueRun = info.createRun();
        dateValueRun.setText(date1);
        dateValueRun.setUnderline(UnderlinePatterns.SINGLE);
        dateValueRun.setFontSize(INFO_FONT_SIZE);
        dateValueRun.setFontFamily("宋体");

        // 添加课程开始日期
        XWPFRun startDateRun = info.createRun();
        startDateRun.setText(" 课程开始日期：");
        startDateRun.setFontSize(INFO_FONT_SIZE);
        startDateRun.setFontFamily("宋体");

        XWPFRun startDateValueRun = info.createRun();
        startDateValueRun.setText(date2);
        startDateValueRun.setUnderline(UnderlinePatterns.SINGLE);
        startDateValueRun.setFontSize(INFO_FONT_SIZE);
        startDateValueRun.setFontFamily("宋体");

        // 添加词数
        XWPFRun wordCountRun = info.createRun();
        wordCountRun.setText(" 词数：");
        wordCountRun.setFontSize(INFO_FONT_SIZE);
        wordCountRun.setFontFamily("宋体");

        XWPFRun wordCountValueRun = info.createRun();
        wordCountValueRun.setText(wordCount);
        wordCountValueRun.setUnderline(UnderlinePatterns.SINGLE);
        wordCountValueRun.setFontSize(INFO_FONT_SIZE);
        wordCountValueRun.setFontFamily("宋体");

        XWPFRun wordRun = info.createRun();
        wordRun.setText("词");
        wordRun.setFontSize(INFO_FONT_SIZE);
        wordRun.setFontFamily("宋体");

        // 添加空行
        doc.createParagraph();

        // 创建复习日期表格
        XWPFTable reviewTable = doc.createTable(3, 11);
        setTableWidth(reviewTable, "100%");
        setReviewTableBorders(reviewTable);

        // 设置第一个表格的列宽
        CTTblGrid grid = reviewTable.getCTTbl().addNewTblGrid();
        for (int i = 0; i < 11; i++) {
            CTTblGridCol gridCol = grid.addNewGridCol();
            gridCol.setW(BigInteger.valueOf(i == 0 ? 1000 : 900));
        }

        // 填充表头
        String[] headers = {"", "第1天", "第2天", "第3天", "第5天", "第7天",
                "第9天", "第12天", "第14天", "第17天", "第21天"};
        XWPFTableRow headerRow = reviewTable.getRow(0);
        for (int i = 0; i < headers.length; i++) {
            setCellText(headerRow.getCell(i), headers[i], INFO_FONT_SIZE);
            headerRow.getCell(i).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        }

        // 填充复习日期行
        XWPFTableRow dateRow = reviewTable.getRow(1);
        setCellText(dateRow.getCell(0), "复习日期", Forget_WORD_FONT_SIZE);
        for (int i = 0; i < reviewDates.size(); i++) {
            setCellText(dateRow.getCell(i + 1), reviewDates.get(i), INFO_FONT_SIZE);
        }

        // 遗忘词数行
        XWPFTableRow forgetRow = reviewTable.getRow(2);
        setCellText(forgetRow.getCell(0), "遗忘词数", Forget_WORD_FONT_SIZE);

        // 添加空行
        doc.createParagraph();

        // 创建单词对照表格
        int numRows = (wordPairs.size() + 2) / 3;
        XWPFTable wordTable = doc.createTable(numRows, 6);
        setTableWidth(wordTable, "100%");
        setWordTableBorders(wordTable);

        // 设置单词表格的固定列宽
        CTTblGrid wordGrid = wordTable.getCTTbl().addNewTblGrid();
        for (int i = 0; i < 6; i++) {
            CTTblGridCol gridCol = wordGrid.addNewGridCol();
            gridCol.setW(BigInteger.valueOf(CELL_WIDTH));
        }

        // 设置表格属性
        CTTbl ttbl = wordTable.getCTTbl();
        CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
        CTTblLayoutType tblLayout = tblPr.getTblLayout() == null ? tblPr.addNewTblLayout() : tblPr.getTblLayout();
        tblLayout.setType(STTblLayoutType.FIXED);

        // 填充单词表格
        int rowIndex = 0;
        int pairIndex = 0;
        while (pairIndex < wordPairs.size()) {
            XWPFTableRow row = wordTable.getRow(rowIndex);
            setRowHeight(row, CELL_HEIGHT);

            for (int i = 0; i < 3 && pairIndex < wordPairs.size(); i++) {
                Pair<String, Pair<String, String>> pair = wordPairs.get(pairIndex++);
                XWPFTableCell cellX = row.getCell(i * 2);
                XWPFTableCell cellY = row.getCell(i * 2 + 1);

                setCellProperties(cellX, false);  // X列右边是单线
                // 判断是否是最右侧的Y列 (i == 2 表示最后一组)
                boolean isLastColumn = (i == 2);
                setCellProperties(cellY, !isLastColumn);  // 如果不是最后一列，则使用双线；否则使用单线

                // 设置X单元格（单词）
                XWPFParagraph xPara = cellX.getParagraphs().get(0);
                xPara.setAlignment(ParagraphAlignment.CENTER);  // 水平居中
                xPara.setVerticalAlignment(TextAlignment.CENTER);  // 垂直居中
                xPara.setSpacingBefore(0);
                xPara.setSpacingAfter(0);
                cellX.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);  // 单元格级别的垂直居中

                XWPFRun xRun = xPara.createRun();
                xRun.setFontSize(WORD_FONT_SIZE);  // 字体大小
                xRun.setFontFamily("Arial");    // 字体
                xRun.setBold(true);                // 加粗
                xRun.setText(pair.getFirst());

                // 设置Y单元格（音标和释义）
                XWPFParagraph yPara = cellY.getParagraphs().get(0);
                yPara.setAlignment(ParagraphAlignment.CENTER);  // 水平居中
                yPara.setVerticalAlignment(TextAlignment.CENTER);  // 垂直居中
                yPara.setSpacingBefore(0);
                yPara.setSpacingAfter(0);
                cellY.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);  // 单元格级别的垂直居中

                // 添加音标
                XWPFRun phoneticRun = yPara.createRun();
                phoneticRun.setFontSize(MEANING_FONT_SIZE);  // 字体大小
                phoneticRun.setFontFamily("Arial");       // 字体
//                phoneticRun.setBold(true);                   // 加粗
                phoneticRun.setText(pair.getSecond().getFirst());

                // 添加释义
                XWPFRun meaningRun = yPara.createRun();
                meaningRun.setFontSize(MEANING_FONT_SIZE);  // 11号字体
                meaningRun.setFontFamily("宋体");       // 字体
//                meaningRun.setBold(true);                   // 加粗
                meaningRun.setText(" " + pair.getSecond().getSecond());
            }
            rowIndex++;
        }

        // 保存文档
        try (FileOutputStream out = new FileOutputStream(outputPath)) {
            doc.write(out);
        }
    }

    private static void setCellProperties(XWPFTableCell cell, boolean isRightBorder) {
        cell.setWidth(String.valueOf(CELL_WIDTH));
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.TOP);

        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();

        // 设置单元格边距
        CTTcMar tcMar = tcPr.isSetTcMar() ? tcPr.getTcMar() : tcPr.addNewTcMar();

        CTTblWidth top = CTTblWidth.Factory.newInstance();
        top.setType(STTblWidth.DXA);
        top.setW(BigInteger.valueOf(100));
        tcMar.setTop(top);

        CTTblWidth bottom = CTTblWidth.Factory.newInstance();
        bottom.setType(STTblWidth.DXA);
        bottom.setW(BigInteger.valueOf(100));
        tcMar.setBottom(bottom);

        CTTblWidth left = CTTblWidth.Factory.newInstance();
        left.setType(STTblWidth.DXA);
        left.setW(BigInteger.valueOf(100));
        tcMar.setLeft(left);

        CTTblWidth right = CTTblWidth.Factory.newInstance();
        right.setType(STTblWidth.DXA);
        right.setW(BigInteger.valueOf(100));
        tcMar.setRight(right);

        // 设置单元格边框
        CTTcBorders borders = tcPr.isSetTcBorders() ? tcPr.getTcBorders() : tcPr.addNewTcBorders();

        // 设置右边框样式
        CTBorder rightBorder = borders.isSetRight() ? borders.getRight() : borders.addNewRight();
        if (isRightBorder) {
            rightBorder.setVal(STBorder.DOUBLE);
            rightBorder.setSz(BigInteger.valueOf(8));  // 增加双线宽度,原来是4
        } else {
            rightBorder.setVal(STBorder.SINGLE);
            rightBorder.setSz(BigInteger.valueOf(4));  // 保持单线宽度不变
        }
        rightBorder.setSpace(BigInteger.valueOf(0));
        rightBorder.setColor("000000");
    }

    private static void setRowHeight(XWPFTableRow row, int height) {
        CTRow ctRow = row.getCtRow();
        CTTrPr trPr = ctRow.isSetTrPr() ? ctRow.getTrPr() : ctRow.addNewTrPr();
        CTHeight ctHeight = trPr.sizeOfTrHeightArray() == 0 ? trPr.addNewTrHeight() : trPr.getTrHeightArray(0);
        ctHeight.setVal(BigInteger.valueOf(height));
        ctHeight.setHRule(STHeightRule.AT_LEAST);
    }

    private static void setCellText(XWPFTableCell cell, String text, int fontSize) {
        XWPFParagraph paragraph = cell.getParagraphs().get(0);
        XWPFRun run = paragraph.createRun();
        run.setFontSize(fontSize);
        run.setFontFamily("宋体");
        run.setText(text);
    }

    private static void setTableWidth(XWPFTable table, String width) {
        table.setWidthType(TableWidthType.PCT);
        table.setWidth(width);
    }

    private static void setReviewTableBorders(XWPFTable table) {
        // 设置所有边框为单线
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setInsideHBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
    }

    private static void setWordTableBorders(XWPFTable table) {
        // 设置表格外边框为单线
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 4, 0, "000000");
        // 设置水平内边框为双线
        table.setInsideHBorder(XWPFTable.XWPFBorderType.DOUBLE, 8, 0, "000000");
    }

    public static class Pair<T, U> {
        private final T first;
        private final U second;

        public Pair(T first, U second) {
            this.first = first;
            this.second = second;
        }

        public T getFirst() {
            return first;
        }

        public U getSecond() {
            return second;
        }
    }
}