package org.cosine.word;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.cosine.cache.FileLoader;
import org.cosine.cache.TemplateManager;
import org.cosine.model.ElLabel;
import org.cosine.model.WordImage;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * word.docx 2007版本word文档填充
 * @author wenbing.li
 * @date 2020/6/25 13:23
 */
public class FillWord07 {
    /**
     * 填充word
     * @param inputStream 模板流
     * @param map 填充参数
     * @return 填充后的文档
     * @exception  IOException  if an I/O error occurs. In particular,
     *             an <code>IOException</code> may be thrown if the
     *             output stream has been closed.
     */
    public XWPFDocument fillWord(InputStream inputStream, Map<String, Object> map) throws IOException {
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        this.fillWord(xwpfDocument, map);
        return xwpfDocument;
    }

    /**
     * 填充word
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param map 填充参数
     * @return 填充后的文档
     * @exception  IOException  if an I/O error occurs. In particular,
     *             an <code>IOException</code> may be thrown if the
     *             output stream has been closed.
     */
    public XWPFDocument fillWord(String template, Map<String, Object> map) throws IOException {
        XWPFDocument xwpfDocument = TemplateManager.getXWPFDocument(template);
        this.fillWord(xwpfDocument, map);
        return xwpfDocument;
    }

    /**
     * 同一个模板，填充多组数据
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @param list 模板参数
     * @param ifPagination 多页数据之间是否插入分页
     * @return 填充后的文档
     */
    public XWPFDocument fillWord(String template, List<Map<String, Object>> list,boolean ifPagination) throws IOException {
        if (list == null || list.size() == 0) {
            return null;
        } else if (list.size() == 1) {
            XWPFDocument xwpfDocument = TemplateManager.getXWPFDocument(template);
            this.fillWord(xwpfDocument, list.get(0));
            return xwpfDocument;
        } else {
            XWPFDocument xwpfDocument = TemplateManager.getXWPFDocument(template);
            this.fillWord(xwpfDocument, list.get(0));
            //插入分页
            if(ifPagination) {
                insertPage(xwpfDocument,xwpfDocument);
            }
            for (int i = 1; i < list.size(); i++) {
                XWPFDocument tempDoc = TemplateManager.getXWPFDocument(template);
                this.fillWord(tempDoc, list.get(i));
                if(ifPagination) {
                    if ((i + 1) < list.size()) {
                        insertPage(tempDoc, xwpfDocument);
                    }
                }
                xwpfDocument.getDocument().addNewBody().set(tempDoc.getDocument().getBody());
            }
            return xwpfDocument;
        }
    }

    /**
     * 不同模板，填充不同也数据
     * @param mapList 模板参数 Map<模板地址, Map<String, Object>>
     * @param ifPagination 多页数据之间是否插入分页
     * @return 填充后的文档
     */
    public XWPFDocument fillWord(Map<String, Map<String, Object>> mapList,boolean ifPagination) throws IOException {
        if (mapList == null || mapList.isEmpty()) {
            return null;
        }
        List<String> templates = new ArrayList<>(mapList.keySet());
        if (templates.size() == 1) {
            XWPFDocument xwpfDocument = TemplateManager.getXWPFDocument(templates.get(0));
            this.fillWord(xwpfDocument, mapList.get(templates.get(0)));
            return xwpfDocument;
        } else {
            XWPFDocument xwpfDocument = TemplateManager.getXWPFDocument(templates.get(0));
            this.fillWord(xwpfDocument, mapList.get(templates.get(0)));
            //插入分页
            if(ifPagination) {
                insertPage(xwpfDocument,xwpfDocument);
            }
            for (int i = 1; i < templates.size(); i++) {
                XWPFDocument tempDoc = TemplateManager.getXWPFDocument(templates.get(i));
                this.fillWord(tempDoc, mapList.get(templates.get(i)));
                if(ifPagination) {
                    if ((i + 1) < templates.size()) {
                        insertPage(tempDoc, tempDoc);
                    }
                }
                xwpfDocument.getDocument().addNewBody().set(tempDoc.getDocument().getBody());
            }
            return xwpfDocument;
        }
    }

    private void fillWord(XWPFDocument xwpfDocument, Map<String, Object> map) {
        if(xwpfDocument == null){
            throw new NullPointerException("XWPFDocument is null");
        }
        //填充段落
        fillAllParagraph(xwpfDocument.getParagraphs(), map);
        //填充页眉
        fillHeaderAndFoot(xwpfDocument,map);
        //填充表格
        fillAllTable(xwpfDocument.getTablesIterator(),map);
    }


    /**
     * 插入分页
     * @param tempDoc 插入分页的上一个页面
     * @param xwpfDocument 插入分页的页面
     */
    private static void insertPage(XWPFDocument tempDoc, XWPFDocument xwpfDocument){
        //创建一个分页段落，添加分页符
        XWPFParagraph pageP = tempDoc.createParagraph();
        copyXWPFParagraphCTP(pageP,xwpfDocument.getParagraphs().get(0));
        pageP.setPageBreak(true);
    }

    /**
     * 填充所有段落
     * @param paragraphs 段落集合
     * @param map 填充参数
     */
    private void fillAllParagraph(List<XWPFParagraph> paragraphs, Map<String, Object> map) {
        for(int i = 0; i < paragraphs.size(); ++i) {
            XWPFParagraph paragraph = paragraphs.get(i);
            if (paragraph.getText().contains(ElLabel.START_LABEL)) {
                this.fillThisParagraph(paragraph, map);
            }
        }
    }

    /**
     * 填充当前段落
     * @param paragraph 段落
     * @param map 填充参数
     */
    private void fillThisParagraph(XWPFParagraph paragraph, Map<String, Object> map) {
        XWPFRun currentRun = null;
        String currentText = "";
        boolean ifFind = false;
        List<Integer> runIndex = new ArrayList<>();
        // 解析段落中，每个区域的文本
        for(int i = 0; i < paragraph.getRuns().size(); ++i) {
            XWPFRun run = paragraph.getRuns().get(i);
            String text = run.getText(0);
            if (text != null) {
                if (ifFind) {
                    currentText = currentText + text;
                    if (!currentText.contains(ElLabel.START_LABEL)) {
                        ifFind = false;
                        runIndex.clear();
                    } else {
                        runIndex.add(i);
                    }
                    if (currentText.contains(ElLabel.END_LABEL)) {
                        this.fillValue(paragraph, currentRun, currentText, runIndex, map);
                        currentText = "";
                        ifFind = false;
                    }
                } else if (text.contains(ElLabel.START_LABEL) || ElLabel.START_LABEL.contains(text)) {
                    currentText = text;
                    ifFind = true;
                    currentRun = run;
                } else {
                    currentText = "";
                }

                if (currentText.contains(ElLabel.END_LABEL)) {
                    this.fillValue(paragraph, currentRun, currentText, runIndex, map);
                    ifFind = false;
                }
            }
        }
    }

    /**
     * 填充页眉
     * @param xwpfDocument 文档
     * @param map 填充参数
     */
    private void fillHeaderAndFoot(XWPFDocument xwpfDocument, Map<String, Object> map) {
        List<XWPFHeader> headerList = xwpfDocument.getHeaderList();
        Iterator<XWPFHeader> xwpfHeaders = headerList.iterator();
        while(xwpfHeaders.hasNext()) {
            XWPFHeader xwpfHeader = xwpfHeaders.next();
            for(int i = 0; i < xwpfHeader.getListParagraph().size(); ++i) {
                this.fillThisParagraph(xwpfHeader.getListParagraph().get(i), map);
            }
        }
        List<XWPFFooter> footerList = xwpfDocument.getFooterList();
        Iterator<XWPFFooter> xwpfFooters = footerList.iterator();
        while(xwpfFooters.hasNext()) {
            XWPFFooter xwpfFooter = xwpfFooters.next();
            for(int i = 0; i < xwpfFooter.getListParagraph().size(); ++i) {
                this.fillThisParagraph(xwpfFooter.getListParagraph().get(i), map);
            }
        }
    }

    /**
     * 填充word中所有的表格
     * @param tableIterator 表格集合
     * @param map 填充参数
     */
    private void fillAllTable(Iterator<XWPFTable> tableIterator, Map<String, Object> map) {
        while(tableIterator.hasNext()) {
            XWPFTable table = tableIterator.next();
            if (table.getText().contains(ElLabel.START_LABEL)) {
                this.fillThisTable(table, map);
            }
        }
    }

    /**
     * 填充当前表格
     * @param table 表格
     * @param map 填充参数
     */
    private void fillThisTable(XWPFTable table,Map<String, Object> map) {
        for(int i = 0; i < table.getNumberOfRows(); ++i) {
            XWPFTableRow row = table.getRow(i);
            List<XWPFTableCell> cells = row.getTableCells();
            // 检查是否为表格填充，并返回填充列表参数
            Object list = checkThisTableIsNeedIterator(cells.get(0), map);
            if (list == null) {
                this.fillThisRow(cells, map);
            } else if(list instanceof Collection){
                Collection<Object> coll = (Collection) list;
                this.fillNextRowAndAddRow(table, i, coll.iterator());
                i = i + coll.size() - 1;
            } else {
                this.fillThisRow(cells, map);
            }
        }
    }

    /**
     * 填充表格行和添加行
     * @param table 表格
     * @param index 索引
     * @param iterator 填充的List数据集
     */
    private void fillNextRowAndAddRow(XWPFTable table, int index, Iterator<Object> iterator) {
        XWPFTableRow uponRow = table.getRow(index);
        XWPFTableRow currentRow = table.getRow(index);
        String[] params = parseCurrentRowGetParams(currentRow);
        String itemName = params[0];
        boolean isCreate = itemName.contains(ElLabel.TRAVERSE_LABEL);
        String alias = itemName.substring(0,itemName.indexOf(ElLabel.START_LABEL) + ElLabel.START_LABEL.length());
        itemName = itemName.replace(alias+ElLabel.TRAVERSE_LABEL, "").replace(ElLabel.START_LABEL, "");
        String[] keys = itemName.replaceAll("\\s{1,}", " ").trim().split(" ");
        params[0] = keys[1];
        List<XWPFTableCell> tempCellList = table.getRow(index).getTableCells();

        Object fillValue;
        String field;
        while(iterator.hasNext()) {
            Object obj = iterator.next();
            currentRow = isCreate ? table.insertNewTableRow(index++) : table.getRow(index++);
            // 复制表格行属性
            currentRow.setHeight(uponRow.getHeight());
            currentRow.getCtRow().setTrPr(uponRow.getCtRow().getTrPr());

            int cellIndex;
            for(cellIndex = 0; cellIndex < currentRow.getTableCells().size(); ++cellIndex) {
                field = getCellMapField(params[cellIndex]);
                fillValue = getCellMapValue(field, obj);
                XWPFTableCell cell = tempCellList.get(cellIndex);
                XWPFTableCell nextCell = currentRow.getTableCells().get(cellIndex);
                nextCell.setText("");
                copyCellAndSetValue(cell, nextCell, fillValue);
            }

            while(cellIndex < params.length) {
                field = getCellMapField(params[cellIndex]);
                fillValue = getCellMapValue(field, obj);
                XWPFTableCell cell = tempCellList.get(cellIndex);
                copyCellAndSetValue(cell, currentRow.createCell(), fillValue);
                ++cellIndex;
            }
        }
        table.removeRow(index);
    }

    /**
     * 填充当前行
     * @param cells 当前行单元格集合
     * @param map 填充参数
     */
    private void fillThisRow(List<XWPFTableCell> cells, Map<String, Object> map) {
        Iterator<XWPFTableCell> iterator = cells.iterator();
        while(iterator.hasNext()) {
            XWPFTableCell cell = iterator.next();
            this.fillAllParagraph(cell.getParagraphs(), map);
        }
    }


    /**
     * 填充文本、图片
     * @param paragraph 段落
     * @param currentRun 文本区域
     * @param currentText 文本
     * @param runIndex 文本区域索引
     * @param map 填充值
     */
    private void fillValue(XWPFParagraph paragraph, XWPFRun currentRun, String currentText, List<Integer> runIndex, Map<String, Object> map) {
        Object obj = parseCurrentText(currentText, map);
        if (obj instanceof WordImage) {
            currentRun.setText("", 0);
            this.setWordImage(currentRun, (WordImage)obj);
        } else {
            currentText = obj.toString();
            this.setWordText(paragraph, currentRun, currentText);
        }
        for (Integer index : runIndex) {
            paragraph.getRuns().get(index).setText("", 0);
        }
        runIndex.clear();
    }

    /**
     * 填充文本，换行的位置使用 \r\n 标记
     * @param currentRun 文本区域
     * @param currentText 文本
     */
    private void setWordText(XWPFParagraph xwpfParagraph, XWPFRun currentRun, String currentText) {
        if (currentText != null && !"".equals(currentText)) {
            //硬回车
            String[] carriageReturnValues = currentText.split(ElLabel.CARRIAGE_RETURN_ESCAPE);
            IBody iBody = xwpfParagraph.getBody();
            int i = 0;
            if (carriageReturnValues.length > 0) {
                for(int le = carriageReturnValues.length - 1; i < le; ++i) {
                    //插入段落
                    XWPFParagraph newParagraph = iBody.insertNewParagraph(xwpfParagraph.getCTP().newCursor());
                    newParagraph.getCTP().setPPr(xwpfParagraph.getCTP().getPPr());
                    XWPFRun xwpfRun = newParagraph.createRun();
                    xwpfRun.getCTR().set(currentRun.getCTR());
                    setSoftCarriageReturnText(xwpfRun,carriageReturnValues[i]);
                }
                setSoftCarriageReturnText(currentRun,carriageReturnValues[carriageReturnValues.length - 1]);
            } else if(currentText.contains(ElLabel.CARRIAGE_RETURN_ESCAPE)){
                //插入段落
                XWPFParagraph newParagraph = iBody.insertNewParagraph(xwpfParagraph.getCTP().newCursor());
                newParagraph.getCTP().setPPr(xwpfParagraph.getCTP().getPPr());
                currentRun.setText("",0);
            }
        } else {
            currentRun.setText("", 0);
        }
    }

    /**
     * 填充word图片
     * @param currentRun 文本区域
     * @param image 图片对象
     */
    private void setWordImage(XWPFRun currentRun, WordImage image) {
        addWordImage(currentRun,image);
    }

    /**
     * 解析文本并填充
     * @param currentText 需要解析的文本
     * @param map 填充参数
     * @return 解析后的文本
     */
    private static Object parseCurrentText(String currentText, Map<String, Object> map){
        String params = "";
        while(currentText.contains(ElLabel.START_LABEL)) {
            params = currentText.substring(currentText.indexOf(ElLabel.START_LABEL) + 2, currentText.indexOf(ElLabel.END_LABEL));
            Object obj = map.get(params.trim());
            if (obj instanceof WordImage || obj instanceof List) {
                return obj;
            }
            if (obj != null) {
                currentText = currentText.replace(ElLabel.START_LABEL + params + ElLabel.END_LABEL, obj.toString());
            } else {
                currentText = currentText.replace(ElLabel.START_LABEL + params + ElLabel.END_LABEL, "");
            }
        }
        return currentText;
    }

    /**
     * 解析行参数
     * @param currentRow 需要解析的行
     * @return 行参数
     */
    private static String[] parseCurrentRowGetParams(XWPFTableRow currentRow) {
        List<XWPFTableCell> cells = currentRow.getTableCells();
        String[] params = new String[cells.size()];
        for(int i = 0; i < cells.size(); ++i) {
            String text = cells.get(i).getText();
            params[i] = text == null ? "" : text.trim().replace(ElLabel.START_LABEL, "").replace(ElLabel.END_LABEL, "");
        }
        return params;
    }

    /**
     * 获取表格单元格字段映射值
     * @param fieldName 字段参数
     * @param object 填充值
     * @return 映射值
     */
    private static Object getCellMapValue(String fieldName, Object object) {
        try {
            Class<?> clazz = object.getClass();
            Field field = clazz.getDeclaredField(fieldName);
            field.setAccessible(true);
            return field.get(object);
        }catch (Exception e){
            return "";
        }
    }

    /**
     * 获取表格单元格字段映射值
     * @param params 参数
     * @return 参数对应的值
     */
    private static String getCellMapField(String params) {
        String[] paramsArr = params.split("\\.");
        return paramsArr[1].trim();
    }

    /**
     * 检查单元格是否需要填充
     * @param cell 单元格
     * @param map 填充参数
     * @return 填充的List数据
     */
    private static Object checkThisTableIsNeedIterator(XWPFTableCell cell, Map<String, Object> map) {
        String text = cell.getText().trim();
        if (text.contains(ElLabel.TRAVERSE_LABEL) && text.startsWith(ElLabel.START_LABEL)) {
            String alias = text.substring(text.indexOf(ElLabel.START_LABEL) + ElLabel.START_LABEL.length(),text.indexOf(ElLabel.TRAVERSE_LABEL));
            text = text.replace(alias + ElLabel.TRAVERSE_LABEL, "").replace(ElLabel.START_LABEL, "");
            String[] keys = text.replaceAll("\\s{1,}", " ").trim().split(" ");
            Object result = map.get(keys[0]);
            return Objects.nonNull(result) ? result : new ArrayList<>(0);
        } else {
            return null;
        }
    }

    /**
     * 软回车
     * @param xwpfRun 软回车的文本区域
     * @param text 带回车符的文本
     */
    private static void setSoftCarriageReturnText(XWPFRun xwpfRun, String text){
        String[] softCarriageReturnValues = text.split(ElLabel.SOFT_CARRIAGE_RETURN_VALUES);
        if(softCarriageReturnValues.length > 0) {
            for (int j = 0, s = softCarriageReturnValues.length - 1; j < s; ++j) {
                xwpfRun.setText(softCarriageReturnValues[j],j);
                xwpfRun.addBreak();
            }
            xwpfRun.setText(softCarriageReturnValues[softCarriageReturnValues.length - 1], softCarriageReturnValues.length-1);
        } else if(text.contains(ElLabel.SOFT_CARRIAGE_RETURN_VALUES)){
            xwpfRun.setText("",0);
            xwpfRun.addBreak();
        }
    }

    /**
     * 填充word图片
     * @param currentRun 图片填充的区域
     * @param image 图片实体
     */
    public static void addWordImage(XWPFRun currentRun, WordImage image) {
        if(image == null){
            return;
        }
        try {
            if(image.getImageBytes() == null){
                image.setImageBytes(FileLoader.loaderFile(image.getImagePath()));
            }
            try (ByteArrayInputStream byteInputStream = new ByteArrayInputStream(image.getImageBytes())) {
                currentRun.addPicture(byteInputStream, Document.PICTURE_TYPE_JPEG, image.getImageName(), Units.toEMU(image.getWidth()), Units.toEMU(image.getHeight()));
            }
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    /**
     * 复制填充单元格样式和文本/图片
     * @param tmpCell 原单元格
     * @param cell 需要填充的单元格
     * @param fillValue 填充的值
     */
    public static void copyCellAndSetValue(XWPFTableCell tmpCell, XWPFTableCell cell, Object fillValue) {
        CTTc cttc2 = tmpCell.getCTTc();
        CTTcPr ctPr2 = cttc2.getTcPr();
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        if (tmpCell.getColor() != null) {
            cell.setColor(tmpCell.getColor());
        }

        if (tmpCell.getVerticalAlignment() != null) {
            cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
        }

        if (ctPr2.getTcW() != null) {
            ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
        }

        if (ctPr2.getVAlign() != null) {
            ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
        }

        if (cttc2.getPList().size() > 0) {
            CTP ctp = cttc2.getPList().get(0);
            if (ctp.getPPr() != null && ctp.getPPr().getJc() != null) {
                cttc.getPList().get(0).addNewPPr().addNewJc().setVal(ctp.getPPr().getJc().getVal());
            }
        }

        if (ctPr2.getTcBorders() != null) {
            ctPr.setTcBorders(ctPr2.getTcBorders());
        }

        XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
        XWPFParagraph cellP = cell.getParagraphs().get(0);
        XWPFRun tmpR = null;
        if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
            tmpR = tmpP.getRuns().get(0);
        }
        XWPFRun cellR = cellP.createRun();
        if(fillValue instanceof WordImage){
            addWordImage(cellR, (WordImage) fillValue);
        }else{
            cellR.setText(fillValue != null ? fillValue.toString() : "");
        }

        if (tmpR != null) {
            cellR.setBold(tmpR.isBold());
            cellR.setItalic(tmpR.isItalic());
            cellR.setStrikeThrough(tmpR.isStrikeThrough());
            cellR.setUnderline(tmpR.getUnderline());
            cellR.setColor(tmpR.getColor());
            cellR.setTextPosition(tmpR.getTextPosition());
            if (tmpR.getFontSize() != -1) {
                cellR.setFontSize(tmpR.getFontSize());
            }

            if (tmpR.getFontFamily() != null) {
                cellR.setFontFamily(tmpR.getFontFamily());
            }

            if (tmpR.getCTR() != null && tmpR.getCTR().isSetRPr()) {
                CTRPr tmpRPr = tmpR.getCTR().getRPr();
                if (tmpRPr.isSetRFonts()) {
                    CTFonts tmpFonts = tmpRPr.getRFonts();
                    CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR.getCTR().getRPr() : cellR.getCTR().addNewRPr();
                    CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr.getRFonts() : cellRPr.addNewRFonts();
                    cellFonts.setAscii(tmpFonts.getAscii());
                    cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
                    cellFonts.setCs(tmpFonts.getCs());
                    cellFonts.setCstheme(tmpFonts.getCstheme());
                    cellFonts.setEastAsia(tmpFonts.getEastAsia());
                    cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
                    cellFonts.setHAnsi(tmpFonts.getHAnsi());
                    cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
                }
            }
        }

        if (tmpP.getAlignment() != null) {
            cellP.setAlignment(tmpP.getAlignment());
        }

        if (tmpP.getVerticalAlignment() != null) {
            cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
        }

        if (tmpP.getBorderBetween() != null) {
            cellP.setBorderBetween(tmpP.getBorderBetween());
        }

        if (tmpP.getBorderBottom() != null) {
            cellP.setBorderBottom(tmpP.getBorderBottom());
        }

        if (tmpP.getBorderLeft() != null) {
            cellP.setBorderLeft(tmpP.getBorderLeft());
        }

        if (tmpP.getBorderRight() != null) {
            cellP.setBorderRight(tmpP.getBorderRight());
        }

        if (tmpP.getBorderTop() != null) {
            cellP.setBorderTop(tmpP.getBorderTop());
        }

        cellP.setPageBreak(tmpP.isPageBreak());
        if (tmpP.getCTP() != null && tmpP.getCTP().getPPr() != null) {
            CTPPr tmpPPr = tmpP.getCTP().getPPr();
            CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP.getCTP().getPPr() : cellP.getCTP().addNewPPr();
            CTSpacing tmpSpacing = tmpPPr.getSpacing();
            if (tmpSpacing != null) {
                CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr.getSpacing() : cellPPr.addNewSpacing();
                if (tmpSpacing.getAfter() != null) {
                    cellSpacing.setAfter(tmpSpacing.getAfter());
                }

                if (tmpSpacing.getAfterAutospacing() != null) {
                    cellSpacing.setAfterAutospacing(tmpSpacing.getAfterAutospacing());
                }

                if (tmpSpacing.getAfterLines() != null) {
                    cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
                }

                if (tmpSpacing.getBefore() != null) {
                    cellSpacing.setBefore(tmpSpacing.getBefore());
                }

                if (tmpSpacing.getBeforeAutospacing() != null) {
                    cellSpacing.setBeforeAutospacing(tmpSpacing.getBeforeAutospacing());
                }

                if (tmpSpacing.getBeforeLines() != null) {
                    cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
                }

                if (tmpSpacing.getLine() != null) {
                    cellSpacing.setLine(tmpSpacing.getLine());
                }

                if (tmpSpacing.getLineRule() != null) {
                    cellSpacing.setLineRule(tmpSpacing.getLineRule());
                }
            }

            CTInd tmpInd = tmpPPr.getInd();
            if (tmpInd != null) {
                CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd() : cellPPr.addNewInd();
                if (tmpInd.getFirstLine() != null) {
                    cellInd.setFirstLine(tmpInd.getFirstLine());
                }

                if (tmpInd.getFirstLineChars() != null) {
                    cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
                }

                if (tmpInd.getHanging() != null) {
                    cellInd.setHanging(tmpInd.getHanging());
                }

                if (tmpInd.getHangingChars() != null) {
                    cellInd.setHangingChars(tmpInd.getHangingChars());
                }

                if (tmpInd.getLeft() != null) {
                    cellInd.setLeft(tmpInd.getLeft());
                }

                if (tmpInd.getLeftChars() != null) {
                    cellInd.setLeftChars(tmpInd.getLeftChars());
                }

                if (tmpInd.getRight() != null) {
                    cellInd.setRight(tmpInd.getRight());
                }

                if (tmpInd.getRightChars() != null) {
                    cellInd.setRightChars(tmpInd.getRightChars());
                }
            }
        }

    }

    /**
     * 复制段落样式到目标段落
     * @param targetParagraph 目标段落
     * @param sourceParagraph 段落源头
     */
    public static void copyXWPFParagraphCTP(XWPFParagraph targetParagraph, XWPFParagraph sourceParagraph){
        if (targetParagraph == null || sourceParagraph == null) {
            throw new IllegalArgumentException("targetParagraph、sourceParagraph不能为空！");
        }
        targetParagraph.setStyle(sourceParagraph.getStyle());
        targetParagraph.setAlignment(sourceParagraph.getAlignment());
        targetParagraph.setFontAlignment(sourceParagraph.getFontAlignment());
        targetParagraph.setFirstLineIndent(sourceParagraph.getFirstLineIndent());
        targetParagraph.setIndentationLeft(sourceParagraph.getIndentationLeft());
        targetParagraph.setIndentationRight(sourceParagraph.getIndentationRight());
        targetParagraph.setIndentFromLeft(sourceParagraph.getIndentFromLeft());
        targetParagraph.setIndentFromRight(sourceParagraph.getIndentFromRight());
        targetParagraph.setIndentationFirstLine(sourceParagraph.getIndentationFirstLine());
        targetParagraph.setSpacingAfter(sourceParagraph.getSpacingAfter());
        targetParagraph.setSpacingAfterLines(sourceParagraph.getSpacingAfterLines());
        targetParagraph.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
    }

    /**
     * 复制段落样式到目标
     * @param targetRun 目标段落
     * @param sourceRun 段落源头
     */
    public static void copyXWPFRunCTP(XWPFRun targetRun, XWPFRun sourceRun){
        if (targetRun == null || sourceRun == null) {
            throw new IllegalArgumentException("targetRun、sourceRun不能为空！");
        }
        targetRun.getCTR().setRPr(sourceRun.getCTR().getRPr());
    }
}
