import cn.hutool.core.io.IoUtil;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

/**
 * @author dengwenqi
 */
public class WordReplaceUtil {
    private XWPFDocument doc;
    private JSONObject paragraphsParams = new JSONObject();
    private JSONObject paramsTables = new JSONObject();
    private List<JSONObject> needDealTable = new LinkedList<>();

    public static WordReplaceUtil me() {
        return new WordReplaceUtil();
    }

    public WordReplaceUtil build(byte[] array) throws Exception {
        try (BufferedInputStream is = new BufferedInputStream(new ByteArrayInputStream(array))) {
            this.doc = new XWPFDocument(OPCPackage.open(is));
            return this;
        }
    }


    public WordReplaceUtil appendParagraphsParams(JSONObject paragraphsParams) {
        this.paragraphsParams = paragraphsParams;
        return this;
    }

    public WordReplaceUtil appendParamsTables(JSONObject paramsTables) {
        this.paramsTables = paramsTables;
        return this;
    }

    public byte[] execute() throws Exception {
        dealParagraphsParams(this.paragraphsParams);
        dealTableParams(this.paragraphsParams, this.paramsTables);
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            this.doc.write(os);
            return os.toByteArray();
        }
    }

    /**
     * 处理段落数据
     *
     * @param paragraphsParams 段落数据
     */
    private void dealParagraphsParams(JSONObject paragraphsParams) {
        if (null != paragraphsParams && !paragraphsParams.isEmpty()) {
            List<XWPFParagraph> parasList = doc.getParagraphs();
            replaceParagraphs(parasList, paragraphsParams);
        }
    }

    /**
     * 处理表格数据
     *
     * @param paragraphsParams 段落数据
     * @param paramsTables     表格数据
     * @throws Exception Exception
     */
    private void dealTableParams(JSONObject paragraphsParams, JSONObject paramsTables) throws Exception {
        List<XWPFTable> xwpfTableList = this.doc.getTables();
        for (int i = 0; i < xwpfTableList.size(); i++) {
            XWPFTable table = xwpfTableList.get(i);
            List<XWPFTableRow> rows = table.getRows();
            for (int j = 0; j < rows.size(); j++) {
                XWPFTableRow currentRow = rows.get(j);
                for (XWPFTableCell cell : currentRow.getTableCells()) {
                    List<XWPFParagraph> paragraphsList = cell.getParagraphs();
                    String wordParamKey;
                    if (null != (wordParamKey = isActiveTable(paragraphsList, paramsTables))) {
                        JSONObject tableProp = new JSONObject();
                        tableProp.put("tableIndex", i);
                        tableProp.put("rowIndex", j);
                        tableProp.put("wordParamKey", wordParamKey);
                        needDealTable.add(tableProp);
                        j++;
                        break;
                    }
                    this.replaceParagraphs(paragraphsList, paragraphsParams);

                }
            }
        }

        this.insertActiveTableParams(paramsTables);
    }

    /**
     * 插入动态表数据
     *
     * @param paramsTables 表格数据
     * @throws Exception Exception
     */
    private void insertActiveTableParams(JSONObject paramsTables) throws Exception {
        List<XWPFTable> xwpfTableList = this.doc.getTables();
        for (JSONObject item : needDealTable
        ) {
            int tableIndex = item.getInteger("tableIndex");
            int rowIndex = item.getInteger("rowIndex");
            String wordParamKey = item.getString("wordParamKey");
            XWPFTable table = xwpfTableList.get(tableIndex);
            List<JSONObject> tableParams = paramsTables.getObject(wordParamKey, List.class);
            XWPFTableRow templateRow = table.getRows().get(rowIndex + 1);
            InputStream is = templateRow.getCtRow().newInputStream();
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            IoUtil.copy(is, bos);
            byte[] bytes = bos.toByteArray();
            bos.close();
            CTRow templateCtRow;
            XWPFTableRow currentRow;
            table.removeRow(rowIndex + 1);
            table.removeRow(rowIndex);
            for (int j = 0; j < tableParams.size(); j++) {
                JSONObject tableParam = tableParams.get(j);
                templateCtRow = CTRow.Factory.parse(new ByteArrayInputStream(bytes));
                currentRow = new XWPFTableRow(templateCtRow, table);
                for (XWPFTableCell cell : currentRow.getTableCells()) {
                    replaceParagraphs(cell.getParagraphs(), tableParam);
                }
                table.addRow(currentRow, j + rowIndex);
            }
        }
    }

    /**
     * 判断是否动态表
     *
     * @param paragraphsList 段落集合
     * @param paramsTables   表格数据
     * @return key值
     */
    private String isActiveTable(List<XWPFParagraph> paragraphsList, JSONObject paramsTables) {
        for (XWPFParagraph paragraph : paragraphsList) {
            List<XWPFRun> runs = paragraph.getRuns();
            String allRunText = "";
            for (XWPFRun run : runs) {
                String runText = run.getText(run.getTextPosition());
                allRunText += null == runText ? "" : runText;
            }
            String paramKey = getWordParamKey(allRunText);
            if (paramsTables.containsKey(paramKey)) {
                return paramKey;
            }
        }
        return null;
    }


    /**
     * 获取word文档中的参数
     *
     * @param allRunText run中所有的内容
     * @return 文档中的参数
     */
    private String getWordParamKey(String allRunText) {
        String oneParaString = allRunText;
        if (StringUtils.isNotBlank(oneParaString)) {
            //word文档会根据文字输入的顺序对run进行进组，字符拆分问题
            String mid;
            if (oneParaString.contains("}}") && !"}}".equals(oneParaString)) {
                mid = oneParaString.substring(!oneParaString.contains("{{") ? 0 : oneParaString.indexOf("{{"), oneParaString.indexOf("}}"));
                oneParaString = mid;
            }
            oneParaString = oneParaString.replaceAll("\\{\\{", "");
            oneParaString = oneParaString.replaceAll("}}", "");
        }

        return oneParaString;
    }

    /**
     * 获取word文档中参数替换后的值
     *
     * @param allRunText       run中所有的内容
     * @param paragraphsParams 段落数据
     * @return word文档中参数替换后的值
     */
    private String getWordParamsValue(String allRunText, JSONObject paragraphsParams) {
        String oneParaString = allRunText;
        oneParaString = getWordParamsValueDfs(oneParaString, paragraphsParams);
        return oneParaString;
    }

    private String getWordParamsValueDfs(String oneParaString, JSONObject paragraphsParams) {
        if (StringUtils.isNotBlank(oneParaString)) {
            //word文档会根据文字输入的顺序对run进行分组，解决字符拆分问题
            String front = "";
            String prev = "";
            String mid;
            if (oneParaString.contains("}}") && !"}}".equals(oneParaString)) {
                int leftIndex = !oneParaString.contains("{{") ?
                        0 : oneParaString.indexOf("{{") > oneParaString.indexOf("}}") ? 0 : oneParaString.indexOf("{{");

                prev = oneParaString.substring(0, leftIndex);
                leftIndex = leftIndex > oneParaString.indexOf("}}") ? 0 : leftIndex;
                mid = oneParaString.substring(leftIndex, oneParaString.indexOf("}}"));
                front = oneParaString.substring(oneParaString.indexOf("}}") + 2);
                oneParaString = mid;
            }
            oneParaString = oneParaString.replaceAll("\\{\\{", "");
            oneParaString = oneParaString.replaceAll("}}", "");
            oneParaString = paragraphsParams.containsKey(oneParaString)
                    ? oneParaString.replaceAll(oneParaString, null == paragraphsParams.getString(oneParaString)?"":paragraphsParams.getString(oneParaString)) : oneParaString;            oneParaString = prev + oneParaString + front;
            if (oneParaString.contains("}}") || oneParaString.contains("{{")) {
                oneParaString = getWordParamsValueDfs(oneParaString, paragraphsParams);
            }
        }
        return oneParaString;
    }

    /**
     * 替换
     *
     * @param paragraphsList   段落集合
     * @param paragraphsParams 段落数据
     */
    private void replaceParagraphs(List<XWPFParagraph> paragraphsList, JSONObject paragraphsParams) {
        for (XWPFParagraph paragraph : paragraphsList) {
            List<XWPFRun> runs = paragraph.getRuns();
            String allRunText = "";
            for (XWPFRun run : runs) {
                String runText = run.getText(run.getTextPosition());
                allRunText += null == runText ? "" : runText;
                run.setText("", 0);
            }
            if (StringUtils.isNotBlank(allRunText)) {
                runs.get(0).setText(getWordParamsValue(allRunText, paragraphsParams), 0);
            }
        }
    }
}
