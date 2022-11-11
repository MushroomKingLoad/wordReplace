import cn.hutool.core.io.IoUtil;
import com.alibaba.fastjson.JSONObject;

import java.io.*;
import java.util.Calendar;
import java.util.LinkedList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        JSONObject paragraph = new JSONObject();
        Calendar calendar = Calendar.getInstance();
        int year = calendar.get(Calendar.YEAR);
        int month = calendar.get(Calendar.MONTH);
        int day = calendar.get(Calendar.DATE);
        paragraph.put("year",year);
        paragraph.put("month",month);
        paragraph.put("day",day);
        paragraph.put("weather","雨");
        paragraph.put("field1","A");
        paragraph.put("field2","B");
        paragraph.put("field3","C");
        paragraph.put("field4","D");
        paragraph.put("field5","E");
        paragraph.put("field6","F");
        paragraph.put("field7","G");
        paragraph.put("class1","动物");
        paragraph.put("class2","植物");
        paragraph.put("class3","非动物也非植物");
        paragraph.put("level1","高等");
        paragraph.put("level2","低等");
        paragraph.put("type1","XDXL-100");
        paragraph.put("type2","SDBX-200");
        paragraph.put("type3","231");
        paragraph.put("type4","665");


        String[] templates = new String[]{"","A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "N"
                , "M", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "S", "Y", "Z"};
        JSONObject paramsTables = new JSONObject();
        List<JSONObject> tableParamsList = new LinkedList<>();
        for (int i = 0; i < 10; i++) {
            JSONObject tableParam = new JSONObject();
            for (int j = 0; j < 7; j++) {
                tableParam.put("tableField"+(j+1),templates[j+1]);
            }
            tableParamsList.add(tableParam);
        }
        paramsTables.put("table1",tableParamsList);
        paramsTables.put("table2",tableParamsList);

        FileInputStream fis = new FileInputStream(new File("./Word套打工具测试模板.docx"));
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        IoUtil.copy(fis,os);
        byte[] arr = os.toByteArray();
        byte[] arrResult = WordReplaceUtil.me().build(arr).appendParagraphsParams(paragraph).appendParamsTables(paramsTables).execute();
        FileOutputStream fos = new FileOutputStream(new File("./result.docx"));
        fos.write(arrResult);
        fos.close();
    }
}
