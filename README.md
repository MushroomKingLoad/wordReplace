# wordReplace

## WordReplace是对添加了占位符的word模板进行文字替换，实现所需效果的工具。从而减少编写word代码的时间，降低学习API成本，提高工作效率，并且能保留word原本的字体样式


### 版本1.0

1. 支持段落
2. 支持静态表格
3. 支持动态表格
4. 自持混合表格

### 使用方式
采用{{xxx}}格式标识占位符，在word模板中增加占位符，后端添加数据执行方法即可。
### 一、段落的套打
#### 直接使用{{xx}} 标识字段即可，word模板中的表现为
![image](https://user-images.githubusercontent.com/55369986/201274816-8c32bc9f-de19-4ec8-b7b4-90e3e8617202.png)
#### code
``` java
        //构建数据
        JSONObject paragraph = new JSONObject();
        Calendar calendar = Calendar.getInstance();
        int year = calendar.get(Calendar.YEAR);
        int month = calendar.get(Calendar.MONTH);
        int day = calendar.get(Calendar.DATE);
        paragraph.put("year",year);
        paragraph.put("month",month);
        paragraph.put("day",day);
        //构建word
        FileInputStream fis = new FileInputStream(new File("./Word套打工具测试模板.docx"));
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        IoUtil.copy(fis,os);
        byte[] arr = os.toByteArray();
        byte[] arrResult = WordReplaceUtil.me().build(arr).appendParagraphsParams(paragraph).execute();
        FileOutputStream fos = new FileOutputStream(new File("./result.docx"));
        fos.write(arrResult);
        fos.close();
```
![image](https://user-images.githubusercontent.com/55369986/201275218-8740ce4d-8b46-41d6-95dd-af23236463a5.png)

### 二、静态表格的套打
#### 与段落相同直接使用{{xx}} 标识字段即可，word模板中的表现为
![image](https://user-images.githubusercontent.com/55369986/201275607-9ec734f4-8a4e-4382-be8c-9eee4abc6223.png)
#### code
``` java
        //构建数据
        JSONObject paragraph = new JSONObject();
        paragraph.put("field1","A");
        paragraph.put("field2","B");
        paragraph.put("field3","C");
        paragraph.put("field4","D");
        paragraph.put("field5","E");
        paragraph.put("field6","F");
        paragraph.put("field7","G");
        //构建word
        FileInputStream fis = new FileInputStream(new File("./Word套打工具测试模板.docx"));
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        IoUtil.copy(fis,os);
        byte[] arr = os.toByteArray();
        //与段落相同使用的appendParagraphsParams方法
        byte[] arrResult = WordReplaceUtil.me().build(arr).appendParagraphsParams(paragraph).execute();
        FileOutputStream fos = new FileOutputStream(new File("./result.docx"));
        fos.write(arrResult);
        fos.close();
```
![image](https://user-images.githubusercontent.com/55369986/201276053-68367176-1db8-4adc-8f0a-e81877473612.png)
### 三、动态表格的套打
#### 在表格头部增加一行第一列使用{{xx}}标识这是一个动态表格，然后再增加一行，在每一列中使用{{xx}} 标识字段即可，word模板中的表现为
![image](https://user-images.githubusercontent.com/55369986/201276376-edf439fe-845b-465f-9967-20afd8bbf3a7.png)
#### code
``` java
        //构建数据
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
        //构建word
        FileInputStream fis = new FileInputStream(new File("./Word套打工具测试模板.docx"));
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        IoUtil.copy(fis,os);
        byte[] arr = os.toByteArray();
        byte[] arrResult = WordReplaceUtil.me().build(arr).appendParamsTables(paramsTables).execute();
        FileOutputStream fos = new FileOutputStream(new File("./result.docx"));
        fos.write(arrResult);
        fos.close();
```
![image](https://user-images.githubusercontent.com/55369986/201276637-fae52b37-29dd-47f2-ba4d-878028387bdb.png)
![image](https://user-images.githubusercontent.com/55369986/201276053-68367176-1db8-4adc-8f0a-e81877473612.png)
### 四、混合表格的套打
#### 静态表格部分{{xx}} 标识字段即可，混合表格部分在表格头部增加一行第一列使用{{xx}}标识这是一个动态表格，然后再增加一行，在每一列中使用{{xx}} 标识字段即可，word模板中的表现为
![image](https://user-images.githubusercontent.com/55369986/201276970-b1a18f97-124f-4c9b-b8fc-b8e611706e86.png)
#### code
``` java
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
```
![image](https://user-images.githubusercontent.com/55369986/201277084-a502dc69-2c21-463a-bb66-281bcde88173.png)

