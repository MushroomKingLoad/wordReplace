# wordReplace

## WordReplace是对添加了占位符的word模板进行文字替换，实现所需效果的工具。从而减少编写word代码的时间，降低学习API成本，提高工作效率。

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
```
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

