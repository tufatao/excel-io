# excel-io
基于CSV与Map的相互转化需求，无需创建相应的数据类。
注意：
* 不提供Pojo的转化支持，因为同类轮子太多。
* 约定大于配置。这里约定sheet中只有一行title，且title从最小列开始
，data行紧跟title行。

### Basic
---
* jdk >= 1.7
* 基于apache poi-ooxml 3.17

### Features
---
* 链式操作，更加直观；
* 读取Csv(File or InputStream) 转化为 List<Map<String, Object>> 或其Json形式；
* 将List<Map<String, Object>> 或其Json形式 转化为Csv 并写入文件或OutputStream；

### How to Use
---
Csv to map list
* read方法支持file，inputStream 类型，及其加密形式，这个没得说，只是做了一下代理；
* toMaplist方法可按需指定sheetIndex，sheetName，titleRowIndex；

#### Example
---
```
List<Map<String, Object>> result;
//默认方式
result = ExcelIO.getReader().read(file).toMaplist();
//指定sheetIndex, titleRowIndex的加密文件
result = ExcelIO.getReader().read(file, password).toMaplist(sheetIndex, titleRowIndex);
```
Map list to local file
* generate方法按需指定sheeName，默认sheetName(sheet1)
，mapList字段中的Object, 数字类型(int, float, double) 已针对处理
  , 其他类型请先自行处理, 或默认按toString()取值
* write方法可按需写入File 或 OutputStream，这个也只是个代理；

#### Example
---
```
//默认方式
ExcelIO.getWriter().generate(mapList).write(file);
//指定sheetName
ExcelIO.getWriter().generate(mapList, "sheetName").write(file);
```
 