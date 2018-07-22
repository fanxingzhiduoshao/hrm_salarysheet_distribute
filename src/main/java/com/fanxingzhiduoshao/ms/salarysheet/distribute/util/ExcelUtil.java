package com.fanxingzhiduoshao.ms.salarysheet.distribute.util;
/*

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

*/
/**
 * excel 导入导出工具类 提供excel 导入导出 业务方法
 *
 * @author wy.wang
 * @version 1.0.0
 *//*

public final class ExcelUtil {
    // 构造器私有 确保不被实例化
    private ExcelUtil() {

    }
    public static final String TYPE_XLS = ".xls";
    public static final String TYPE_XLSX = ".xlsx";
    private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    */
/**
     *
     * @MethodName : excelToList
     *
     * @Description : 将Excel转化为List
     * @param in 承载着Excel的输入流
     * @param type excel文件的后缀名 ".xls","xlsx"
     * @param entityClass List中对象的类型（Excel中的每一行都要转化为该类型的对象）
     * @param fieldMap Excel中的中文列头和类的英文属性的对应关系Map
     * @param <T>
     * @return
     * @throws ExcelException
     * @throws IOException
     *//*

    public static <T> List<T> excelToList(InputStream in, String type, Class<T> entityClass, LinkedHashMap<String, String> fieldMap)
            throws ExcelException, IOException {

        // 定义要返回的list
        List<T> resultList = new ArrayList<T>();
        Workbook workbook = null;
        try {

            // 根据Excel数据源创建WorkBook
            // 创建一个工作薄
            if (TYPE_XLS.equals(type)) {
                workbook = new HSSFWorkbook(in);
            } else if (TYPE_XLSX.equals(type)) {
                // Office 2007+ XML
                workbook = new XSSFWorkbook(in);
            } else {
                throw new ExcelException("文件格式错误！");
            }

            // 获取工作表 只导入第一页工作表
            Sheet sheet = workbook.getSheetAt(0);
            // 获取工作表的有效行数
            int rownums = sheet.getLastRowNum();

            // 如果Excel中没有数据则提示错误
            if (rownums < 1) {
                throw new ExcelException("Excel文件中没有任何数据");
            }

            // 获取Excel中的列名
            Row header = sheet.getRow(0);
            LinkedHashMap<String, Integer> colMap = new LinkedHashMap<String, Integer>();
            for (int i = 0; i < header.getLastCellNum(); i++) {
                Cell cell = header.getCell(i);
                colMap.put(cell.getStringCellValue(), i);
            }
            logger.debug("导入excel第字段列表" + colMap.toString());
            // 判断需要的字段在Excel中是否都存在
            boolean isExist = true;
            for (String cnName : fieldMap.values()) {
                if (!colMap.containsKey(cnName)) {
                    logger.debug("导入的数据表中缺少字段：" + cnName);
                    isExist = false;
                    break;
                }
            }
            if (!isExist) {
                throw new ExcelException("Excel中缺少必要的字段，或字段名称有误");
            }

            // 将sheet转换为list
            for (int i = 1; i <= rownums; i++) {
                // 新建要转换的对象
                T entity;
                entity = entityClass.newInstance();
                // 获取行
                Row row = sheet.getRow(i);
                // 给对象中的字段赋值
                for (Entry<String, String> entry : fieldMap.entrySet()) {
                    // 获取中文字段名
                    String cnNormalName = entry.getValue();
                    // 获取英文字段名
                    String enNormalName = entry.getKey();
                    // 根据中文字段名获取列号
                    int col = colMap.get(cnNormalName);

                    // 获取当前单元格中的内容
                    Cell cell = row.getCell(col);

                    // 给对象赋值
                    try {
                        setFieldValueByName(enNormalName, cell, entity);
                    } catch (IllegalArgumentException | IllegalAccessException e) {
                        logger.debug("第%d行%d列%s字段属性与excel中对应字段%s格式不匹配!", i + 1, col + 1, enNormalName, cnNormalName);
                        String msg = String.format("第%d行%d列%s字段属性与excel中对应字段%s格式不匹配!", i + 1, col + 1, enNormalName,
                                cnNormalName);
                        throw new ExcelException(msg);
                    }
                }

                resultList.add(entity);
            }
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
            // 如果是ExcelException，则直接抛出
            if (e instanceof ExcelException) {
                throw (ExcelException) e;

                // 否则将其它异常包装成ExcelException再抛出
            } else {
                e.printStackTrace();
                throw new ExcelException("导入Excel失败");
            }
        } finally {
            if (workbook != null)
                workbook.close();
            in.close();
        }
        return resultList;
    }

    */
/**
     * 将给定集合list的对象的 包含在FiledMap中的属性插入 至excel 表中，并将excel 结果写入out流中(默认格式为xls)
     *
     * @param list      给定对象的集合<br>
     *                  当集合为空时 返回空excel表
     * @param fieldMap  <p>
     *                  需要插入的属性 和 对应的列名
     *                  </p>
     *                  示例：<br>
     *            LinkedHashMap<String, String> fieldMap = new LinkedHashMap
     *            <String,String>();<br>
     *                  fieldMap.put("question", "问题");<br>
     *                  fieldMap.put("plainText", "回复");<br>
     *                  fieldMap.put("classifyId", "类型编号");<br>
     *                  fieldMap.put("beginTime", "生效时间");<br>
     *                  fieldMap.put("endTime", "失效时间");<br>
     *                  fieldMap.put("characterId", "角色编号");<br>
     *                  fieldMap.put("fQuestion", "主句");<br>
     *                  fieldMap.put("Question.classify.name", "问题分类名"); 支持路径格式<br>
     * @param sheetName 工作表名
     * @param sheetSize 一页工作表 包含行数
     * @param out       输出流
     * @throws Exception
     *//*

    public static <T> void listToExcel(List<T> list, LinkedHashMap<String, String> fieldMap, String sheetName,
                                       int sheetSize, OutputStream out) throws Exception {
        listToExcel(list, TYPE_XLS, fieldMap, sheetName, sheetSize, out);

    }

    */
/**
     * @param list      数据源
     * @param fieldMap  类的英文属性和Excel中的中文列名的对应关系
     * @param sheetSize 每个工作表中记录的最大个数
     * @param response  使用response可以导出到浏览器
     * @throws ExcelException
     * @MethodName : listToExcel
     * @Description : 导出Excel（导出到浏览器，可以自定义工作表的大小）
     *//*

    public static <T> void listToExcel(List<T> list, LinkedHashMap<String, String> fieldMap, String sheetName,
                                       int sheetSize, HttpServletResponse response) throws ExcelException {

        // 设置默认文件名为当前时间：年月日时分秒
        String fileName = sheetName + new SimpleDateFormat("yyyyMMddhhmmss").format(new Date()).toString();

        // 设置response头信息
        response.reset();
        response.setContentType("application/vnd.ms-excel"); // 改成输出excel文件
        response.setHeader("Content-disposition", "attachment; filename=" + fileName + ".xls");

        // 创建工作簿并发送到浏览器
        try {

            OutputStream out = response.getOutputStream();
            listToExcel(list, fieldMap, sheetName, sheetSize, out);

        } catch (Exception e) {
            e.printStackTrace();

            // 如果是ExcelException，则直接抛出
            if (e instanceof ExcelException) {
                throw (ExcelException) e;

                // 否则将其它异常包装成ExcelException再抛出
            } else {
                throw new ExcelException("导出Excel失败");
            }
        }
    }

    */
/**
     * @param list      数据源
     * @param fieldMap  类的英文属性和Excel中的中文列名的对应关系
     * @param sheetSize 每个工作表中记录的最大个数
     * @throws ExcelException
     * @MethodName : listToExcel
     * @Description : 导出Excel（导出到浏览器，可以自定义工作表的大小）
     *//*

    public static <T> void listToExcel(List<T> list, LinkedHashMap<String, String> fieldMap, String sheetName,
                                       int sheetSize, File file) throws ExcelException {

        String type = file.getName().substring(file.getName().lastIndexOf("."));
        // 创建工作簿写入指定文件
        try {

            OutputStream out = new FileOutputStream(file);
            listToExcel(list, type, fieldMap, sheetName, sheetSize, out);

        } catch (Exception e) {
            e.printStackTrace();

            // 如果是ExcelException，则直接抛出
            if (e instanceof ExcelException) {
                throw (ExcelException) e;

                // 否则将其它异常包装成ExcelException再抛出
            } else {
                throw new ExcelException("写入指定文件失败");
            }
        }
    }

    private static <T> void listToExcel(List<T> list, String type, LinkedHashMap<String, String> fieldMap, String sheetName,
                                        int sheetSize, OutputStream out) throws Exception {
        if (sheetSize > 65535 || sheetSize < 1) {
            sheetSize = 65535;
        }

        // 创建一个工作薄
        Workbook workbook = null;
        // 根据Excel数据源创建WorkBook
        // 创建一个工作薄
        if (TYPE_XLS.equals(type)) {
            workbook = new HSSFWorkbook();
        } else if (TYPE_XLSX.equals(type)) {
            // Office 2007+ XML
            workbook = new XSSFWorkbook();
        } else {
            throw new ExcelException("文件格式错误！");
        }

        // 定义 一些 数据保存格式
        Map<String, CellStyle> styles = createCellStyles(workbook);

        if (list.size() == 0 || list == null) {
            Sheet sheet = workbook.createSheet(sheetName);
            fillSheet(sheet, fieldMap, list, 0, list.size() - 1, styles);
            sheet.setDefaultColumnWidth((short) 15);
        }

        // 因为2003的Excel一个工作表最多可以有65536条记录，除去列头剩下65535条
        // 所以如果记录太多，需要放到多个工作表中，其实就是个分页的过程
        // 1.计算一共有多少个工作表
        double sheetNum = Math.ceil(list.size() / new Integer(sheetSize).doubleValue());

        // 2.创建相应的工作表，并向其中填充数据
        for (int i = 0; i < sheetNum; i++) {
            // 如果只有一个工作表的情况
            if (1 == sheetNum) {
                Sheet sheet = workbook.createSheet(sheetName);
                fillSheet(sheet, fieldMap, list, 0, list.size() - 1, styles);
                sheet.setDefaultColumnWidth((short) 15);
                // 有多个工作表的情况
            } else {
                Sheet sheet = workbook.createSheet(sheetName + (i + 1));
                // 获取开始索引和结束索引
                int firstIndex = i * sheetSize;
                int lastIndex = (i + 1) * sheetSize - 1 > list.size() - 1 ? list.size() - 1 : (i + 1) * sheetSize - 1;
                // 填充工作表
                fillSheet(sheet, fieldMap, list, firstIndex, lastIndex, styles);
                sheet.setDefaultColumnWidth((short) 15);
            }
        }

        // 将内容写给定的输出流
        workbook.write(out);
        workbook.close();
        out.close();

    }

    private static HashMap<String, CellStyle> createCellStyles(Workbook workbook) {

        Font headerfont = workbook.createFont();
        headerfont.setFontHeightInPoints((short) 10);
        headerfont.setBold(true);

        Font contentfont = workbook.createFont();
        headerfont.setFontHeightInPoints((short) 10);

        DataFormat format = workbook.createDataFormat();

        HashMap<String, CellStyle> map = new HashMap<String, CellStyle>();

        CellStyle header = workbook.createCellStyle();
        header.setFillForegroundColor(HSSFColor.WHITE.index);
        header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        header.setBorderBottom(BorderStyle.THIN);
        header.setBorderLeft(BorderStyle.THIN);
        header.setBorderRight(BorderStyle.THIN);
        header.setBorderTop(BorderStyle.THIN);
        header.setAlignment(HorizontalAlignment.CENTER);
        header.setFillBackgroundColor(HSSFColor.GREEN.index);
        header.setFont(headerfont);
        header.setDataFormat(format.getFormat("@"));

        CellStyle integer = workbook.createCellStyle();
        integer.setFillForegroundColor(HSSFColor.WHITE.index);
        integer.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        integer.setBorderBottom(BorderStyle.THIN);
        integer.setBorderLeft(BorderStyle.THIN);
        integer.setBorderRight(BorderStyle.THIN);
        integer.setBorderTop(BorderStyle.THIN);
        integer.setAlignment(HorizontalAlignment.GENERAL);
        integer.setFont(contentfont);
        integer.setDataFormat(format.getFormat("0"));

        CellStyle decimals = workbook.createCellStyle();
        decimals.setFillForegroundColor(HSSFColor.WHITE.index);
        decimals.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        decimals.setBorderBottom(BorderStyle.THIN);
        decimals.setBorderLeft(BorderStyle.THIN);
        decimals.setBorderRight(BorderStyle.THIN);
        decimals.setBorderTop(BorderStyle.THIN);
        decimals.setAlignment(HorizontalAlignment.GENERAL);
        decimals.setFont(contentfont);
        decimals.setDataFormat(format.getFormat("0.0"));

        CellStyle date = workbook.createCellStyle();
        date.setFillForegroundColor(HSSFColor.WHITE.index);
        date.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        date.setBorderBottom(BorderStyle.THIN);
        date.setBorderLeft(BorderStyle.THIN);
        date.setBorderRight(BorderStyle.THIN);
        date.setBorderTop(BorderStyle.THIN);
        date.setAlignment(HorizontalAlignment.GENERAL);
        date.setFont(contentfont);
        date.setDataFormat(format.getFormat("yyyy/M/d"));

        CellStyle string = workbook.createCellStyle();
        string.setFillForegroundColor(HSSFColor.WHITE.index);
        string.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        string.setBorderBottom(BorderStyle.THIN);
        string.setBorderLeft(BorderStyle.THIN);
        string.setBorderRight(BorderStyle.THIN);
        string.setBorderTop(BorderStyle.THIN);
        string.setAlignment(HorizontalAlignment.GENERAL);
        string.setFont(contentfont);
        string.setDataFormat(format.getFormat("@"));

        map.put("header", header);
        map.put("integer", integer);
        map.put("decimals", decimals);
        map.put("date", date);
        map.put("string", string);

        return map;
    }

    */
/**
     * 将数据 插入 工作表中
     *
     * @param sheet
     * @param fieldMap
     * @param list
     * @param firstIndex
     * @param lastIndex
     * @throws Exception
     *//*

    private static <T> void fillSheet(Sheet sheet, LinkedHashMap<String, String> fieldMap, List<T> list, int firstIndex,
                                      int lastIndex, Map<String, CellStyle> styles) throws Exception {

        // 定义存属性字段名和excel表列字段名的数组
        String[] enFields = new String[fieldMap.size()];
        String[] cnFields = new String[fieldMap.size()];
        // 填充数组
        int count = 0;
        for (Entry<String, String> entry : fieldMap.entrySet()) {
            enFields[count] = entry.getKey();
            cnFields[count] = entry.getValue();
            count++;
        }
        // 填充表头
        int rowNum = 0;
        Row headline = sheet.createRow(rowNum++);
        Cell cell;
        for (int i = 0; i < cnFields.length; i++) {
            cell = headline.createCell(i, CellType.STRING);// 根据表格行创建单元格
            cell.setCellValue(String.valueOf(cnFields[i]));
            cell.setCellStyle(styles.get("header"));
        }

        logger.debug("导出{}行数据。", list.size());
        // 填入数据
        for (int i = firstIndex; i <= lastIndex; i++) {
            // 创建行
            Row row = sheet.createRow(rowNum++);
            // 获取对象
            T entity = list.get(i);
            // 将对象对应excel 表列的值插入
            for (int j = 0; j < enFields.length; j++) {
                cell = row.createCell(j, CellType.STRING);
                Object value = getFieldValueByNameSequence(enFields[j], entity);
                if (value != null) {

                    if (value instanceof Date) {
                        cell.setCellValue((Date) value);
                        cell.setCellStyle(styles.get("date"));
                        continue;
                    }
                    if (value instanceof Integer) {
                        double num = ((Integer) value).doubleValue();
                        cell.setCellValue(num);
                        cell.setCellStyle(styles.get("integer"));
                        continue;
                    }
                    if (value instanceof Double) {
                        double num = ((Double) value).doubleValue();
                        cell.setCellValue(num);
                        cell.setCellStyle(styles.get("decimals"));
                        continue;
                    }
                    if (value instanceof Long) {

                        cell.setCellStyle(styles.get("integer"));
                        double num = ((Long) value).doubleValue();
                        cell.setCellValue(num);

                        continue;
                    }
                    cell.setCellValue(String.valueOf(value));
                    cell.setCellStyle(styles.get("string"));
                    continue;
                } else {
                    cell.setCellValue("");
                    cell.setCellStyle(styles.get("string"));
                    continue;
                }

            }

        }
    }

    */
/**
     * @param fieldName 字段名
     * @param object    对象
     * @return 字段值
     * @MethodName : getFieldValueByName
     * @Description : 根据字段名获取字段值
     *//*

    private static Object getFieldValueByName(String fieldName, Object object) throws Exception {

        Object value = null;
        Field field = getFieldByName(fieldName, object.getClass());

        if (field != null) {
            field.setAccessible(true);
            Class<?> type = field.getType();
            value = field.get(object);
        } else {
            throw new ExcelException(object.getClass().getSimpleName() + "类不存在字段名 " + fieldName);
        }

        return value;
    }

    */
/**
     * @param fieldName 字段名
     * @param clazz     包含该字段的类
     * @return 字段
     * @MethodName : getFieldByName
     * @Description : 根据字段名获取字段
     *//*

    private static Field getFieldByName(String fieldName, Class<?> clazz) {
        // 拿到本类的所有字段
        Field[] selfFields = clazz.getDeclaredFields();

        // 如果本类中存在该字段，则返回
        for (Field field : selfFields) {
            if (field.getName().equals(fieldName)) {
                return field;
            }
        }

        // 否则，查看父类中是否存在此字段，如果有则返回
        Class<?> superClazz = clazz.getSuperclass();
        if (superClazz != null && superClazz != Object.class) {
            return getFieldByName(fieldName, superClazz);
        }

        // 如果本类和父类都没有，则返回空
        return null;
    }

    */
/**
     * @param fieldNameSequence 带路径的属性名或简单属性名
     * @param object            对象
     * @return 属性值
     * @throws Exception
     * @MethodName : getFieldValueByNameSequence
     * @Description : 根据带路径或不带路径的属性名获取属性值
     * 即接受简单属性名，如userName等，又接受带路径的属性名，如student.department.name等
     *//*

    private static Object getFieldValueByNameSequence(String fieldNameSequence, Object object) throws Exception {

        Object value = null;

        // 将fieldNameSequence进行拆分
        String[] attributes = fieldNameSequence.split("\\.");
        if (attributes.length == 1) {
            value = getFieldValueByName(fieldNameSequence, object);
        } else {
            // 根据属性名获取属性对象
            Object fieldObj = getFieldValueByName(attributes[0], object);
            String subFieldNameSequence = fieldNameSequence.substring(fieldNameSequence.indexOf(".") + 1);
            value = getFieldValueByNameSequence(subFieldNameSequence, fieldObj);
        }
        return value;

    }

    */
/*
     * @MethodName : setFieldValueByName
     *
     * @Description : 根据字段名给对象的字段赋值
     *
     * @param fieldName 字段名
     *
     * @param fieldValue 字段值
     *
     * @param o 对象
     *//*

    private static void setFieldValueByName(String fieldName, Cell cell, Object entity)
            throws IllegalArgumentException, IllegalAccessException {

        Field field = getFieldByName(fieldName, entity.getClass());

        if (field != null) {
            field.setAccessible(true);

            if (cell == null) {
                field.set(entity, null);
            } else {

                CellType cellType = cell.getCellTypeEnum();

                switch (cellType) {
                    case BLANK:
                        // 根据字段类型给字段赋值
                        initWithBlankValue(entity, field);
                        break;
                    case BOOLEAN:
                        boolean value = cell.getBooleanCellValue();
                        initWithBooleanValue(entity, field, value);
                        break;
                    case STRING:
                        //TODO  对时间类型 兼容转换 待完成
                        String value2 = cell.getStringCellValue();
                        initWithStringValue(entity, field, value2);
                        break;
                    case NUMERIC:
                        double value3 = cell.getNumericCellValue();
                        initWithNumericValue(entity, field, cell);
                        break;
                    default:
                }
            }

        } else {
            throw new ExcelException(entity.getClass().getSimpleName() + "类不存在字段名 " + fieldName);
        }
    }

    private static void initWithBlankValue(Object entity, Field field) throws IllegalAccessException {
        Class<?> fieldType = field.getType();
        if (String.class == fieldType) {
            field.set(entity, "");
        } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
            field.set(entity, 0);
        } else if (Long.TYPE == fieldType || Long.class == fieldType) {
            field.set(entity, 0L);
        } else if (Float.TYPE == fieldType || Float.class == fieldType) {
            field.set(entity, 0.0f);
        } else if (Short.TYPE == fieldType || Short.class == fieldType) {
            field.set(entity, (short) 0);
        } else if (Byte.TYPE == fieldType || Byte.class == fieldType) {
            field.set(entity, (byte) 0);
        } else if (Double.TYPE == fieldType || Double.class == fieldType) {
            field.set(entity, 0.0);
        } else if (Boolean.TYPE == fieldType || Boolean.class == fieldType) {
            field.set(entity, false);
        } else if (Character.TYPE == fieldType || Character.class == fieldType) {
            field.set(entity, ' ');
        } else if (Date.class == fieldType) {
            field.set(entity, new Date());
        }
    }

    private static void initWithStringValue(Object entity, Field field, String value) throws IllegalAccessException {
        Class<?> fieldType = field.getType();
        if (String.class == fieldType) {
            field.set(entity, value);
        } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
            field.set(entity, Integer.parseInt(value));
        } else if (Long.TYPE == fieldType || Long.class == fieldType) {
            field.set(entity, Long.parseLong(value));
        } else if (Float.TYPE == fieldType || Float.class == fieldType) {
            field.set(entity, Float.parseFloat(value));
        } else if (Short.TYPE == fieldType || Short.class == fieldType) {
            field.set(entity, Short.parseShort(value));
        } else if (Byte.TYPE == fieldType || Byte.class == fieldType) {
            field.set(entity, Byte.parseByte(value));
        } else if (Double.TYPE == fieldType || Double.class == fieldType) {
            field.set(entity, Double.parseDouble(value));
        }
    }

    private static void initWithNumericValue(Object entity, Field field, Cell cell) throws IllegalAccessException {
        Class<?> fieldType = field.getType();
        double value = cell.getNumericCellValue();
        if (String.class == fieldType) {
            field.set(entity, String.valueOf(value));
        } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
            field.set(entity, (int) value);
        } else if (Long.TYPE == fieldType || Long.class == fieldType) {
            field.set(entity, (long) value);
        } else if (Float.TYPE == fieldType || Float.class == fieldType) {
            field.set(entity, (float) value);
        } else if (Short.TYPE == fieldType || Short.class == fieldType) {
            field.set(entity, (short) value);
        } else if (Byte.TYPE == fieldType || Byte.class == fieldType) {
            field.set(entity, (byte) value);
        } else if (Double.TYPE == fieldType || Double.class == fieldType) {
            field.set(entity, value);
        } else if (Date.class == fieldType) {
            field.set(entity, cell.getDateCellValue());
        }

    }

    private static void initWithBooleanValue(Object entity, Field field, boolean value) throws IllegalAccessException {
        Class<?> fieldType = field.getType();
        if (String.class == fieldType) {
            field.set(entity, Boolean.valueOf(value).toString());
        }
    }

}
*/
