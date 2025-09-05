package com.cookiegege;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.io.resource.ClassPathResource;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.handler.WriteHandler;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.alibaba.excel.write.metadata.fill.FillWrapper;
import com.cookiegege.convert.ExcelBigNumberConvert;
import com.cookiegege.excel.*;
import com.cookiegege.exception.ExcelException;
import com.cookiegege.strategy.CellMergeStrategy;
import com.cookiegege.util.FileUtils;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.AccessLevel;
import lombok.NoArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.CountDownLatch;

/**
 * Excel相关处理
 *
 * @author JoSuper
 */
@NoArgsConstructor(access = AccessLevel.PRIVATE)
@Slf4j
public class EasyExcelUtil {

    private static final String COMMA_SEPARATOR = ",";

    /**
     * 同步导入(适用于小数据量)
     *
     * @param is 输入流
     * @return 转换后集合
     */
    public static <T> List<T> importExcel(InputStream is, Class<T> clazz) {
        return EasyExcel.read(is).head(clazz).autoCloseStream(false).sheet().doReadSync();
    }


    /**
     * 使用校验监听器 异步导入 同步返回
     *
     * @param is         输入流
     * @param clazz      对象类型
     * @param isValidate 是否 Validator 检验 默认为是
     * @return 转换后集合
     */
    public static <T> ExcelResult<T> importExcel(InputStream is, Class<T> clazz, boolean isValidate) {
        DefaultExcelListener<T> listener = new DefaultExcelListener<>(isValidate);
        EasyExcel.read(is, clazz, listener).sheet().doRead();
        return listener.getExcelResult();
    }

    /**
     * 使用自定义监听器 异步导入 自定义返回
     *
     * @param is       输入流
     * @param clazz    对象类型
     * @param listener 自定义监听器
     * @return 转换后集合
     */
    public static <T> ExcelResult<T> importExcel(InputStream is, Class<T> clazz, ExcelListener<T> listener) {
        EasyExcel.read(is, clazz, listener).sheet().doRead();
        return listener.getExcelResult();
    }

    /**
     * 使用自定义监听器 异步导入所有sheet 自定义返回
     *
     * @param is       输入流
     * @param clazz    对象类型
     * @param listener 自定义监听器
     * @return 转换后集合
     */
    public static <T> List<T> importAllExcel(InputStream is, Class<T> clazz, MergeExcelReadListener<T> listener) {
        return importExcel(is, clazz, listener, true);
    }

    /**
     * 使用自定义监听器 异步导入第一个sheet 自定义返回
     *
     * @param is       输入流
     * @param clazz    对象类型
     * @param listener 自定义监听器
     * @return 转换后集合
     */
    public static <T> List<T> importExcel(InputStream is, Class<T> clazz, MergeExcelReadListener<T> listener) {
        return importExcel(is, clazz, listener, false);
    }

    /**
     * 使用自定义监听器 异步导入 自定义返回
     *
     * @param is       输入流
     * @param clazz    对象类型
     * @param listener 自定义监听器
     * @param readAll  是否读取所有sheet
     * @return 转换后集合
     */
    public static <T> List<T> importExcel(InputStream is, Class<T> clazz, MergeExcelReadListener<T> listener, Boolean readAll) {
        ExcelReaderBuilder excelReaderBuilder;
        if (clazz == null) {
            excelReaderBuilder = EasyExcel.read(is, listener);
        } else {
            excelReaderBuilder = EasyExcel.read(is, clazz, listener);
        }
        if (listener.getHeadRowNumber() != 1) {
            excelReaderBuilder.headRowNumber(listener.getHeadRowNumber());
        }
        excelReaderBuilder = excelReaderBuilder.extraRead(CellExtraTypeEnum.MERGE);
        if (readAll) {
            excelReaderBuilder.doReadAll();
        } else {
            excelReaderBuilder.sheet().doRead();
        }
        HashMap<String, List<CellExtra>> extraMergeInfoMap = listener.getMergeDataMap();

        List<CellExtra> mergeDataList = listener.getMergeDataList();


        //没有合并单元格情况，直接返回即可
        if (isEmpty(mergeDataList)) {
            return listener.getDataList();
        }
        CountDownLatch computerLatch = new CountDownLatch(extraMergeInfoMap.keySet().size());
        List<T> list = Collections.synchronizedList(new ArrayList<>());
        extraMergeInfoMap.forEach((key, value) -> {
            new Thread(() -> {
                //存在有合并单元格时，自动获取值，并校对
                List<T> explainMergeData = explainMergeData(listener.getDataBySheet(key), value, listener.getHeadRowNumber());
                list.addAll(explainMergeData);
                computerLatch.countDown();
            }).start();
        });

        try {
            computerLatch.await();
        } catch (InterruptedException e) {
        }
        return list;
    }

    /**
     * 处理合并单元格
     *
     * @param data               解析数据
     * @param extraMergeInfoList 合并单元格信息
     * @param headRowNumber      起始行
     * @return 填充好的解析数据
     */
    public static <T> List<T> explainMergeData(List<T> data, List<CellExtra> extraMergeInfoList, Integer headRowNumber) {
        //循环所有合并单元格信息
        extraMergeInfoList.forEach(cellExtra -> {
            int firstRowIndex = cellExtra.getFirstRowIndex() - headRowNumber;
            int lastRowIndex = cellExtra.getLastRowIndex() - headRowNumber;
            int firstColumnIndex = cellExtra.getFirstColumnIndex();
            int lastColumnIndex = cellExtra.getLastColumnIndex();
            //获取初始值
            Object initValue = getInitValueFromList(firstRowIndex, firstColumnIndex, data);
            //设置值
            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                for (int j = firstColumnIndex; j <= lastColumnIndex; j++) {
                    setInitValueToList(initValue, i, j, data);
                }
            }
        });
        return data;
    }

    /**
     * 设置合并单元格的值
     *
     * @param filedValue  值
     * @param rowIndex    行
     * @param columnIndex 列
     * @param data        解析数据
     */
    public static <T> void setInitValueToList(Object filedValue, Integer rowIndex, Integer columnIndex, List<T> data) {
        T object = data.get(rowIndex);

        for (Field field : object.getClass().getDeclaredFields()) {
            //提升反射性能，关闭安全检查
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == columnIndex) {
                    try {
                        field.set(object, filedValue);
                        break;
                    } catch (IllegalAccessException e) {
                        log.error("设置合并单元格的值异常：{}", e.getMessage());
                    }
                }
            }
        }
    }


    /**
     * 获取合并单元格的初始值
     * rowIndex对应list的索引
     * columnIndex对应实体内的字段
     *
     * @param firstRowIndex    起始行
     * @param firstColumnIndex 起始列
     * @param data             列数据
     * @return 初始值
     */
    private static <T> Object getInitValueFromList(Integer firstRowIndex, Integer firstColumnIndex, List<T> data) {
        Object filedValue = null;
        T object = data.get(firstRowIndex);
        for (Field field : object.getClass().getDeclaredFields()) {
            //提升反射性能，关闭安全检查
            field.setAccessible(true);
            ExcelProperty annotation = field.getAnnotation(ExcelProperty.class);
            if (annotation != null) {
                if (annotation.index() == firstColumnIndex) {
                    try {
                        filedValue = field.get(object);
                        break;
                    } catch (IllegalAccessException e) {
                        log.error("设置合并单元格的初始值异常：{}", e.getMessage());
                    }
                }
            }
        }
        return filedValue;
    }

    /**
     * 判断集合是否为空
     *
     * @param collection
     * @return
     */
    private static boolean isEmpty(Collection<?> collection) {
        return collection == null || collection.isEmpty();
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param response  响应体
     */
    public static <T> void exportExcel(List<T> list, String sheetName, Class<T> clazz, HttpServletResponse response) {
        try {
            resetResponse(sheetName, response);
            ServletOutputStream os = response.getOutputStream();
            exportExcel(list, sheetName, clazz, false, os);
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        }
    }

    /**
     * 导出excel(多个sheet)
     *
     * @param list      导出数据集合
     * @param excelName 工作表的名称
     * @param merge     是否合并单元格
     * @param response  响应体
     */
    public static <T> void exportExcel(List<ExportSheet> list, String excelName, boolean merge, HttpServletResponse response) {
        ExcelWriter build = null;
        try {
            resetResponse(excelName, response);
            ServletOutputStream os = response.getOutputStream();
            build = EasyExcel.write(os).autoCloseStream(false).build();
            for (int i = 0; i < list.size(); i++) {
                ExportSheet exportSheet = list.get(i);
                List data = exportSheet.getList();
                ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(i, exportSheet.getSheetName());
                sheetBuilder.head(exportSheet.getClazz()).registerConverter(new ExcelBigNumberConvert());
                sheetBuilder.registerWriteHandler(new CustomImageModifyHandler());
                if (merge) {
                    sheetBuilder.registerWriteHandler(new CellMergeStrategy(data, true));
                }
                build.write(data, sheetBuilder.build());
            }
//            exportExcel(list, sheetName, clazz, merge, os);
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        } finally {
            if (build != null) {
                build.finish();
            }
        }
    }

    /**
     * 导出excel(多个sheet)
     *
     * @param list      导出数据集合
     * @param excelName 工作表的名称
     * @param merge     是否合并单元格
     * @param response  响应体
     */
    public static void exportStringExcel(List<ExportStringSheet> list, String excelName, boolean merge, HttpServletResponse response) {
        ExcelWriter build = null;
        try {
            resetResponse(excelName, response);
            ServletOutputStream os = response.getOutputStream();
            build = EasyExcel.write(os).autoCloseStream(false).build();
            for (int i = 0; i < list.size(); i++) {
                ExportStringSheet exportSheet = list.get(i);
                List<List<String>> data = exportSheet.getList();
                ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(i, exportSheet.getSheetName());
                sheetBuilder.head(exportSheet.getHead()).registerConverter(new ExcelBigNumberConvert());
                sheetBuilder.registerWriteHandler(new CustomImageModifyHandler());
                if (merge) {
                    sheetBuilder.registerWriteHandler(new CellMergeStrategy(data, true));
                }
                build.write(data, sheetBuilder.build());
            }
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        } finally {
            if (build != null) {
                build.finish();
            }
        }
    }

    /**
     * 导出excel(多个sheet)
     *
     * @param list      导出数据集合
     * @param excelName 工作表的名称
     * @param merge     是否合并单元格
     * @param response  响应体
     */
    public static void exportStringExcel(
            List<ExportStringSheet> list, String excelName,
            boolean merge,
            HttpServletResponse response,
            List<WriteHandler> handlers
    ) {
        handlers.add(new CustomImageModifyHandler());
        ExcelWriter build = null;
        try {
            resetResponse(excelName, response);
            ServletOutputStream os = response.getOutputStream();
            build = EasyExcel.write(os).autoCloseStream(false).build();
            for (int i = 0; i < list.size(); i++) {
                ExportStringSheet exportSheet = list.get(i);
                List<List<String>> data = exportSheet.getList();
                ExcelWriterSheetBuilder sheetBuilder = EasyExcel.writerSheet(i, exportSheet.getSheetName());
                sheetBuilder.head(exportSheet.getHead()).registerConverter(new ExcelBigNumberConvert());
                handlers.forEach(sheetBuilder::registerWriteHandler);
//                sheetBuilder.registerWriteHandler(new CustomImageModifyHandler());
                if (merge) {
                    sheetBuilder.registerWriteHandler(new CellMergeStrategy(data, true));
                }
                build.write(data, sheetBuilder.build());
            }
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        } finally {
            if (build != null) {
                build.finish();
            }
        }
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param merge     是否合并单元格
     * @param response  响应体
     */
    public static <T> void exportExcel(List<T> list, String sheetName, Class<T> clazz, boolean merge, HttpServletResponse response) {
        try {
            resetResponse(sheetName, response);
            ServletOutputStream os = response.getOutputStream();
            exportExcel(list, sheetName, clazz, merge, os);
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        }
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param merge     是否合并单元格
     * @param response  响应体
     */
    public static <T> void exportExcel(
            List<T> list,
            String sheetName,
            Class<T> clazz,
            boolean merge,
            HttpServletResponse response,
            List<WriteHandler> writeHandlers
    ) {
        try {
            resetResponse(sheetName, response);
            ServletOutputStream os = response.getOutputStream();
            exportExcel(list, sheetName, clazz, merge, os, writeHandlers);
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        }
    }


    public static void exportCustomExcel(List<List<String>> list, String sheetName, List<List<String>> head, boolean merge, HttpServletResponse response) {
        try {
            resetResponse(sheetName, response);
            ServletOutputStream os = response.getOutputStream();
            exportCustomExcel(list, sheetName, head, merge, os);
        } catch (IOException e) {
            throw new ExcelException("导出Excel异常");
        }
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param head      表头
     * @param merge     是否合并单元格
     * @param os        输出流
     */
    public static void exportCustomExcel(List<List<String>> list, String sheetName, List<List<String>> head, boolean merge, OutputStream os) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(os).head(head)
                .autoCloseStream(false)
                .registerWriteHandler(new CustomImageModifyHandler())
                // 自动适配
//            .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                // 大数值自动转换 防止失真
                .registerConverter(new ExcelBigNumberConvert())
                .sheet(sheetName);
        if (merge) {
            // 合并处理器
            builder.registerWriteHandler(new CellMergeStrategy(list, true));
        }
        builder.doWrite(list);
    }

    /**
     * 导出excel
     *
     * @param list          导出数据集合
     * @param sheetName     工作表的名称
     * @param head          表头
     * @param writeHandlers 自定义处理器
     * @param os            输出流
     */
    public static void exportCustomExcel(List<List<String>> list, String sheetName, List<List<String>> head, List<WriteHandler> writeHandlers, OutputStream os) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(os).head(head)
                .autoCloseStream(false)
                .registerWriteHandler(new CustomImageModifyHandler())
                // 自动适配
//            .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
                // 大数值自动转换 防止失真
                .registerConverter(new ExcelBigNumberConvert())
                .sheet(sheetName);
        // 合并处理器
        writeHandlers.forEach(builder::registerWriteHandler);

        builder.doWrite(list);
    }


    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param os        输出流
     */
    public static <T> void exportExcel(List<T> list, String sheetName, Class<T> clazz, OutputStream os) {
        exportExcel(list, sheetName, clazz, false, os);
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param os        输出流
     */
    public static <T> void exportExcel(List<T> list, String sheetName, Class<T> clazz, OutputStream os, List<WriteHandler> handlers) {
        exportExcel(list, sheetName, clazz, false, os, handlers);
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param merge     是否合并单元格
     * @param os        输出流
     */
    public static <T> void exportExcel(List<T> list, String sheetName, Class<T> clazz, boolean merge, OutputStream os) {
        ExcelWriterSheetBuilder builder = EasyExcel.write(os, clazz)
                .autoCloseStream(false)
                .registerWriteHandler(new CustomImageModifyHandler())
                // 自动适配
//            .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()
                // 大数值自动转换 防止失真
                .registerConverter(new ExcelBigNumberConvert())
                .sheet(sheetName);
        if (merge) {
            // 合并处理器
            builder.registerWriteHandler(new CellMergeStrategy(list, true));
        }
        builder.doWrite(list);
    }

    /**
     * 导出excel
     *
     * @param list      导出数据集合
     * @param sheetName 工作表的名称
     * @param clazz     实体类
     * @param merge     是否合并单元格
     * @param os        输出流
     */
    public static <T> void exportExcel(
            List<T> list,
            String sheetName,
            Class<T> clazz,
            boolean merge,
            OutputStream os,
            List<WriteHandler> handlers
    ) {

        ExcelWriterBuilder builder = EasyExcel.write(os, clazz);
        builder.autoCloseStream(false);
        ArrayList<WriteHandler> writeHandlers = new ArrayList<>();
        writeHandlers.add(new CustomImageModifyHandler());
        writeHandlers.addAll(handlers);
        writeHandlers.forEach(builder::registerWriteHandler);

        builder.registerConverter(new ExcelBigNumberConvert());
        ExcelWriterSheetBuilder writerSheetBuilder = builder.sheet(sheetName);
        if (merge) {
            // 合并处理器
            builder.registerWriteHandler(new CellMergeStrategy(list, true));
        }
        writerSheetBuilder.doWrite(list);
    }

    /**
     * 单表多数据模板导出 模板格式为 {.属性}
     *
     * @param filename     文件名
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param data         模板需要的数据
     * @param response     响应体
     */
    public static void exportTemplate(List<Object> data, String filename, String templatePath, HttpServletResponse response) {
        try {
            resetResponse(filename, response);
            ServletOutputStream os = response.getOutputStream();
            exportTemplate(data, templatePath, os);
        } catch (IOException e) {
            throw new RuntimeException("导出Excel异常");
        }
    }

    /**
     * 单表多数据模板导出 模板格式为 {.属性}
     *
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param data         模板需要的数据
     * @param os           输出流
     */
    public static void exportTemplate(List<Object> data, String templatePath, OutputStream os) {
        ClassPathResource templateResource = new ClassPathResource(templatePath);
        ExcelWriter excelWriter = EasyExcel.write(os)
                .withTemplate(templateResource.getStream())
                .autoCloseStream(false)
                // 大数值自动转换 防止失真
                .registerConverter(new ExcelBigNumberConvert())
                .build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollUtil.isEmpty(data)) {
            throw new IllegalArgumentException("数据为空");
        }
        // 单表多数据导出 模板格式为 {.属性}
        for (Object d : data) {
            excelWriter.fill(d, writeSheet);
        }
        excelWriter.finish();
    }

    /**
     * 多表多数据模板导出 模板格式为 {key.属性}
     *
     * @param filename     文件名
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param data         模板需要的数据
     * @param response     响应体
     */
    public static void exportTemplateMultiList(Map<String, Object> data, String filename, String templatePath, HttpServletResponse response) {
        try {
            resetResponse(filename, response);
            ServletOutputStream os = response.getOutputStream();
            exportTemplateMultiList(data, templatePath, os);
        } catch (IOException e) {
            throw new RuntimeException("导出Excel异常");
        }
    }

    /**
     * 多表多数据模板导出 模板格式为 {key.属性}
     *
     * @param templatePath 模板路径 resource 目录下的路径包括模板文件名
     *                     例如: excel/temp.xlsx
     *                     重点: 模板文件必须放置到启动类对应的 resource 目录下
     * @param data         模板需要的数据
     * @param os           输出流
     */
    public static void exportTemplateMultiList(Map<String, Object> data, String templatePath, OutputStream os) {
        ClassPathResource templateResource = new ClassPathResource(templatePath);
        ExcelWriter excelWriter = EasyExcel.write(os)
                .withTemplate(templateResource.getStream())
                .autoCloseStream(false)
                // 大数值自动转换 防止失真
                .registerConverter(new ExcelBigNumberConvert())
                .build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        if (CollUtil.isEmpty(data)) {
            throw new IllegalArgumentException("数据为空");
        }
        for (Map.Entry<String, Object> map : data.entrySet()) {
            // 设置列表后续还有数据
            FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
            if (map.getValue() instanceof Collection) {
                // 多表导出必须使用 FillWrapper
                excelWriter.fill(new FillWrapper(map.getKey(), (Collection<?>) map.getValue()), fillConfig, writeSheet);
            } else {
                excelWriter.fill(map.getValue(), writeSheet);
            }
        }
        excelWriter.finish();
    }

    /**
     * 重置响应体
     */
    private static void resetResponse(String sheetName, HttpServletResponse response) throws UnsupportedEncodingException {
        String filename = encodingFilename(sheetName);
        FileUtils.setAttachmentResponseHeader(response, filename);
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
    }

    /**
     * 解析导出值 0=男,1=女,2=未知
     *
     * @param propertyValue 参数值
     * @param converterExp  翻译注解
     * @param separator     分隔符
     * @return 解析后值
     */
    public static String convertByExp(String propertyValue, String converterExp, String separator) {
        StringBuilder propertyString = new StringBuilder();
        String[] convertSource = converterExp.split(COMMA_SEPARATOR);
        for (String item : convertSource) {
            String[] itemArray = item.split("=");
            if (StrUtil.containsAny(propertyValue, separator)) {
                for (String value : propertyValue.split(separator)) {
                    if (itemArray[0].equals(value)) {
                        propertyString.append(itemArray[1] + separator);
                        break;
                    }
                }
            } else {
                if (itemArray[0].equals(propertyValue)) {
                    return itemArray[1];
                }
            }
        }
        return StringUtils.stripEnd(propertyString.toString(), separator);
    }

    /**
     * 反向解析值 男=0,女=1,未知=2
     *
     * @param propertyValue 参数值
     * @param converterExp  翻译注解
     * @param separator     分隔符
     * @return 解析后值
     */
    public static String reverseByExp(String propertyValue, String converterExp, String separator) {
        StringBuilder propertyString = new StringBuilder();
        String[] convertSource = converterExp.split(COMMA_SEPARATOR);
        for (String item : convertSource) {
            String[] itemArray = item.split("=");
            if (StringUtils.containsAny(propertyValue, separator)) {
                for (String value : propertyValue.split(separator)) {
                    if (itemArray[1].equals(value)) {
                        propertyString.append(itemArray[0] + separator);
                        break;
                    }
                }
            } else {
                if (itemArray[1].equals(propertyValue)) {
                    return itemArray[0];
                }
            }
        }
        return StringUtils.stripEnd(propertyString.toString(), separator);
    }

    /**
     * 编码文件名
     */
    public static String encodingFilename(String filename) {
        return IdUtil.fastSimpleUUID() + "_" + filename + ".xlsx";
    }

}
