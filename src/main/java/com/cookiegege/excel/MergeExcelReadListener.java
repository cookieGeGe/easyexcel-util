package com.cookiegege.excel;

import cn.hutool.core.util.StrUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.CellExtra;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.stream.Collectors;

@EqualsAndHashCode(callSuper = true)
@Data
public abstract class MergeExcelReadListener<T> extends AnalysisEventListener<T> {

    /**
     * 存储每张表的原始数据
     */
    public HashMap<String, List<T>> dataMap = new HashMap<>();

    /**
     * 存储每个sheet合并行的数据
     */
    public HashMap<String, List<CellExtra>> mergeDataMap = new HashMap<>();

    /**
     * 文件头的行数
     */
    private Integer headRowNumber = 1;

    public MergeExcelReadListener() {
        dataMap = new HashMap<>();
        mergeDataMap = new HashMap<>();
    }

    /**
     * 添加数据到原始数据
     * @param analysisContext
     * @param lineData
     */
    public void addToData(AnalysisContext analysisContext, T lineData) {
        String sheetName = analysisContext.readSheetHolder().getSheetName();
        if (!dataMap.containsKey(sheetName)){
            dataMap.put(sheetName, new ArrayList<>());
        }
        dataMap.get(sheetName).add(lineData);
    }

    /**
     * 添加数据到合并行数据
     * @param analysisContext
     * @param cellExtra
     */
    public void addToMergeData(AnalysisContext analysisContext, CellExtra cellExtra) {
        String sheetName = analysisContext.readSheetHolder().getSheetName();
        if (!mergeDataMap.containsKey(sheetName)){
            mergeDataMap.put(sheetName, new ArrayList<>());
        }
        mergeDataMap.get(sheetName).add(cellExtra);
    }

    /**
     * 获取所有的原始数据
     *
     * @return
     */
    public List<T> getDataList() {
        return dataMap.values().stream().flatMap(Collection::stream).collect(Collectors.toList());
    }

    /**
     * 获取所有的合并行的数据
     *
     * @return
     */
    public List<CellExtra> getMergeDataList() {
        return mergeDataMap.values().stream().flatMap(Collection::stream).collect(Collectors.toList());
    }

    /**
     * 获取指定sheet页的数据
     *
     * @param sheetName sheet页的名称
     * @return
     */
    public List<T> getDataBySheet(String sheetName) {
        return dataMap.get(sheetName);
    }

    /**
     * 获取指定sheet页的数据，带默认值
     *
     * @param sheetName   sheet页的名称
     * @param defaultList 默认数据
     * @return
     */
    public List<T> getDataBySheet(String sheetName, List<T> defaultList) {
        return dataMap.getOrDefault(sheetName, defaultList);
    }

    /**
     * 获取指定sheet页的合并行数据
     *
     * @param sheetName sheet页的名称
     * @return
     */
    public List<CellExtra> getMergeDataBySheet(String sheetName) {
        return mergeDataMap.get(sheetName);
    }

    /**
     * 获取指定sheet页的合并行数据，带默认值
     *
     * @param sheetName   sheet页的名称
     * @param defaultList 默认数据
     * @return
     */
    public List<CellExtra> getMergeDataBySheet(String sheetName, List<CellExtra> defaultList) {
        return mergeDataMap.getOrDefault(sheetName, defaultList);
    }

    /**
     * 公共的sheet页读取完后的操作
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        String msg = StrUtil.format("{} read end!", analysisContext.readSheetHolder().getSheetName());
        System.out.println(msg);
    }

    /**
     * 原始获取结果的方法
     * @return
     */
    public ExcelResult<T> getExcelResult() {
        return null;
    }

    public Integer getHeadRowNumber() {
        return headRowNumber;
    }
}
