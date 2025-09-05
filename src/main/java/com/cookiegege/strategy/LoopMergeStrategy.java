package com.cookiegege.strategy;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * @author JoSuper
 * @date 2025/9/1 14:59
 */
public class LoopMergeStrategy extends AbstractMergeStrategy {

    private int eachRow;   // 每多少行合并一次

    private List<Integer> columnList; // 要合并的列索引（0开始）

    private int mergeLastRow;

    public LoopMergeStrategy(int eachRow, List<Integer> columnList) {
        this.eachRow = eachRow;
        this.columnList = columnList;
        this.mergeLastRow = 0;
    }


    /**
     * merge
     *
     * @param sheet
     * @param cell
     * @param head
     * @param rowIndex
     */
    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer rowIndex) {
        List<String> headNameList = head.getHeadNameList();
        int current = cell.getColumnIndex();
        if (!this.columnList.contains(current)) {
            return;
        }

        int realRowIndex = rowIndex;
        if ((realRowIndex + 1) % eachRow == 0 && realRowIndex != mergeLastRow) {
            this.mergeLastRow = realRowIndex;
            int firstRow = realRowIndex + 1 - eachRow + headNameList.size();
            int lastRow = firstRow + eachRow - 1;
            for (Integer columnIndex : this.columnList) {
                sheet.addMergedRegionUnsafe(new CellRangeAddress(firstRow, lastRow, columnIndex, columnIndex));
            }
        }
    }
}
