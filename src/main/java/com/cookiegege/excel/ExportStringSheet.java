package com.cookiegege.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExportStringSheet {

    /**
     * 导出的表名称
     */
    private String sheetName;

    /**
     * 要导出的数据
     */
    private List<List<String>> list;

    /**
     * 导出头
     */
    private List<List<String>> head;

}
