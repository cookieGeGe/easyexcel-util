package io.github.cookiegege.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExportSheet<T> {

    /**
     * 导出的表名称
     */
    private String sheetName;

    /**
     * 要导出的数据
     */
    private List<T> list;

    /**
     * 导出对象的实体类
     */
    private Class<T> clazz;

}
