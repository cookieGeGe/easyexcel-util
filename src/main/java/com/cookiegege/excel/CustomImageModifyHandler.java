package com.cookiegege.excel;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.ImageData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFShape;

import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.atomic.AtomicReference;

/**
 * @author Wang WenLei
 * @version 1.0
 * @since 2023/6/13 16:22
 */
public class CustomImageModifyHandler implements CellWriteHandler {
    /**
     * 已经处理的Cell
     */
    private final CopyOnWriteArrayList<String> REPEATS = new CopyOnWriteArrayList<>();
    /**
     * 单元格的图片最大张数（每列的单元格图片张数不确定，单元格宽度需按照张数最多的长度来设置）
     */
    private final AtomicReference<Integer> MAX_IMAGE_SIZE = new AtomicReference<>(0);

    /**
     * 标记手动添加的图片，用于排除EasyExcel自动添加的图片
     */
    private final CopyOnWriteArrayList<Integer> CREATE_PIC_INDEX = new CopyOnWriteArrayList<>();

    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, WriteCellData<?> cellData, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        //  在 数据转换成功后 不是头就把类型设置成空
        if (isHead) {
            return;
        }
        //将要插入图片的单元格的type设置为空,下面再填充图片
        if (CollectionUtils.isNotEmpty(cellData.getImageDataList())) {
            cellData.setType(CellDataTypeEnum.EMPTY);
        }
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        //  在 单元格写入完毕后 ，自己填充图片
        if (isHead || CollectionUtils.isEmpty(cellDataList)) {
            return;
        }
        boolean listFlag = false;
        Sheet sheet = cell.getSheet();
        // 此处为ExcelUrlConverterUtil的返回值
        List<ImageData> imageDataList = cellDataList.get(0).getImageDataList();
        if (CollectionUtils.isNotEmpty(imageDataList)) {
            listFlag = true;
        }
        if (!listFlag && imageDataList == null) {
            return;
        }
        String key = cell.getRowIndex() + "_" + cell.getColumnIndex();
        if (REPEATS.contains(key)) {
            return;
        }
        REPEATS.add(key);
        if (imageDataList.size() > MAX_IMAGE_SIZE.get()) {
            MAX_IMAGE_SIZE.set(imageDataList.size());
        }
        // 默认要导出的图片大小为60*60px,60px的行高大约是900,60px列宽大概是248*8
        sheet.getRow(cell.getRowIndex()).setHeight((short) 900);
        sheet.setColumnWidth(cell.getColumnIndex(), listFlag ? 240 * 8 * MAX_IMAGE_SIZE.get() : 240 * 8);

        if (listFlag) {
            for (int i = 0; i < imageDataList.size(); i++) {
                ImageData imageData = imageDataList.get(i);
                if (imageData == null) {
                    continue;
                }
                byte[] image = imageData.getImage();
                this.insertImage(sheet, cell, image, i);
            }
        } else {
            this.insertImage(sheet, cell, imageDataList.get(0).getImage(), 0);
        }

        // 清除EasyExcel自动添加的没有格式的图片
        XSSFDrawing drawingPatriarch = (XSSFDrawing) sheet.getDrawingPatriarch();
        List<XSSFShape> shapes = drawingPatriarch.getShapes();
        for (int i = 0; i < shapes.size(); i++) {
            XSSFShape shape = shapes.get(i);
            if (shape instanceof XSSFPicture && !CREATE_PIC_INDEX.contains(i)) {
                CREATE_PIC_INDEX.add(i);
                XSSFPicture picture = (XSSFPicture) shape;
                picture.resize(0);
            }
        }
    }

    /**
     * 重新插入一个图片
     *
     * @param sheet       Excel页面
     * @param cell        表格元素
     * @param pictureData 图片数据
     * @param i           图片顺序
     */
    private void insertImage(Sheet sheet, Cell cell, byte[] pictureData, int i) {
        int picWidth = Units.pixelToEMU(60);
        int index = sheet.getWorkbook().addPicture(pictureData, HSSFWorkbook.PICTURE_TYPE_PNG);
        CREATE_PIC_INDEX.add(index);
        Drawing<?> drawing = sheet.getDrawingPatriarch();
        if (drawing == null) {
            drawing = sheet.createDrawingPatriarch();
        }
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        // 设置图片坐标
        anchor.setDx1(picWidth * i);
        anchor.setDx2(picWidth + picWidth * i);
        anchor.setDy1(0);
        anchor.setDy2(0);
        //设置图片位置
        int columnIndex = cell.getColumnIndex();
        anchor.setCol1(columnIndex);
        anchor.setCol2(columnIndex);
        int rowIndex = cell.getRowIndex();
        anchor.setRow1(rowIndex);
        anchor.setRow2(rowIndex + 1);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        drawing.createPicture(anchor, index);
    }
}
