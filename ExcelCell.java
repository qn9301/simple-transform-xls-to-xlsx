package xxx.common.entity;

import lombok.Data;

@Data
public class ExcelCell {

    // 第一列
    private int firstColumn;

    // 第一行
    private int firstRow;

    // 最后一列
    private int lastColumn;

    // 最后一列
    private int lastRow;

    // 是否是合饼单元格
    private boolean isMergedRegion;

    //
    private String value;

}

