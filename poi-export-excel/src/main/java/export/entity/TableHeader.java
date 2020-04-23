package export.entity;

import java.util.List;

public class TableHeader {

    /**
     * 表头文字
     */
    private String headerText = "";
    /**
     * 取值字段
     */
    private String field = "";
    /**
     * 列宽 [defult = 15]
     */
    private Integer width = 15;
    /**
     * 表头背景色
     */
    private String background = "";
    /**
     * 单元格内容居中、向左对齐、向右对齐 [defult = left]
     */
    private String align = "left";
    /**
     * 单元格内容是否换行
     */
    private Boolean wrapText = false;
    /**
     * 多级表头
     */
    private List<TableHeader> children;

    public String getHeaderText() {
        return headerText;
    }

    public void setHeaderText(String headerText) {
        this.headerText = headerText;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getAlign() {
        return align;
    }

    public void setAlign(String align) {
        this.align = align;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public List<TableHeader> getChildren() {
        return children;
    }

    public void setChildren(List<TableHeader> children) {
        this.children = children;
    }

    @Override
    public String toString() {
        return "TableHeader{" +
                "headerText='" + headerText + '\'' +
                ", field='" + field + '\'' +
                ", width=" + width +
                ", background='" + background + '\'' +
                ", children=" + children +
                '}';
    }
}
