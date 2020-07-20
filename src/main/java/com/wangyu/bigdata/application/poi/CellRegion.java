package com.wangyu.bigdata.application.poi;

public class CellRegion {

    //记录合并单元格的每行记录起始和终点的行号
    private int startrownum;
    private int endrownum;
    private int startcolumnnum;
    private int endcolumnnum;
    private String value;

    public int getStartcolumnnum() {
        return startcolumnnum;
    }

    public void setStartcolumnnum(int startcolumnnum) {
        this.startcolumnnum = startcolumnnum;
    }

    public int getEndcolumnnum() {
        return endcolumnnum;
    }

    public void setEndcolumnnum(int endcolumnnum) {
        this.endcolumnnum = endcolumnnum;
    }


    public int getStartrownum() {
        return startrownum;
    }

    public void setStartrownum(int startrownum) {
        this.startrownum = startrownum;
    }

    public int getEndrownum() {
        return endrownum;
    }

    public void setEndrownum(int endrownum) {
        this.endrownum = endrownum;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return "CellRegion{" +
                "startrownum=" + startrownum +
                ", endrownum=" + endrownum +
                ", startcolumnnum=" + startcolumnnum +
                ", endcolumnnum=" + endcolumnnum +
                ", value='" + value + '\'' +
                '}';
    }


}
