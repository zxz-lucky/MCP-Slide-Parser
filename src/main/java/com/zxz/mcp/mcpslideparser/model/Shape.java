package com.zxz.mcp.mcpslideparser.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 表示幻灯片中的一个形状元素
 */

public class Shape {
    public enum ShapeType {
        TEXT_BOX, IMAGE, TABLE, CHART, SHAPE
    }

    private String id;                // 形状ID
    private ShapeType type;           // 形状类型
    private List<TextRun> textRuns;   // 文本内容(针对文本框)
    private Style style;              // 形状样式
    private double x;                 // X坐标
    private double y;                 // Y坐标
    private double width;             // 宽度
    private double height;            // 高度

    public Shape() {
        this.textRuns = new ArrayList<>();
    }

    // Getter和Setter方法
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public ShapeType getType() {
        return type;
    }

    public void setType(ShapeType type) {
        this.type = type;
    }

    public List<TextRun> getTextRuns() {
        return textRuns;
    }

    public void addTextRun(TextRun textRun) {
        this.textRuns.add(textRun);
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }

    public double getX() {
        return x;
    }

    public void setX(double x) {
        this.x = x;
    }

    public double getY() {
        return y;
    }

    public void setY(double y) {
        this.y = y;
    }

    public double getWidth() {
        return width;
    }

    public void setWidth(double width) {
        this.width = width;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }
}