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

    private byte[] imageData;         // 图片数据(针对图片形状)
    private String imageType;         // 图片类型(针对图片形状)
    private String name;              // 形状名称
    private String shapeType;         // 具体形状类型(针对SHAPE类型)
    private Table table;              // 表格数据(针对表格形状)

    private String altText;       // 替代文本


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

    public byte[] getImageData() {
        return imageData;
    }

    public void setImageData(byte[] imageData) {
        this.imageData = imageData;
    }

    public String getImageType() {
        return imageType;
    }

    public void setImageType(String imageType) {
        this.imageType = imageType;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getShapeType() {
        return shapeType;
    }

    public void setShapeType(String shapeType) {
        this.shapeType = shapeType;
    }

    public Table getTable() {
        return table;
    }

    public void setTable(Table table) {
        this.table = table;
    }

    public String getAltText() {
        return altText;
    }

    public void setAltText(String altText) {
        this.altText = altText;
    }


}