package com.zxz.mcp.mcpslideparser.model;

import java.util.ArrayList;
import java.util.List;


/**
 * 表示 PPT中的一页幻灯片
 */

public class Slide {
    private int slideNumber;       // 幻灯片编号
    private String title;          // 幻灯片标题
    private List<Shape> shapes;    // 幻灯片中的所有形状
    private String notes;          // 幻灯片备注
    private Style backgroundStyle; // 背景样式

    public Slide() {
        this.shapes = new ArrayList<>();
    }

    // Getter和Setter方法
    public int getSlideNumber() {
        return slideNumber;
    }

    public void setSlideNumber(int slideNumber) {
        this.slideNumber = slideNumber;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<Shape> getShapes() {
        return shapes;
    }

    public void addShape(Shape shape) {
        this.shapes.add(shape);
    }

    public String getNotes() {
        return notes;
    }

    public void setNotes(String notes) {
        this.notes = notes;
    }

    public Style getBackgroundStyle() {
        return backgroundStyle;
    }

    public void setBackgroundStyle(Style backgroundStyle) {
        this.backgroundStyle = backgroundStyle;
    }
}