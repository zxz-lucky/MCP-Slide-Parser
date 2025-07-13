package com.zxz.mcp.mcpslideparser.model;

/**
 * 表示文本段落中的一个文本片段
 */

public class TextRun {
    private String text;        // 文本内容
    private Style style;        // 文本样式
    private boolean isBold;     // 是否加粗
    private boolean isItalic;   // 是否斜体
    private boolean isUnderlined; // 是否有下划线
    private String hyperlink;   // 超链接地址

    // Getter和Setter方法
    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }

    public boolean isBold() {
        return isBold;
    }

    public void setBold(boolean bold) {
        isBold = bold;
    }

    public boolean isItalic() {
        return isItalic;
    }

    public void setItalic(boolean italic) {
        isItalic = italic;
    }

    public boolean isUnderlined() {
        return isUnderlined;
    }

    public void setUnderlined(boolean underlined) {
        isUnderlined = underlined;
    }

    public String getHyperlink() {
        return hyperlink;
    }

    public void setHyperlink(String hyperlink) {
        this.hyperlink = hyperlink;
    }
}