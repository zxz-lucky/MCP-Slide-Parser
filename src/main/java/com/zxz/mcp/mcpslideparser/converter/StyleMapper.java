package com.zxz.mcp.mcpslideparser.converter;


import com.zxz.mcp.mcpslideparser.model.Shape;
import com.zxz.mcp.mcpslideparser.model.Style;
import com.zxz.mcp.mcpslideparser.model.TextRun;

/**
 * 将 PPT样式映射为 CSS样式
 */

public class StyleMapper {

    /**
     * 获取全局 CSS样式
     * @return CSS字符串
     */
    public String getGlobalCSS() {
        return "        .text-bold { font-weight: bold; }\n" +
               "        .text-italic { font-style: italic; }\n" +
               "        .text-underlined { text-decoration: underline; }\n";
    }
    
    /**
     * 映射幻灯片样式
     * @param style 样式对象
     * @return CSS样式字符串
     */
    public String mapSlideStyle(Style style) {
        if (style == null) {
            return "";
        }
        
        StringBuilder styleBuilder = new StringBuilder();
        
        if (style.getBackgroundColor() != null) {
            styleBuilder.append("background-color:").append(style.getBackgroundColor()).append(";");
        }
        
        return styleBuilder.toString();
    }
    
    /**
     * 映射形状样式
     * @param style 样式对象
     * @param shape 形状对象
     * @return CSS样式字符串
     */
    public String mapShapeStyle(Style style, Shape shape) {
        StringBuilder styleBuilder = new StringBuilder();
        
        // 位置和大小
        styleBuilder.append(String.format("left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;", 
                                        shape.getX(), shape.getY(), shape.getWidth(), shape.getHeight()));
        
        if (style != null) {
            if (style.getFillColor() != null) {
                styleBuilder.append("background-color:").append(style.getFillColor()).append(";");
            }
            
            if (style.getBorderColor() != null && style.getBorderWidth() > 0) {
                styleBuilder.append("border:")
                          .append(style.getBorderWidth()).append("px solid ")
                          .append(style.getBorderColor()).append(";");
            }
            
            if (style.getOpacity() > 0) {
                styleBuilder.append("opacity:").append(style.getOpacity()).append(";");
            }
        }
        
        return styleBuilder.toString();
    }
    
    /**
     * 映射文本样式
     * @param style 样式对象
     * @param textRun 文本片段对象
     * @return CSS样式字符串
     */
    public String mapTextStyle(Style style, TextRun textRun) {
        if (style == null) {
            return "";
        }
        
        StringBuilder styleBuilder = new StringBuilder();
        
        if (style.getFontFamily() != null) {
            styleBuilder.append("font-family:").append(style.getFontFamily()).append(";");
        }
        
        if (style.getFontSize() > 0) {
            styleBuilder.append("font-size:").append(style.getFontSize()).append("px;");
        }
        
        if (style.getFontColor() != null) {
            styleBuilder.append("color:").append(style.getFontColor()).append(";");
        }
        
        if (style.getBackgroundColor() != null) {
            styleBuilder.append("background-color:").append(style.getBackgroundColor()).append(";");
        }
        
        if (textRun.isBold()) {
            styleBuilder.append("font-weight:bold;");
        }
        
        if (textRun.isItalic()) {
            styleBuilder.append("font-style:italic;");
        }
        
        if (textRun.isUnderlined()) {
            styleBuilder.append("text-decoration:underline;");
        }
        
        return styleBuilder.toString();
    }
}