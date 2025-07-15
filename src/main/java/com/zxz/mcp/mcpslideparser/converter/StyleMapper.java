package com.zxz.mcp.mcpslideparser.converter;


import com.zxz.mcp.mcpslideparser.model.Shape;
import com.zxz.mcp.mcpslideparser.model.Style;
import com.zxz.mcp.mcpslideparser.model.Table;
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
        return """
        
                /* 新增连接图形样式 */
                        .shape.connector {
                            z-index: 2; /* 确保连接线在顶层 */
                            mix-blend-mode: multiply; /* 重叠部分颜色加深 */
                        }
                       \s
                        /* 圆形/线条拼接修正 */
                        .shape[shape-type*="Line"],\s
                        .shape[shape-type*="Arc"] {
                            position: absolute;
                            transform: translateZ(0); /* 强制GPU渲染避免错位 */
                        }
                
                /* 幻灯片容器 - 保持PPT原始尺寸比例 */
                        .slide {
                            position: relative;
                            width: 1920px;  /* 标准PPT宽度 */
                            height: 1080px;  /* 标准PPT高度(16:9) */
                            margin: 20px auto;
                            background: white;
                            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                            overflow: visible;
                        }
                       \s
                        /* 横向布局容器 */
                        .slide-layout-container {
                            position: relative;
                            width: 100%;
                            height: 100%;
                        }
                       \s
                        /* 形状绝对定位 */
                        .shape {
                            position: absolute;
                            margin: 0;
                        }
                       \s
                        /* 图片尺寸控制 */
                        .ppt-image {
                            position: absolute;
                            max-width: none;  /* 禁用响应式限制 */
                            width: auto;
                            height: auto;
                            object-fit: contain;
                        }
                       \s
                        /* 图片容器 */
                        .image-container {
                            position: absolute;
                            overflow: hidden;
                        }
        
        
        
        /* 基础样式 */
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f9f9f9;
        }
        
        /* 幻灯片容器 */
        .slide-container {
            max-width: 900px;
            margin: 0 auto;
            padding: 20px;
        }
        
        /* 单张幻灯片 */
        .slide {
            background: white;
            margin-bottom: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 15px rgba(0,0,0,0.1);
            overflow: hidden;
            page-break-after: always;
        }
        
        /* 幻灯片内容区域 */
        .slide-content {
            padding: 30px;
        }
        
        /* 标题系统 */
        .slide-title {
            color: #2c3e50;
            margin: 0 0 20px 0;
            padding-bottom: 10px;
            border-bottom: 2px solid #eee;
        }
        
        h1.slide-title { font-size: 28px; }
        h2.slide-title { font-size: 24px; }
        h3.slide-title { font-size: 20px; }
        h4.slide-title { font-size: 18px; }
        
        /* 形状容器 */
        .shape {
            margin: 15px 0;
        }
        
        /* 文本段落 */
        .text-run {
            white-space: pre-wrap;
            margin: 5px 0;
        }
        
        /* 列表样式 */
        ul, ol {
            padding-left: 25px;
            margin: 15px 0;
        }
        
        li {
            margin-bottom: 8px;
        }
        
        /* 表格样式 */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        
        th, td {
            padding: 12px 15px;
            text-align: left;
            border: 1px solid #ddd;
        }
        
        th {
            background-color: #f2f2f2;
            font-weight: bold;
        }
        
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        
        /* 图片样式 */
        img {
            max-width: 100%;
            height: auto;
            display: block;
            margin: 10px auto;
        }
        
        /* 备注区域 */
        .slide-notes {
            background: #f8f9fa;
            padding: 15px;
            margin-top: 20px;
            font-size: 14px;
            color: #666;
            border-top: 1px solid #eee;
        }
        
        /* 响应式设计 */
        @media (max-width: 768px) {
            .slide-container {
                padding: 10px;
            }
            
            .slide-content {
                padding: 20px;
            }
            
            .slide-title {
                font-size: 22px;
            }
            
            th, td {
                padding: 8px 10px;
            }
        }
        """;
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

        // 使用PPT中的原始坐标和尺寸
        styleBuilder.append(String.format(
                "left:%.2fpx;top:%.2fpx;width:%.2fpx;height:%.2fpx;",
                shape.getX(), shape.getY(), shape.getWidth(), shape.getHeight()
        ));

        if (style != null) {
            if (style.getFillColor() != null) {
                styleBuilder.append("background-color:").append(style.getFillColor()).append(";");
            }

            if (style.getBorderColor() != null && style.getBorderWidth() > 0) {
                styleBuilder.append("border:")
                        .append(style.getBorderWidth()).append("px solid ")
                        .append(style.getBorderColor()).append(";");
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

        // 添加行高和字间距
        styleBuilder.append("line-height:1.5;letter-spacing:0.5px;");

        return styleBuilder.toString();
    }


    /**
     * 映射图片样式
     */
    public String mapImageStyle(Style style, Shape shape) {
        StringBuilder styleBuilder = new StringBuilder();

        // 使用PPT中的原始图片尺寸
        styleBuilder.append(String.format(
                "width:%.2fpx;height:%.2fpx;",
                shape.getWidth(), shape.getHeight()
        ));

        // 保持宽高比
        styleBuilder.append("object-fit:contain;");

        if (style != null && style.getBorderColor() != null && style.getBorderWidth() > 0) {
            styleBuilder.append("border:")
                    .append(style.getBorderWidth()).append("px solid ")
                    .append(style.getBorderColor()).append(";");
        }

        return styleBuilder.toString();
    }


    /**
     * 映射表格样式
     */
    public String mapTableStyle(Style style, Shape shape) {
        StringBuilder styleBuilder = new StringBuilder();

        // 基础形状样式
        styleBuilder.append(mapShapeStyle(style, shape));

        if (style != null) {
            // 表格特有样式
            if (style.getBackgroundColor() != null) {
                styleBuilder.append("background-color:").append(style.getBackgroundColor()).append(";");
            }
        }

        return styleBuilder.toString();
    }

    /**
     * 映射表格单元格样式
     */
    public String mapTableCellStyle(Style style, Table.Cell cell) {
        if (style == null) {
            return "";
        }

        StringBuilder styleBuilder = new StringBuilder();

        // 文本样式
        if (style.getFontFamily() != null) {
            styleBuilder.append("font-family:").append(style.getFontFamily()).append(";");
        }

        if (style.getFontSize() > 0) {
            styleBuilder.append("font-size:").append(style.getFontSize()).append("px;");
        }

        if (style.getFontColor() != null) {
            styleBuilder.append("color:").append(style.getFontColor()).append(";");
        }

        // 单元格特有样式
        if (style.getFillColor() != null) {
            styleBuilder.append("background-color:").append(style.getFillColor()).append(";");
        }

        if (style.getBorderColor() != null && style.getBorderWidth() > 0) {
            styleBuilder.append("border:").append(style.getBorderWidth()).append("px solid ")
                    .append(style.getBorderColor()).append(";");
        }

        if (style.getAlignment() != null) {
            styleBuilder.append("text-align:").append(style.getAlignment()).append(";");
        }

        // 添加内边距和垂直对齐
        styleBuilder.append("padding:8px 12px;vertical-align:middle;");

        return styleBuilder.toString();
    }

    /**
     * 映射其他形状样式
     */
    public String mapGenericStyle(Style style, Shape shape) {
        StringBuilder styleBuilder = new StringBuilder();

        // 基础形状样式
        styleBuilder.append(mapShapeStyle(style, shape));

        if (style != null) {
            // 特殊形状处理
            if (shape.getShapeType() != null && shape.getShapeType().contains("ARROW")) {
                // 箭头形状特殊样式
                styleBuilder.append("width:0;height:0;border-style:solid;");
                if (style.getFillColor() != null) {
                    styleBuilder.append("border-color:transparent transparent ")
                            .append(style.getFillColor()).append(" transparent;");
                }
            } else {
                // 默认其他形状样式
                if (style.getFillColor() != null) {
                    styleBuilder.append("background-color:").append(style.getFillColor()).append(";");
                }

                // 圆角处理
                if (shape.getShapeType() != null &&
                        (shape.getShapeType().contains("ROUND") || shape.getShapeType().contains("OVAL"))) {
                    styleBuilder.append("border-radius:50%;");
                }
            }

            if (style.getBorderColor() != null && style.getBorderWidth() > 0) {
                styleBuilder.append("border:").append(style.getBorderWidth()).append("px solid ")
                        .append(style.getBorderColor()).append(";");
            }
        }

        return styleBuilder.toString();
    }


}