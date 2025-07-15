package com.zxz.mcp.mcpslideparser.converter;


import com.zxz.mcp.mcpslideparser.model.Shape;
import com.zxz.mcp.mcpslideparser.model.Slide;
import com.zxz.mcp.mcpslideparser.model.Table;
import com.zxz.mcp.mcpslideparser.model.TextRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Base64;
import java.util.List;

/**
 * 将幻灯片数据转换为 HTML格式
 */

public class HTMLConverter {

    private static final Logger logger = LoggerFactory.getLogger(HTMLConverter.class);
    private final StyleMapper styleMapper;
    
    public HTMLConverter() {
        this.styleMapper = new StyleMapper();
    }
    
    /**
     * 将幻灯片列表转换为完整 HTML文档
     * @param slides 幻灯片列表
     * @return HTML字符串
     */
    public String convertToHTML(List<Slide> slides) {
        logger.info("开始将幻灯片转换为HTML");
        
        StringBuilder htmlBuilder = new StringBuilder();
        
        // HTML 头部
        htmlBuilder.append("<!DOCTYPE html>\n")
                .append("<html>\n")
                .append("<head>\n")
                .append("    <meta charset=\"UTF-8\">\n")
                .append("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n")
                .append("    <title>PPT to HTML Conversion</title>\n")
                .append("    <style>\n")
                .append("        body { font-family: 'Arial', sans-serif; line-height: 1.6; color: #333; background: #f5f5f5; }\n")
                .append("        .slide-container { max-width: 900px; margin: 0 auto; padding: 20px; }\n")
                .append("        .slide { background: white; margin-bottom: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden; }\n")
                .append("        .slide-content { padding: 25px; }\n")
                .append("        .slide-title { font-size: 28px; margin: 0 0 20px 0; color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; }\n")
                .append("        .shape { margin: 15px 0; }\n")
                .append("        .text-run { white-space: pre-wrap; margin: 5px 0; }\n")
                .append("        .slide-notes { background: #f8f9fa; padding: 15px; margin-top: 20px; font-size: 14px; color: #666; }\n")
                .append("        ul, ol { padding-left: 25px; margin: 10px 0; }\n")
                .append("        li { margin-bottom: 8px; }\n")
                .append("        table { width: 100%; border-collapse: collapse; margin: 15px 0; }\n")
                .append("        th, td { padding: 12px; text-align: left; border: 1px solid #ddd; }\n")
                .append("        th { background-color: #f2f2f2; }\n")
                .append("        img { max-width: 100%; height: auto; display: block; }\n")
                .append(styleMapper.getGlobalCSS())
                .append("    </style>\n")
                .append("</head>\n")
                .append("<body>\n")
                .append("<div class=\"slide-container\">\n");
        
        // 幻灯片内容
        for (Slide slide : slides) {
            htmlBuilder.append(convertSlideToHTML(slide));
        }
        
        // HTML 尾部
        htmlBuilder.append("</body>\n")
                  .append("</html>");
        
        logger.info("HTML转换完成");
        return htmlBuilder.toString();
    }
    
    /**
     * 将单个幻灯片转换为 HTML片段
     * @param slide 幻灯片对象
     * @return HTML字符串
     */
    private String convertSlideToHTML(Slide slide) {
        StringBuilder slideBuilder = new StringBuilder();

        // 幻灯片容器 - 添加横向布局容器
        slideBuilder.append(String.format(
                "<div class=\"slide\" id=\"slide-%d\" style=\"%s\">\n",
                slide.getSlideNumber(),
                styleMapper.mapSlideStyle(slide.getBackgroundStyle())
        ));

        // 添加横向布局容器
        slideBuilder.append("    <div class=\"slide-layout-container\">\n");

        // 标题处理
        String title = slide.getTitle();
        if (title != null && !title.trim().isEmpty()) {
            int headerLevel = detectHeaderLevel(title);
            slideBuilder.append(String.format(
                    "        <h%d class=\"slide-title\">%s</h%d>\n",
                    headerLevel,
                    escapeHTML(title),
                    headerLevel
            ));
        }

        // 形状处理 - 保持原有位置
        for (Shape shape : slide.getShapes()) {
            slideBuilder.append(convertShapeToHTML(shape));
        }

        slideBuilder.append("    </div>\n"); // 结束横向布局容器

        // 备注处理
        if (slide.getNotes() != null && !slide.getNotes().isEmpty()) {
            slideBuilder.append("    <div class=\"slide-notes\">\n")
                    .append("        <h4>Notes:</h4>\n")
                    .append("        <p>").append(escapeHTML(slide.getNotes().replace("\n", "<br>"))).append("</p>\n")
                    .append("    </div>\n");
        }

        slideBuilder.append("</div>\n");
        return slideBuilder.toString();
    }


    // 新增方法：检测标题层级
    private int detectHeaderLevel(String title) {
        if (title.startsWith("# ")) return 1;
        if (title.startsWith("## ")) return 2;
        if (title.startsWith("### ")) return 3;
        if (title.startsWith("#### ")) return 4;
        return 2; // 默认h2
    }

    
    /**
     * 将形状转换为 HTML片段
     * @param shape 形状对象
     * @return HTML字符串
     */
    private String convertShapeToHTML(Shape shape) {
        StringBuilder shapeBuilder = new StringBuilder();

        // 使用绝对定位保持原有位置
        shapeBuilder.append(String.format(
                "<div class=\"shape shape-%s\" id=\"shape-%s\" style=\"%s\">\n",
                shape.getType().name().toLowerCase(),
                shape.getId(),
                styleMapper.mapShapeStyle(shape.getStyle(), shape)
        ));

        switch (shape.getType()) {
            case TEXT_BOX:
                for (TextRun textRun : shape.getTextRuns()) {
                    shapeBuilder.append(convertTextRunToHTML(textRun));
                }
                break;
            case IMAGE:
                shapeBuilder.append(convertImageToHTML(shape));
                break;
            case TABLE:
                shapeBuilder.append(convertTableToHTML(shape));
                break;
            default:
                shapeBuilder.append(convertGenericToHTML(shape));
        }

        shapeBuilder.append("</div>\n");

        return shapeBuilder.toString();
    }
    
    /**
     * 将文本片段转换为 HTML片段
     * @param textRun 文本片段对象
     * @return HTML字符串
     */
    private String convertTextRunToHTML(TextRun textRun) {
        StringBuilder runBuilder = new StringBuilder();

        String tag = "span";
        if (textRun.getHyperlink() != null) {
            tag = "a";
        }

        // 检测是否是列表项
        boolean isListItem = textRun.getText() != null &&
                (textRun.getText().startsWith("- ") ||
                        textRun.getText().matches("^\\d+\\.\\s.+"));

        if (isListItem) {
            runBuilder.append("<li class=\"ppt-list-item\" style=\"")
                    .append(styleMapper.mapTextStyle(textRun.getStyle(), textRun))
                    .append("\">")
                    .append(escapeHTML(textRun.getText().replaceFirst("^[-\\d]+\\.?\\s+", "")));
        } else {
            runBuilder.append(String.format(
                    "<%s class=\"text-run\" style=\"%s\"%s>%s</%s>",
                    tag,
                    styleMapper.mapTextStyle(textRun.getStyle(), textRun),
                    textRun.getHyperlink() != null ? " href=\"" + escapeHTML(textRun.getHyperlink()) + "\"" : "",
                    escapeHTML(textRun.getText()),
                    tag
            ));
        }

        return runBuilder.toString();
    }


    /**
     * 转换图片为HTML
     */
    private String convertImageToHTML(Shape shape) {
        StringBuilder imageBuilder = new StringBuilder();

        if (shape.getImageData() != null) {
            String base64Image = Base64.getEncoder().encodeToString(shape.getImageData());
            String imageType = shape.getImageType() != null ?
                    shape.getImageType().replace("image/", "") : "png";

            imageBuilder.append(String.format(
                    "<img src=\"data:image/%s;base64,%s\" alt=\"%s\" class=\"ppt-image\" style=\"%s\" />",
                    imageType,
                    base64Image,
                    escapeHTML(shape.getAltText() != null ? shape.getAltText() : ""),
                    styleMapper.mapImageStyle(shape.getStyle(), shape)
            ));
        } else {
            imageBuilder.append(String.format(
                    "<div class=\"image-placeholder\" style=\"%s\">[Image]</div>",
                    styleMapper.mapImageStyle(shape.getStyle(), shape)
            ));
        }

        return imageBuilder.toString();
    }


    /**
     * 转换表格为HTML
     */
    private String convertTableToHTML(Shape shape) {
        StringBuilder tableBuilder = new StringBuilder();

        if (shape.getTable() != null) {
            Table table = shape.getTable();
            tableBuilder.append("<table>\n");

            // 表头处理
            boolean hasHeader = false;
            // 检查第一行是否有特殊样式可以作为表头
            if (table.getRows() > 0) {
                for (int i = 0; i < table.getColumns(); i++) {
                    Table.Cell cell = findTableCell(table, 0, i);
                    if (cell != null && (cell.getStyle() != null && cell.getStyle().getFillColor() != null)) {
                        hasHeader = true;
                        break;
                    }
                }
            }

            if (hasHeader) {
                tableBuilder.append("<thead><tr>\n");
                for (int i = 0; i < table.getColumns(); i++) {
                    Table.Cell cell = findTableCell(table, 0, i);
                    if (cell != null) {
                        tableBuilder.append("<th style=\"").append(styleMapper.mapTableCellStyle(cell.getStyle(), cell))
                                .append("\">").append(escapeHTML(cell.getText())).append("</th>\n");
                    } else {
                        tableBuilder.append("<th>&nbsp;</th>\n");
                    }
                }
                tableBuilder.append("</tr></thead>\n<tbody>\n");

                // 数据行从第二行开始
                for (int i = 1; i < table.getRows(); i++) {
                    tableBuilder.append("<tr>\n");
                    for (int j = 0; j < table.getColumns(); j++) {
                        Table.Cell cell = findTableCell(table, i, j);
                        if (cell != null) {
                            tableBuilder.append("<td style=\"").append(styleMapper.mapTableCellStyle(cell.getStyle(), cell))
                                    .append("\">").append(escapeHTML(cell.getText())).append("</td>\n");
                        } else {
                            tableBuilder.append("<td>&nbsp;</td>\n");
                        }
                    }
                    tableBuilder.append("</tr>\n");
                }
                tableBuilder.append("</tbody>\n");
            } else {
                tableBuilder.append("<tbody>\n");
                for (int i = 0; i < table.getRows(); i++) {
                    tableBuilder.append("<tr>\n");
                    for (int j = 0; j < table.getColumns(); j++) {
                        Table.Cell cell = findTableCell(table, i, j);
                        if (cell != null) {
                            tableBuilder.append("<td style=\"").append(styleMapper.mapTableCellStyle(cell.getStyle(), cell))
                                    .append("\">").append(escapeHTML(cell.getText())).append("</td>\n");
                        } else {
                            tableBuilder.append("<td>&nbsp;</td>\n");
                        }
                    }
                    tableBuilder.append("</tr>\n");
                }
                tableBuilder.append("</tbody>\n");
            }

            tableBuilder.append("</table>\n");
        } else {
            tableBuilder.append("<div class=\"table-placeholder\">[Table Content]</div>\n");
        }

        return tableBuilder.toString();
    }

    /**
     * 转换其他形状为HTML
     */
    private String convertGenericToHTML(Shape shape) {
        StringBuilder shapeHtml = new StringBuilder();

        String shapeClass = "generic-shape";
        if (shape.getShapeType() != null) {
            shapeClass = shape.getShapeType().toLowerCase().replace("_", "-") + "-shape";
        }

        shapeHtml.append("<div class=\"shape-container\">\n")
                .append(String.format(
                        "    <div class=\"%s\" style=\"%s\">\n",
                        shapeClass,
                        styleMapper.mapGenericStyle(shape.getStyle(), shape)
                ));

        // 形状内容
        if (shape.getName() != null && !shape.getName().isEmpty()) {
            shapeHtml.append("        <div class=\"shape-label\">")
                    .append(escapeHTML(shape.getName()))
                    .append("</div>\n");
        }

        // 特殊形状处理
        if (shape.getShapeType() != null && shape.getShapeType().contains("ARROW")) {
            shapeHtml.append("        <div class=\"arrow-head\"></div>\n");
        }

        shapeHtml.append("    </div>\n")
                .append("</div>\n");

        return shapeHtml.toString();
    }


    /**
     * 转义 HTML特殊字符
     * @param text 原始文本
     * @return 转义后的文本
     */
    private String escapeHTML(String text) {
        if (text == null) {
            return "";
        }
        
        return text.replace("&", "&amp;")
                  .replace("<", "&lt;")
                  .replace(">", "&gt;")
                  .replace("\"", "&quot;")
                  .replace("'", "&#39;")
                  .replace("\n", "<br>");
    }


    /**
     * 辅助方法：查找表格中指定位置的单元格
     */
    private Table.Cell findTableCell(Table table, int row, int column) {
        for (Table.Cell cell : table.getCells()) {
            if (cell.getRow() == row && cell.getColumn() == column) {
                return cell;
            }
        }
        return null;
    }

}