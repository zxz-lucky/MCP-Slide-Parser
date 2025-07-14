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
                  .append("    <title>PPT to HTML Conversion</title>\n")
                  .append("    <style>\n")
                  .append("        .slide { margin-bottom: 50px; border: 1px solid #ddd; padding: 20px; position: relative; min-height: 500px; }\n")
                  .append("        .slide-title { font-size: 24px; margin-bottom: 20px; }\n")
                  .append("        .shape { position: absolute; }\n")
                  .append("        .text-run { white-space: pre-wrap; }\n")
                  .append("        .slide-notes { margin-top: 20px; padding: 10px; background-color: #f5f5f5; }\n")
                  .append(styleMapper.getGlobalCSS())
                  .append("    </style>\n")
                  .append("</head>\n")
                  .append("<body>\n");
        
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
        
        // 幻灯片容器
        slideBuilder.append(String.format(
            "<div class=\"slide\" id=\"slide-%d\" style=\"%s\">\n",
            slide.getSlideNumber(),
            styleMapper.mapSlideStyle(slide.getBackgroundStyle())
        ));
        
        // 幻灯片标题
        slideBuilder.append(String.format(
            "    <h1 class=\"slide-title\">%s</h1>\n",
            escapeHTML(slide.getTitle())
        ));
        
        // 幻灯片形状
        for (Shape shape : slide.getShapes()) {
            slideBuilder.append(convertShapeToHTML(shape));
        }
        
        // 幻灯片备注
        if (slide.getNotes() != null && !slide.getNotes().isEmpty()) {
            slideBuilder.append("    <div class=\"slide-notes\">\n")
                      .append("        <h3>Notes:</h3>\n")
                      .append("        <p>").append(escapeHTML(slide.getNotes())).append("</p>\n")
                      .append("    </div>\n");
        }
        
        slideBuilder.append("</div>\n");
        
        return slideBuilder.toString();
    }
    
    /**
     * 将形状转换为 HTML片段
     * @param shape 形状对象
     * @return HTML字符串
     */
    private String convertShapeToHTML(Shape shape) {
        StringBuilder shapeBuilder = new StringBuilder();
        
        // 形状容器
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
        
        runBuilder.append(String.format(
            "<%s class=\"text-run\" style=\"%s\"%s>%s</%s>\n",
            tag,
            styleMapper.mapTextStyle(textRun.getStyle(), textRun),
            textRun.getHyperlink() != null ? " href=\"" + escapeHTML(textRun.getHyperlink()) + "\"" : "",
            escapeHTML(textRun.getText()),
            tag
        ));
        
        return runBuilder.toString();
    }



    /**
     * 转换图片为HTML
     */
    private String convertImageToHTML(Shape shape) {
        StringBuilder imageBuilder = new StringBuilder();

        if (shape.getImageData() != null) {
            // 将图片数据转换为Base64编码
            String base64Image = Base64.getEncoder().encodeToString(shape.getImageData());
            String imageType = shape.getImageType() != null ?
                    shape.getImageType().replace("image/", "") : "png";

            imageBuilder.append(String.format(
                    "<img src=\"data:image/%s;base64,%s\" alt=\"%s\" style=\"%s\" />\n",
                    imageType,
                    base64Image,
                    escapeHTML(shape.getAltText() != null ? shape.getAltText() : ""),
                    styleMapper.mapImageStyle(shape.getStyle(), shape)
            ));
        } else {
            imageBuilder.append(String.format(
                    "<div class=\"image-placeholder\" style=\"%s\">[Image: %s]</div>\n",
                    styleMapper.mapImageStyle(shape.getStyle(), shape),
                    escapeHTML(shape.getName() != null ? shape.getName() : "")
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
            tableBuilder.append(String.format(
                    "<table class=\"table-container\" style=\"%s\">\n",
                    styleMapper.mapTableStyle(shape.getStyle(), shape)
            ));

            // 创建表格行
            for (int i = 0; i < table.getRows(); i++) {
                tableBuilder.append("<tr>\n");

                // 创建表格单元格
                for (int j = 0; j < table.getColumns(); j++) {
                    Table.Cell cell = findTableCell(table, i, j);
                    if (cell != null) {
                        tableBuilder.append(String.format(
                                "<td class=\"table-cell\" style=\"%s\">%s</td>\n",
                                styleMapper.mapTableCellStyle(cell.getStyle(), cell),
                                escapeHTML(cell.getText() != null ? cell.getText() : "")
                        ));
                    } else {
                        // 空单元格
                        tableBuilder.append("<td class=\"table-cell\">&nbsp;</td>\n");
                    }
                }

                tableBuilder.append("</tr>\n");
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

        // 根据形状类型添加不同的HTML
        if (shape.getShapeType() != null && shape.getShapeType().contains("ARROW")) {
            // 箭头形状特殊处理
            shapeHtml.append(String.format(
                    "<div class=\"arrow-shape\" style=\"%s\"></div>\n",
                    styleMapper.mapGenericStyle(shape.getStyle(), shape)
            ));
        } else {
            // 默认其他形状
            shapeHtml.append(String.format(
                    "<div class=\"other-shape\" style=\"%s\">%s</div>\n",
                    styleMapper.mapGenericStyle(shape.getStyle(), shape),
                    escapeHTML(shape.getName() != null ? shape.getName() : "")
            ));
        }

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