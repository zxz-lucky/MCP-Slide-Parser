package com.zxz.mcp.mcpslideparser.parser;

import com.zxz.mcp.mcpslideparser.model.*;
import com.zxz.mcp.mcpslideparser.model.Shape;
import org.apache.poi.hslf.usermodel.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PPTParser implements PowerPointParser {
    private static final Logger logger = LoggerFactory.getLogger(PPTParser.class);

    @Override
    public List<Slide> parse(File file) throws IOException, IllegalArgumentException {
        List<Slide> slides = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(file);
             HSLFSlideShow ppt = new HSLFSlideShow(fis)) {

            logger.info("开始解析PPT文件: {}", file.getName());

            for (HSLFSlide hslfSlide : ppt.getSlides()) {
                Slide slide = new Slide();
                slide.setSlideNumber(hslfSlide.getSlideNumber() + 1); // 调整为1-based编号

                // 处理幻灯片标题
                String title = hslfSlide.getTitle();
                slide.setTitle(title != null ? title : "Slide " + slide.getSlideNumber());
                logger.debug("解析幻灯片 {}: {}", slide.getSlideNumber(), slide.getTitle());

                // 处理幻灯片背景
                Style backgroundStyle = new Style();
                HSLFBackground background = hslfSlide.getBackground();
                if (background != null && background.getFill() != null) {
                    Color color = background.getFill().getForegroundColor();
                    if (color != null) {
                        backgroundStyle.setBackgroundColor(formatColorToHex(color));
                    }
                }
                slide.setBackgroundStyle(backgroundStyle);

                // 处理幻灯片备注
                HSLFNotes notes = hslfSlide.getNotes();
                if (notes != null) {
                    slide.setNotes(extractNotesText(notes));
                }

                // 处理幻灯片中的形状
                for (HSLFShape shape : hslfSlide.getShapes()) {
                    try {
                        processShape(shape, slide);
                    } catch (Exception e) {
                        logger.warn("处理形状时出错 (ID: {})", shape.getShapeId(), e);
                    }
                }

                slides.add(slide);
            }

            logger.info("成功解析 {} 张幻灯片", slides.size());
        } catch (Exception e) {
            throw new IOException("PPT文件解析失败: " + e.getMessage(), e);
        }

        return slides;
    }

    private String extractNotesText(HSLFNotes notes) {
        StringBuilder notesText = new StringBuilder();
        for (List<HSLFTextParagraph> paragraphList : notes.getTextParagraphs()) {
            for (HSLFTextParagraph paragraph : paragraphList) {
                for (HSLFTextRun run : paragraph.getTextRuns()) {
                    notesText.append(run.getRawText());
                }
                notesText.append("\n");
            }
        }
        return notesText.toString();
    }

    private void processShape(HSLFShape hslfShape, Slide slide) {
        Shape shape = new Shape();
        shape.setId(Integer.toString(hslfShape.getShapeId()));

        // 设置形状位置和大小
        java.awt.geom.Rectangle2D anchor = hslfShape.getAnchor();
        shape.setX(anchor.getX());
        shape.setY(anchor.getY());
        shape.setWidth(anchor.getWidth());
        shape.setHeight(anchor.getHeight());

        // 设置形状样式
        if (hslfShape instanceof HSLFSimpleShape) {
            HSLFSimpleShape simpleShape = (HSLFSimpleShape) hslfShape;
            Style shapeStyle = new Style();

            if (simpleShape.getFillColor() != null) {
                shapeStyle.setFillColor(formatColorToHex(simpleShape.getFillColor()));
            }

            if (simpleShape.getLineColor() != null) {
                shapeStyle.setBorderColor(formatColorToHex(simpleShape.getLineColor()));
                shapeStyle.setBorderWidth((int) simpleShape.getLineWidth());
            }

            shape.setStyle(shapeStyle);
        }

        // 处理不同类型的形状
        if (hslfShape instanceof HSLFTextShape) {
            shape.setType(Shape.ShapeType.TEXT_BOX);
            // 可添加文本处理逻辑
            processTextShape((HSLFTextShape) hslfShape, shape);
        } else if (hslfShape instanceof HSLFPictureShape) {
            shape.setType(Shape.ShapeType.IMAGE);
            // 可添加图片处理逻辑
            processPictureShape((HSLFPictureShape) hslfShape, shape);
        } else if (hslfShape instanceof HSLFTable) {
            shape.setType(Shape.ShapeType.TABLE);
            // 可添加表格处理逻辑
            processTableShape((HSLFTable) hslfShape, shape);
        } else {
            shape.setType(Shape.ShapeType.SHAPE);
            // 可添加其他类型处理逻辑
            processGenericShape(hslfShape, shape);
        }

        slide.addShape(shape);
    }

    private void processTextShape(HSLFTextShape textShape, Shape shape) {
        try {
            // 1. 获取文本段落 - 修正后的获取方式
            List<HSLFTextParagraph> paragraphs = textShape.getTextParagraphs();

            if (paragraphs == null || paragraphs.isEmpty()) {
                return;
            }

            for (HSLFTextParagraph paragraph : paragraphs) {
                if (paragraph == null) {
                    continue;
                }

                // 2. 获取文本运行
                List<HSLFTextRun> runs = paragraph.getTextRuns();
                if (runs == null || runs.isEmpty()) {
                    continue;
                }

                for (HSLFTextRun run : runs) {
                    if (run == null) {
                        continue;
                    }

                    TextRun textRun = new TextRun();
                    textRun.setText(run.getRawText() != null ? run.getRawText() : "");

                    Style textStyle = new Style();

                    // 3. 字体名称（安全获取）
                    textStyle.setFontFamily(run.getFontFamily() != null ? run.getFontFamily() : "Arial");

                    // 4. 字体大小处理（安全获取）
                    int fontSize = 12; // 默认值
                    try {
                        fontSize = run.getLength(); //getSize()
                        if (fontSize <= 0) fontSize = 12;
                    } catch (Exception e) {
                        logger.warn("获取字体大小失败，使用默认值", e);
                    }
                    textStyle.setFontSize(fontSize);

                    // 5. 字体颜色处理（安全获取）
                    String fontColorHex = "#000000"; // 默认黑色
                    try {
                        java.awt.Color color = (Color) run.getFontColor();
                        if (color != null) {
                            fontColorHex = String.format("#%06X", 0xFFFFFF & color.getRGB());
                        }
                    } catch (Exception e) {
                        logger.warn("获取字体颜色失败，使用默认值", e);
                    }
                    textStyle.setFontColor(fontColorHex);

                    textRun.setStyle(textStyle);
                    textRun.setBold(run.isBold());
                    textRun.setItalic(run.isItalic());
                    textRun.setUnderlined(run.isUnderlined());

                    // 6. 超链接处理
                    try {
                        HSLFHyperlink link = run.getHyperlink();
                        if (link != null && link.getAddress() != null) {
                            textRun.setHyperlink(link.getAddress());
                        }
                    } catch (Exception e) {
                        logger.warn("获取超链接失败", e);
                    }

                    shape.addTextRun(textRun);
                }
            }
        } catch (Exception e) {
            logger.error("处理文本形状时发生错误", e);
        }
    }


    private void processPictureShape(HSLFPictureShape pictureShape, Shape shape) {
        try {
            // 获取图片数据
            HSLFPictureData pictureData = pictureShape.getPictureData();
            if (pictureData != null) {
                // 设置图片数据
                shape.setImageData(pictureData.getData());

                // 设置图片类型
                String contentType = pictureData.getContentType();
                shape.setImageType(contentType != null ? contentType : "image/png");

                // 设置图片尺寸
                Dimension dimension = pictureShape.getPictureData().getImageDimension();
                if (dimension != null) {
                    shape.setWidth(dimension.getWidth());
                    shape.setHeight(dimension.getHeight());
                }

                // 设置图片名称（如果有）
                if (pictureShape.getShapeName() != null) {
                    shape.setName(pictureShape.getShapeName());
                }
            }
        } catch (Exception e) {
            logger.warn("处理图片形状时出错 (ID: {})", pictureShape.getShapeId(), e);
        }
    }

    private void processTableShape(HSLFTable tableShape, Shape shape) {
        try {
            // 创建表格对象
            Table table = new Table();
            table.setRows(tableShape.getNumberOfRows());
            table.setColumns(tableShape.getNumberOfColumns());

            // 创建表格默认样式
            Style tableStyle = new Style();

            // 方法1：尝试通过表格中的第一个单元格获取样式
            if (tableShape.getNumberOfRows() > 0 && tableShape.getNumberOfColumns() > 0) {
                HSLFTableCell firstCell = tableShape.getCell(0, 0);
                if (firstCell != null) {
                    // 获取填充颜色
                    if (firstCell.getFill() != null && firstCell.getFill().getForegroundColor() != null) {
                        tableStyle.setFillColor(formatColorToHex(firstCell.getFill().getForegroundColor()));
                    }

                    // 获取边框颜色和宽度
                    if (firstCell.getLineColor() != null) {
                        tableStyle.setBorderColor(formatColorToHex(firstCell.getLineColor()));
                        tableStyle.setBorderWidth(1); // HSLF中边框宽度通常为1
                    }
                }
            }

            // 方法2：如果第一种方法不可行，使用默认样式
            if (tableStyle.getFillColor() == null) {
                tableStyle.setFillColor("#FFFFFF"); // 默认白色填充
            }
            if (tableStyle.getBorderColor() == null) {
                tableStyle.setBorderColor("#000000"); // 默认黑色边框
                tableStyle.setBorderWidth(1); // 默认1px边框
            }

            shape.setStyle(tableStyle);

            // 处理表格单元格
            for (int i = 0; i < tableShape.getNumberOfRows(); i++) {
                for (int j = 0; j < tableShape.getNumberOfColumns(); j++) {
                    HSLFTableCell cell = tableShape.getCell(i, j);
                    if (cell != null) {
                        Table.Cell tableCell = new Table.Cell();
                        tableCell.setRow(i);
                        tableCell.setColumn(j);

                        // 设置单元格文本内容
                        StringBuilder cellText = new StringBuilder();
                        for (HSLFTextParagraph paragraph : cell.getTextParagraphs()) {
                            for (HSLFTextRun run : paragraph.getTextRuns()) {
                                cellText.append(run.getRawText());
                            }
                        }
                        tableCell.setText(cellText.toString());

                        // 设置单元格样式
                        Style cellStyle = new Style();
                        if (cell.getFill() != null && cell.getFill().getForegroundColor() != null) {
                            cellStyle.setFillColor(formatColorToHex(cell.getFill().getForegroundColor()));
                        } else {
                            cellStyle.setFillColor(tableStyle.getFillColor()); // 继承表格样式
                        }

                        if (cell.getLineColor() != null) {
                            cellStyle.setBorderColor(formatColorToHex(cell.getLineColor()));
                            cellStyle.setBorderWidth(1);
                        } else {
                            cellStyle.setBorderColor(tableStyle.getBorderColor());
                            cellStyle.setBorderWidth(tableStyle.getBorderWidth());
                        }

                        tableCell.setStyle(cellStyle);
                        table.addCell(tableCell);
                    }
                }
            }

            shape.setTable(table);
        } catch (Exception e) {
            logger.warn("处理表格形状时出错 (ID: {})", tableShape.getShapeId(), e);
        }
    }

    private void processGenericShape(HSLFShape genericShape, Shape shape) {
        try {
            // 设置形状名称
            if (genericShape.getShapeName() != null) {
                shape.setName(genericShape.getShapeName());
            }

            // 对于AutoShape可以获取具体形状类型
            if (genericShape instanceof HSLFAutoShape) {
                HSLFAutoShape autoShape = (HSLFAutoShape) genericShape;
                shape.setShapeType(autoShape.getShapeType().name());
            }

            // 设置形状样式
            if (genericShape instanceof HSLFSimpleShape) {
                HSLFSimpleShape simpleShape = (HSLFSimpleShape) genericShape;
                Style shapeStyle = new Style();

                if (simpleShape.getFillColor() != null) {
                    shapeStyle.setFillColor(formatColorToHex(simpleShape.getFillColor()));
                }

                if (simpleShape.getLineColor() != null) {
                    shapeStyle.setBorderColor(formatColorToHex(simpleShape.getLineColor()));
                    shapeStyle.setBorderWidth((int) simpleShape.getLineWidth());
                }

                shape.setStyle(shapeStyle);
            }
        } catch (Exception e) {
            logger.warn("处理通用形状时出错 (ID: {})", genericShape.getShapeId(), e);
        }
    }



    private TextRun createTextRunFromHSLF(HSLFTextRun hslfRun) {
        TextRun textRun = new TextRun();
        textRun.setText(hslfRun.getRawText() != null ? hslfRun.getRawText() : "");

        Style textStyle = new Style();
        textStyle.setFontFamily(hslfRun.getFontFamily() != null ?
                hslfRun.getFontFamily() : "Arial");

        // 字体大小安全获取
        try {
            textStyle.setFontSize(hslfRun.getLength());
        } catch (Exception e) {
            logger.warn("字体大小获取失败，使用默认值12", e);
            textStyle.setFontSize(12);
        }

        // 字体颜色安全获取
        try {
            Color color = (Color) hslfRun.getFontColor();
            textStyle.setFontColor(color != null ?
                    formatColorToHex(color) : "#000000");
        } catch (Exception e) {
            logger.warn("字体颜色获取失败，使用默认黑色", e);
            textStyle.setFontColor("#000000");
        }

        textRun.setStyle(textStyle);
        textRun.setBold(hslfRun.isBold());
        textRun.setItalic(hslfRun.isItalic());
        textRun.setUnderlined(hslfRun.isUnderlined());

        // 超链接安全获取
        try {
            HSLFHyperlink link = hslfRun.getHyperlink();
            if (link != null && link.getAddress() != null) {
                textRun.setHyperlink(link.getAddress());
            }
        } catch (Exception e) {
            logger.warn("超链接获取失败", e);
        }

        return textRun;
    }

    private static String formatColorToHex(Color color) {
        return String.format("#%02X%02X%02X",
                color.getRed(), color.getGreen(), color.getBlue());
    }

    @Override
    public boolean supportsFormat(String extension) {
        return extension != null &&
                extension.matches("(?i)ppt|pot|pps|dps|dpt");
    }
}