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
            processTextShape((HSLFTextShape) hslfShape, shape);
        } else if (hslfShape instanceof HSLFPictureShape) {
            shape.setType(Shape.ShapeType.IMAGE);
            // 可添加图片处理逻辑
        } else if (hslfShape instanceof HSLFTable) {
            shape.setType(Shape.ShapeType.TABLE);
            // 可添加表格处理逻辑
        } else {
            shape.setType(Shape.ShapeType.SHAPE);
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