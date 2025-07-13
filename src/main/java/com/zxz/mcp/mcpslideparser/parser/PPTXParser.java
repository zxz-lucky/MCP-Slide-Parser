package com.zxz.mcp.mcpslideparser.parser;

import com.zxz.mcp.mcpslideparser.model.Shape;
import com.zxz.mcp.mcpslideparser.model.Slide;
import com.zxz.mcp.mcpslideparser.model.Style;
import com.zxz.mcp.mcpslideparser.model.TextRun;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.xslf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


/**
 * PPTX文件解析器
 */

public class PPTXParser implements PowerPointParser {

    private static final Logger logger = LoggerFactory.getLogger(PPTXParser.class);

    @Override
    public List<Slide> parse(File file) throws IOException, IllegalArgumentException {
        List<Slide> slides = new ArrayList<>();

        try (OPCPackage pkg = OPCPackage.open(file);
             XMLSlideShow pptx = new XMLSlideShow(pkg)) {
            logger.info("开始解析 PPTX 文件: {}", file.getName());

            for (XSLFSlide xslfSlide : pptx.getSlides()) {
                Slide slide = new Slide();
                slide.setSlideNumber(xslfSlide.getSlideNumber());

                // 处理幻灯片标题
                String title = xslfSlide.getTitle();
                slide.setTitle(title != null ? title : "Slide " + slide.getSlideNumber());
                logger.debug("解析幻灯片 {}: {}", slide.getSlideNumber(), slide.getTitle());

                // 处理幻灯片背景
                Style backgroundStyle = new Style();
                XSLFBackground background = xslfSlide.getBackground();
                if (background != null && background.getFillColor() != null) {
                    backgroundStyle.setBackgroundColor(String.format("#%06X", (0xFFFFFF & background.getFillColor().getRGB())));
                }
                slide.setBackgroundStyle(backgroundStyle);

                // 处理幻灯片备注
                XSLFNotes notes = xslfSlide.getNotes();
                if (notes != null) {
                    StringBuilder notesText = new StringBuilder();
                    for (List<XSLFTextParagraph> paragraphList : notes.getTextParagraphs()) {
                        for (XSLFTextParagraph paragraph : paragraphList) {
                            for (XSLFTextRun run : paragraph.getTextRuns()) {
                                notesText.append(run.getRawText());
                            }
                            notesText.append("\n");
                        }
                    }
                    slide.setNotes(notesText.toString());
                }

                // 处理幻灯片中的形状
                for (XSLFShape shape : xslfSlide.getShapes()) {
                    processShape(shape, slide);
                }

                slides.add(slide);
            }

            logger.info("成功解析 {} 张幻灯片", slides.size());
        } catch (OpenXML4JException e) {
            logger.error("解析PPTX文件失败", e);
            throw new IOException("Failed to parse PPTX file", e);
        }

        return slides;
    }

    private void processShape(XSLFShape xslfShape, Slide slide) {
        Shape shape = new Shape();
        shape.setId(Integer.toString(xslfShape.getShapeId()));

        // 设置形状位置和大小
        java.awt.geom.Rectangle2D anchor = xslfShape.getAnchor();
        shape.setX(anchor.getX());
        shape.setY(anchor.getY());
        shape.setWidth(anchor.getWidth());
        shape.setHeight(anchor.getHeight());

        // 设置形状样式
        Style shapeStyle = new Style();
        if (xslfShape instanceof XSLFSimpleShape) {
            XSLFSimpleShape simpleShape = (XSLFSimpleShape) xslfShape;
            if (simpleShape.getFillColor() != null) {
                shapeStyle.setFillColor(String.format("#%06X", (0xFFFFFF & simpleShape.getFillColor().getRGB())));
            }
            if (simpleShape.getLineColor() != null) {
                shapeStyle.setBorderColor(String.format("#%06X", (0xFFFFFF & simpleShape.getLineColor().getRGB())));
                shapeStyle.setBorderWidth((int) simpleShape.getLineWidth());
            }
        }
        shape.setStyle(shapeStyle);

        // 处理不同类型的形状
        if (xslfShape instanceof XSLFTextShape) {
            shape.setType(Shape.ShapeType.TEXT_BOX);
            processTextShape((XSLFTextShape) xslfShape, shape);
        } else if (xslfShape instanceof XSLFPictureShape) {
            shape.setType(Shape.ShapeType.IMAGE);
        } else if (xslfShape instanceof XSLFTable) {
            shape.setType(Shape.ShapeType.TABLE);
        } else {
            shape.setType(Shape.ShapeType.SHAPE);
        }

        slide.addShape(shape);
    }

    private void processTextShape(XSLFTextShape textShape, Shape shape) {
        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                TextRun textRun = new TextRun();
                textRun.setText(run.getRawText());

                Style textStyle = new Style();
                textStyle.setFontFamily(run.getFontFamily());

                // 修正1: 字体大小处理
                double fontSize = run.getFontSize();
                textStyle.setFontSize(fontSize > 0 ? (int)Math.round(fontSize) : 12); // 默认12pt

                // 修正2: 字体颜色处理
                PaintStyle paintStyle = run.getFontColor();
                if (paintStyle != null) {
                    if (paintStyle instanceof PaintStyle.SolidPaint) {
                        Color color = ((PaintStyle.SolidPaint)paintStyle).getSolidColor().getColor();
                        if (color != null) {
                            textStyle.setFontColor(String.format("#%06X", 0xFFFFFF & color.getRGB()));
                        }
                    }
                }

                textRun.setStyle(textStyle);
                textRun.setBold(run.isBold());
                textRun.setItalic(run.isItalic());
                textRun.setUnderlined(run.isUnderlined());

                if (run.getHyperlink() != null) {
                    textRun.setHyperlink(run.getHyperlink().getAddress());
                }

                shape.addTextRun(textRun);
            }
        }
    }

    @Override
    public boolean supportsFormat(String extension) {
        return extension != null &&
                (extension.equalsIgnoreCase("pptx") ||
                        extension.equalsIgnoreCase("pptm") ||
                        extension.equalsIgnoreCase("potx") ||
                        extension.equalsIgnoreCase("potm") ||
                        extension.equalsIgnoreCase("ppsx") ||
                        extension.equalsIgnoreCase("ppsm"));
    }
}