package com.zxz.mcp.mcpslideparser.parser;

import com.zxz.mcp.mcpslideparser.model.Shape;
import com.zxz.mcp.mcpslideparser.model.*;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


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

                    // 新增：检查是否为纯色背景
                    if (xslfSlide.getBackground().getFillStyle().getPaint() instanceof Color) {
                        Color bgColor = (Color) xslfSlide.getBackground().getFillStyle().getPaint();
                        backgroundStyle.setBackgroundColor(formatColorToHex(bgColor));
                    }

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
            processPictureShape((XSLFPictureShape) xslfShape, shape);
        } else if (xslfShape instanceof XSLFTable) {
            shape.setType(Shape.ShapeType.TABLE);
            processTableShape((XSLFTable) xslfShape, shape);
        } else {
            shape.setType(Shape.ShapeType.SHAPE);
            processGenericShape(xslfShape, shape);
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

    private void processPictureShape(XSLFPictureShape pictureShape, Shape shape) {
        try {
            XSLFPictureData pictureData = pictureShape.getPictureData();
            if (pictureData != null) {
                // 设置图片数据
                shape.setImageData(pictureData.getData());

                // 设置图片类型
                String contentType = pictureData.getContentType();
                shape.setImageType(contentType != null ? contentType : "image/png");

                // 设置图片名称
                if (pictureShape.getShapeName() != null) {
                    shape.setName(pictureShape.getShapeName());
                }

                // 设置图片尺寸
                Dimension dimension = pictureData.getImageDimension();
                if (dimension != null) {
                    shape.setWidth(dimension.getWidth());
                    shape.setHeight(dimension.getHeight());
                }
            }
        } catch (Exception e) {
            logger.warn("处理图片形状时出错 (ID: {})", pictureShape.getShapeId(), e);
        }
    }




    private void processTableShape(XSLFTable tableShape, Shape shape) {
        try {
            // === 1. 初始化表格对象 ===
            Table table = new Table();

            // 获取所有行（不省略任何安全检查）
            List<XSLFTableRow> rows = tableShape.getRows();
            if (rows == null || rows.isEmpty()) {
                shape.setTable(table); // 返回空表格
                return;
            }

            // 设置行数和列数（完整处理空表情况）
            table.setRows(rows.size());
            int columnCount = 0;
            for (XSLFTableRow row : rows) {
                if (row != null && row.getCells() != null) {
                    columnCount = Math.max(columnCount, row.getCells().size());
                }
            }
            table.setColumns(columnCount);

            // === 2. 遍历每个单元格（完整遍历逻辑）===
            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                XSLFTableRow row = rows.get(rowIndex);
                if (row == null) continue;

                List<XSLFTableCell> cells = row.getCells();
                if (cells == null) continue;

                for (int colIndex = 0; colIndex < cells.size(); colIndex++) {
                    XSLFTableCell cell = cells.get(colIndex);
                    if (cell == null) continue;

                    // === 2.1 创建单元格对象（完整属性设置）===
                    Table.Cell tableCell = new Table.Cell();
                    tableCell.setRow(rowIndex); // 使用实际遍历索引
                    tableCell.setColumn(colIndex);

                    // === 2.2 获取文本内容（不省略任何处理步骤）===
                    StringBuilder cellText = new StringBuilder();
                    try {
                        for (XSLFTextParagraph paragraph : cell.getTextParagraphs()) {
                            if (paragraph == null) continue;

                            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                                if (textRun != null && textRun.getRawText() != null) {
                                    cellText.append(textRun.getRawText());
                                }
                            }
                        }
                    } catch (Exception e) {
                        logger.warn("获取单元格文本失败", e);
                    }
                    tableCell.setText(cellText.toString());

                    // === 2.3 获取单元格样式（完整样式解析）===
                    Style cellStyle = new Style();
                    try {
                        // 获取底层XML对象
                        XmlObject xmlObject = cell.getXmlObject();
                        if (xmlObject instanceof CTTableCell) {
                            CTTableCell ctCell = (CTTableCell) xmlObject;

                            // 背景色处理
                            if (ctCell.isSetTcPr()) {
                                processCellBackgroundColor(ctCell, cellStyle);
                            }

                            // 边框处理（四边完整处理）
                            if (ctCell.isSetTcPr()) {
                                CTTableCellProperties pr = ctCell.getTcPr();
                                processBorder(pr.getLnL(), cellStyle); // 左边框
                                processBorder(pr.getLnR(), cellStyle); // 右边框
                                processBorder(pr.getLnT(), cellStyle); // 上边框
                                processBorder(pr.getLnB(), cellStyle); // 下边框
                            }
                        }
                    } catch (Exception e) {
                        logger.warn("解析单元格样式失败", e);
                    }

                    tableCell.setStyle(cellStyle);
                    table.addCell(tableCell);
                }
            }

            shape.setTable(table);

        } catch (Exception e) {
            logger.error("表格处理发生严重错误", e);
            shape.setTable(new Table()); // 确保始终返回有效对象
        }
    }

    private void processCellBackgroundColor(CTTableCell ctCell, Style cellStyle) {
        try {
            // 方法1：使用XPath直接提取颜色值
            XmlObject[] fillObjects = ctCell.getTcPr().selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
                            ".//a:fill/a:solidFill/a:srgbClr/@val"
            );

            if (fillObjects.length > 0) {
                String rgbStr = fillObjects[0].toString();
                if (rgbStr.length() == 6) {
                    cellStyle.setFillColor("#" + rgbStr);
                }
            }

            // 方法2：备用方案 - 通过XML字符串解析
            if (cellStyle.getFillColor() == null) {
                String xml = ctCell.getTcPr().toString();
                Matcher matcher = Pattern.compile("srgbClr val=\"([A-F0-9]{6})\"").matcher(xml);
                if (matcher.find()) {
                    cellStyle.setFillColor("#" + matcher.group(1));
                }
            }

            // 方法3：最终回退方案
            if (cellStyle.getFillColor() == null) {
                cellStyle.setFillColor("#FFFFFF"); // 默认白色背景
            }

        } catch (Exception e) {
            logger.debug("背景色解析失败，使用默认白色", e);
            cellStyle.setFillColor("#FFFFFF");
        }
    }


    // === 3. 边框处理辅助方法（完整实现）===
    private void processBorder(CTLineProperties border, Style style) {
        if (border == null) return;

        try {
            // 边框颜色
            if (border.isSetSolidFill() && border.getSolidFill().isSetSrgbClr()) {
                byte[] rgb = border.getSolidFill().getSrgbClr().getVal();
                String color = String.format(
                        "#%02X%02X%02X",
                        rgb[0] & 0xFF,
                        rgb[1] & 0xFF,
                        rgb[2] & 0xFF
                );
                style.setBorderColor(color);
            }

            // 边框宽度（精确单位转换）
            if (border.isSetW()) {
                int width = (int)(border.getW() / 12700); // EMU转像素
                style.setBorderWidth(width);
            }
        } catch (Exception e) {
            logger.debug("边框处理异常", e);
        }
    }


    private void processGenericShape(XSLFShape genericShape, Shape shape) {
        try {
            // 设置形状名称
            if (genericShape.getShapeName() != null) {
                shape.setName(genericShape.getShapeName());
            }

            // 对于AutoShape可以获取具体形状类型
            if (genericShape instanceof XSLFAutoShape) {
                XSLFAutoShape autoShape = (XSLFAutoShape) genericShape;
                shape.setShapeType(autoShape.getShapeType().name());
            }

            // 设置形状样式
            if (genericShape instanceof XSLFSimpleShape) {
                XSLFSimpleShape simpleShape = (XSLFSimpleShape) genericShape;
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


    private static String formatColorToHex(Color color) {
        return String.format("#%06X", 0xFFFFFF & color.getRGB());
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