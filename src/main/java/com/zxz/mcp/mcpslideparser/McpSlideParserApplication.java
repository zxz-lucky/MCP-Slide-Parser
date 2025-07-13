package com.zxz.mcp.mcpslideparser;

import com.zxz.mcp.mcpslideparser.converter.HTMLConverter;
import com.zxz.mcp.mcpslideparser.model.Slide;
import com.zxz.mcp.mcpslideparser.parser.PPTParser;
import com.zxz.mcp.mcpslideparser.parser.PPTXParser;
import com.zxz.mcp.mcpslideparser.parser.PowerPointParser;
import com.zxz.mcp.mcpslideparser.util.FileUtils;
import com.zxz.mcp.mcpslideparser.util.ValidationUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

@SpringBootApplication
public class McpSlideParserApplication {

    //日志记录器
    private static final Logger logger = LoggerFactory.getLogger(McpSlideParserApplication.class);

    public static void main(String[] args) {
        if (args.length < 1) {  //若小于 1 , 说明未传入输入文件路径,打印使用说明和示例后退出程序
            System.err.println("Usage: java -jar ppt-to-html-converter.jar <input-file> [output-file]");
            System.err.println("Example: java -jar ppt-to-html-converter.jar presentation.pptx output.html");
            System.exit(1);
        }

        String inputPath = args[0];     //将第一个参数作为输入文件路径 inputPath
        // 若有第二个参数则作为输出文件路径 outputPath ，否则默认在输入文件路径后加 .html 作为输出路径。
        String outputPath = args.length > 1 ? args[1] : inputPath + ".html";

        try {
            // 1. 验证输入文件
            File inputFile = new File(inputPath);
            ValidationUtils.validatePowerPointFile(inputFile);  //验证 PPT文件是否有效
            logger.info("Processing file: {}", inputFile.getAbsolutePath());    //验证通过后记录日志说明正在处理的文件路径

            // 2. 选择合适的解析器
            PowerPointParser parser = getParserForFile(inputFile);  //根据文件扩展名获取合适的解析器
            logger.info("Using parser: {}", parser.getClass().getSimpleName()); //记录使用的解析器类名

            // 3. 解析PPT文件
            List<Slide> slides = parser.parse(inputFile);   //解析 PPT文件
            logger.info("Successfully parsed {} slides", slides.size());    //记录成功解析的幻灯片数量

            // 4. 转换为HTML
            HTMLConverter converter = new HTMLConverter();
            String html = converter.convertToHTML(slides);  //将幻灯片列表转换为完整 HTML文档

            // 5. 保存HTML文件
            File outputFile = new File(outputPath);
            try (FileWriter writer = new FileWriter(outputFile)) {
                writer.write(html);
                logger.info("HTML output saved to: {}", outputFile.getAbsolutePath());  //记录日志说明保存路径
                System.out.println("Conversion successful! Output file: " + outputFile.getAbsolutePath());
            }
        } catch (IllegalArgumentException e) {
            logger.error("Invalid input: {}", e.getMessage());
            System.err.println("Error: " + e.getMessage());
            System.exit(2);
        } catch (IOException e) {
            logger.error("Error processing file: {}", e.getMessage(), e);
            System.err.println("Error processing file: " + e.getMessage());
            System.exit(3);
        }
    }

    /**
     * 根据文件扩展名获取合适的解析器
     * @param file PPT文件
     * @return 解析器实例
     */
    private static PowerPointParser getParserForFile(File file) {

        String extension = FileUtils.getFileExtension(file);    //得到文件扩展名 extension

        if (extension.equalsIgnoreCase("pptx") ||
                extension.equalsIgnoreCase("pptm") ||
                extension.equalsIgnoreCase("potx") ||
                extension.equalsIgnoreCase("potm") ||
                extension.equalsIgnoreCase("ppsx") ||
                extension.equalsIgnoreCase("ppsm")) {
            return new PPTXParser();
        } else {
            return new PPTParser();
        }
    }

}
