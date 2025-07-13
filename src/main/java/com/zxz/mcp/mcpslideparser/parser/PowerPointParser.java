package com.zxz.mcp.mcpslideparser.parser;

import com.zxz.mcp.mcpslideparser.model.Slide;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * PowerPoint文件解析器接口
 */

public interface PowerPointParser {

    /**
     * 解析 PPT文件
     * @param file PPT文件
     * @return 幻灯片列表
     * @throws IOException 文件读取异常
     * @throws IllegalArgumentException 文件格式不支持
     */
    List<Slide> parse(File file) throws IOException, IllegalArgumentException;
    
    /**
     * 检查是否支持该格式
     * @param extension 文件扩展名
     * @return 是否支持
     */
    boolean supportsFormat(String extension);
}