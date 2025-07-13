package com.zxz.mcp.mcpslideparser.util;

import java.io.File;

/**
 * 文件操作工具类
 */

public class FileUtils {

    /**
     * 获取文件扩展名
     * @param file 文件对象
     * @return 扩展名(不带点)
     */
    public static String getFileExtension(File file) {
        String fileName = file.getName();
        int dotIndex = fileName.lastIndexOf('.');
        return (dotIndex == -1) ? "" : fileName.substring(dotIndex + 1);
    }
    
    /**
     * 检查是否是支持的 PPT文件
     * @param file 文件对象
     * @return 是否是PPT文件
     */
    public static boolean isPowerPointFile(File file) {
        String extension = getFileExtension(file);
        return extension.equalsIgnoreCase("ppt") ||
               extension.equalsIgnoreCase("pptx") ||
               extension.equalsIgnoreCase("pot") ||
               extension.equalsIgnoreCase("potx") ||
               extension.equalsIgnoreCase("pps") ||
               extension.equalsIgnoreCase("ppsx") ||
               extension.equalsIgnoreCase("dps") ||
               extension.equalsIgnoreCase("dpt") ||
               extension.equalsIgnoreCase("pptm") ||
               extension.equalsIgnoreCase("potm") ||
               extension.equalsIgnoreCase("ppsm");
    }
}