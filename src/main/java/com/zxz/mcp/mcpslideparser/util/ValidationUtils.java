package com.zxz.mcp.mcpslideparser.util;

import java.io.File;

/**
 * 验证工具类
 */

public class ValidationUtils {

    /**
     * 验证 PPT文件是否有效
     * @param file 文件对象
     * @throws IllegalArgumentException 如果文件无效
     */
    public static void validatePowerPointFile(File file) throws IllegalArgumentException {
        if (file == null) {
            throw new IllegalArgumentException("File cannot be null");
        }

        if (!file.exists()) {
            throw new IllegalArgumentException("File does not exist: " + file.getPath());
        }

        if (!file.isFile()) {
            throw new IllegalArgumentException("Path is not a file: " + file.getPath());
        }

        if (!FileUtils.isPowerPointFile(file)) {
            throw new IllegalArgumentException("Unsupported file format: " + FileUtils.getFileExtension(file));
        }
    }
}