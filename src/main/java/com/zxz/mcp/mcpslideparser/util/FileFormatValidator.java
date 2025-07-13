package com.zxz.mcp.mcpslideparser.util;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class FileFormatValidator {
    public static void validate(File file) throws IOException, InvalidFormatException {
        // 基础检查
        if (!file.exists()) {
            throw new IOException("文件不存在: " + file.getAbsolutePath());
        }
        if (!file.isFile()) {
            throw new IOException("路径不是文件: " + file.getAbsolutePath());
        }

        // 格式检测
        byte[] header = Files.readAllBytes(file.toPath());
        if (header.length < 8) {
            throw new InvalidFormatException("文件过小或损坏");
        }

        // 检查是否是加密的OLE2文件
        if (isEncryptedOLE2(header)) {
            throw new EncryptedDocumentException("加密的PPT文件不支持");
        }

        // 检查格式签名
        if (!isPPT(header) && !isPPTX(header)) {
            throw new InvalidFormatException("非PPT/PPTX格式文件");
        }
    }

    private static boolean isPPT(byte[] header) {
        return header[0] == (byte) 0xD0 && 
               header[1] == (byte) 0xCF &&
               header[2] == 0x11 && 
               header[3] == (byte) 0xE0;
    }

    private static boolean isPPTX(byte[] header) {
        return header[0] == 0x50 && header[1] == 0x4B;
    }

    private static boolean isEncryptedOLE2(byte[] header) {
        // OLE2加密文件特征检查
        return isPPT(header) && header[8] == 0x2D && header[9] == 0x00;
    }
}