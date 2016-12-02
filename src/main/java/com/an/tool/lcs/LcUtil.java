package com.an.tool.lcs;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class LcUtil {
    /**
     * @param filePath
     * @param content
     * @param writeType
     * @param createEmptyFile
     * @throws IOException
     */
    static void writeFile(String filePath, String content, int writeType, boolean createEmptyFile) throws IOException {
        if (filePath == null || filePath.isEmpty()) {
            throw new IllegalArgumentException("Invalid filePath " + filePath);
        }
        if (content == null || content.isEmpty()) {
            if (!createEmptyFile) {
                return;
            }
        }
        String dirName = filePath.substring(0, filePath.lastIndexOf(File.separator));
        File dir = new File(dirName);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        File file = new File(filePath);
        if (file.exists()) {
            if (writeType == 0) {
                return;
            } else if (writeType == 1) {
                file.delete();
                file.createNewFile();
            }
        } else {
            file.createNewFile();
        }
        if (!content.isEmpty()) {
            try (BufferedWriter bf = new BufferedWriter(new FileWriter(filePath, writeType == 2))) {
                bf.write(content);
            }
        }
    }
}
