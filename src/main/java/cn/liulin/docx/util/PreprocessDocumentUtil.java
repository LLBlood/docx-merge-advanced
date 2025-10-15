package cn.liulin.docx.util;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/15 15:21
 */
public class PreprocessDocumentUtil {


    /**
     * 预处理文档列表，对列表中的每个文档执行预处理操作
     * 该方法会遍历文档路径列表，对每个文档调用预处理方法，
     * 生成处理后的文档路径列表并返回
     *
     * @param docPathList 包含待处理文档路径的列表
     * @return 包含处理后文档路径的列表
     * @throws Exception 如果在预处理过程中发生错误
     */
    public static List<String> preprocessDocumentList(List<String> docPathList) throws Exception {
        List<String> processedDocPathList = new ArrayList<>();
        // 遍历文档路径列表，对每个文档执行预处理
        for (String docPath : docPathList) {
            String outPath = preprocessDocument(docPath);
            processedDocPathList.add(outPath);
        }
        return processedDocPathList;
    }

    /**
     * 在docx4j加载前预处理文档，替换不兼容的标签
     *
     * @param docPath 原始文档路径
     * @return 处理后的文档路径
     * @throws Exception 处理异常
     */
    private static String preprocessDocument(String docPath) throws Exception {
        Path originalDoc = Paths.get(docPath);
        Path processedDoc = Files.createTempFile("processed_", ".docx");

        // 复制原始文档到临时文件
        Files.copy(originalDoc, processedDoc, StandardCopyOption.REPLACE_EXISTING);

        // 创建一个新的临时文件用于输出
        Path outputDoc = Files.createTempFile("output_", ".docx");

        try (ZipFile zipFile = new ZipFile(processedDoc.toFile());
             ZipOutputStream zipOutputStream = new ZipOutputStream(Files.newOutputStream(outputDoc.toFile().toPath()))) {

            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                ZipEntry entry = entries.nextElement();
                zipOutputStream.putNextEntry(new ZipEntry(entry.getName()));

                if ("word/document.xml".equals(entry.getName())) {
                    // 处理document.xml内容
                    try (InputStream inputStream = zipFile.getInputStream(entry);
                         ByteArrayOutputStream buffer = new ByteArrayOutputStream()) {

                        int nRead;
                        byte[] data = new byte[1024];
                        while ((nRead = inputStream.read(data, 0, data.length)) != -1) {
                            buffer.write(data, 0, nRead);
                        }
                        buffer.flush();

                        String xmlContent = new String(buffer.toByteArray(), StandardCharsets.UTF_8);

                        // 处理不兼容的标签，将w:start和w:end替换为w:left和w:right
                        xmlContent = xmlContent.replaceAll("<w:start\\b", "<w:left");
                        xmlContent = xmlContent.replaceAll("</w:start>", "</w:left");
                        xmlContent = xmlContent.replaceAll("<w:end\\b", "<w:right");
                        xmlContent = xmlContent.replaceAll("</w:end>", "</w:right>");

                        // 移除页眉页脚引用
                        xmlContent = xmlContent.replaceAll("<w:headerReference[^>]*/>", "");
                        xmlContent = xmlContent.replaceAll("<w:footerReference[^>]*/>", "");

                        // 写入处理后的内容
                        zipOutputStream.write(xmlContent.getBytes(StandardCharsets.UTF_8));
                    }
                } else {
                    // 直接复制其他文件
                    try (InputStream inputStream = zipFile.getInputStream(entry)) {
                        byte[] buffer = new byte[1024];
                        int length;
                        while ((length = inputStream.read(buffer)) > 0) {
                            zipOutputStream.write(buffer, 0, length);
                        }
                    }
                }

                zipOutputStream.closeEntry();
            }
        }

        // 删除中间文件
        Files.deleteIfExists(processedDoc);

        return outputDoc.toString();
    }
}