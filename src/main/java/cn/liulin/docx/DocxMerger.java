package cn.liulin.docx;

import cn.liulin.docx.util.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class DocxMerger {
    private static final Logger logger = LoggerUtil.getLogger(DocxMerger.class);
    
    // 存储所有图片引用路径 ((原始路径+文档索引) -> 新名称)
    private Map<String, String> imageReferences = new HashMap<>();
    
    // 图片计数器
    private int imageCounter = 10;

    /**
     * 合并传入的多个文档
     *
     * @param docPathList 传入文档路径list
     * @param outputPath 输出文档路径
     * @author liulin
     * @date 2025/10/15 15:13
     */
    public void mergeList(List<String> docPathList, String outputPath) throws Exception {
        LoggerUtil.logMethodEntry(logger, "mergeList", docPathList, outputPath);
        
        try {
            logger.info("开始合并文档...");

            // 创建输出目录
            File output = new File(outputPath);
            if (!output.getParentFile().exists()) {
                output.getParentFile().mkdirs();
            }

            // 分批处理文档，每批处理5个文档（减少内存使用）
            WordprocessingMLPackage resultDoc = null;
            List<String> batch = new ArrayList<>();
            
            for (int i = 0; i < docPathList.size(); i++) {
                batch.add(docPathList.get(i));
                
                // 每5个文档处理一次，或者到达最后一个文档时处理（减少批处理大小）
                if (batch.size() >= 5 || i == docPathList.size() - 1) {
                    if (resultDoc == null) {
                        // 第一批文档，创建基础文档
                        resultDoc = mergeBatch(batch, null);
                    } else {
                        // 后续批次，将结果合并到已有文档中
                        resultDoc = mergeBatch(batch, resultDoc);
                    }
                    batch.clear();
                    
                    // 更积极地建议垃圾回收
                    System.gc();
                    
                    // 提示当前进度
                    logger.info("已处理 {}/{} 个文档", Math.min(i + 1, docPathList.size()), docPathList.size());
                }
            }

            // 保存最终文档
            resultDoc.save(output);
            
            // 在保存后复制图片到最终文档并更新相关XML文件
            copyImagesToFinalDocument(docPathList, outputPath);
            
            logger.info("文档已成功合并并保存到: {}", outputPath);
            resultDoc.reset();
            
            LoggerUtil.logMethodExit(logger, "mergeList", "合并完成");
        } catch (Exception e) {
            LoggerUtil.logMethodException(logger, "mergeList", e);
            throw e;
        }
    }
    
    /**
     * 分批合并文档
     * @param batchDocPaths 当前批次的文档路径
     * @param baseDoc 已有的基础文档，如果为null则创建新的
     * @return 合并后的文档
     */
    private WordprocessingMLPackage mergeBatch(List<String> batchDocPaths, WordprocessingMLPackage baseDoc) throws Exception {
        logger.info("开始处理批次，包含 {} 个文档", batchDocPaths.size());
        
        // 预处理原始文档
        List<String> processedDocPathList = PreprocessDocumentUtil.preprocessDocumentList(batchDocPaths);

        // 加载当前批次数据（跳过图片加载）
        List<WordprocessingMLPackage> docList = WordProcessingUtils.loadDocListSkipImages(processedDocPathList);

        WordprocessingMLPackage resultDoc;
        if (baseDoc == null) {
            // 如果没有基础文档，使用第一个文档作为基础
            resultDoc = docList.get(0);
        } else {
            // 如果已有基础文档，将其作为第一个文档
            docList.add(0, baseDoc);
            resultDoc = baseDoc;
        }

        // 合并样式，以第一个文档的样式为基准
        StyleReMapperUtil.mergeStyles(docList);

        // 映射编号（避免列表编号混乱）
        NumberingMapperUtil.mapNumbering(docList);

        // 保存图片引用路径，不直接复制图片
        imageCounter = ResourceCopierUtil.saveImageReferences(docList, imageReferences, imageCounter);

        // 保存两个文档的格式信息（暂时保留但不处理表格边框）
        Map<String, String> formatProperties = TableFormatPreserverUtil.saveDocumentFormat(docList);

        // 在合并前应用默认字体大小
        FontMapperUtil.applyDocListDefaultFontSizesBeforeMerge(docList, formatProperties);

        // 获取基础文档的主要部分
        MainDocumentPart main1 = resultDoc.getMainDocumentPart();

        // 移除文档网格设置
        WordProcessingUtils.removeDocumentGridSettingsList(docList);

        // 将合并doc 的所有内容追加到 doc1
        WordProcessingUtils.addDocListToBase(main1, docList);

        // 清理临时文件
        for (String s : processedDocPathList) {
            Files.deleteIfExists(Paths.get(s));
        }

        for (int i = 1; i < docList.size(); i++) {
            docList.get(i).reset();
        }
        logger.info("批次处理完成");
        return resultDoc;
    }
    
    /**
     * 将图片复制到最终的文档中
     * 
     * @param docPathList 原始文档路径列表
     * @param outputPath 输出文档路径
     * @throws Exception IO异常
     */
    private void copyImagesToFinalDocument(List<String> docPathList, String outputPath) throws Exception {
        logger.info("开始将图片复制到最终文档中...");
        
        if (imageReferences.isEmpty()) {
            logger.info("没有图片需要复制");
            return;
        }
        
        // 创建临时文件用于处理
        String tempPath = outputPath + ".tmp";
        
        // 先复制所有ZIP条目并添加图片文件
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(outputPath));
             ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(tempPath))) {
            
            // 复制所有现有条目
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);
                
                byte[] buffer = new byte[8192];
                int len;
                while ((len = zis.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }
                
                zis.closeEntry();
                zos.closeEntry();
            }
            
            // 添加图片文件 ((原始路径+文档索引) -> 新名称)
            for (Map.Entry<String, String> imageRefEntry : imageReferences.entrySet()) {
                String key = imageRefEntry.getKey();
                String newName = imageRefEntry.getValue();
                
                // 解析key获取原始路径和文档索引
                String[] parts = key.split("\\|");
                String originalPath = parts[0];
                int docIndex = Integer.parseInt(parts[1]);
                
                // 从原始文档中提取图片并添加到新文档中
                copyImageFromSourceDoc(docPathList.get(docIndex), originalPath, newName, zos);
            }
        }
        
        // 更新[Content_Types].xml文件
        updateContentTypes(tempPath);
        
        // 更新word/_rels/document.xml.rels文件
        updateDocumentRels(tempPath);
        
        // 替换原文件
        Files.deleteIfExists(Paths.get(outputPath));
        Files.move(Paths.get(tempPath), Paths.get(outputPath));
        
        logger.info("图片复制完成");
    }
    
    /**
     * 更新[Content_Types].xml文件，添加图片类型声明
     * 
     * @param filePath 文件路径
     * @throws Exception IO异常
     */
    private void updateContentTypes(String filePath) throws Exception {
        // 读取并更新[Content_Types].xml
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(filePath));
             ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(filePath + ".tmp"))) {
            
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);
                
                if ("[Content_Types].xml".equals(entry.getName())) {
                    // 处理[Content_Types].xml文件
                    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
                    byte[] data = new byte[8192];
                    int len;
                    while ((len = zis.read(data)) > 0) {
                        buffer.write(data, 0, len);
                    }
                    
                    String content = new String(buffer.toByteArray(), "UTF-8");
                    
                    // 检查是否已存在图片类型声明，避免重复添加
                    if (!content.contains("image/jpeg") && !content.contains("image/png")) {
                        // 在</Types>标签前插入图片类型声明
                        StringBuilder newContent = new StringBuilder(content);
                        int insertPos = newContent.lastIndexOf("</Types>");
                        if (insertPos != -1) {
                            // 插入常用的图片类型声明
                            String imageTypes = 
                                "  <Default ContentType=\"image/jpeg\" Extension=\"jpeg\"/>\n" +
                                "  <Default ContentType=\"image/jpeg\" Extension=\"jpg\"/>\n" +
                                "  <Default ContentType=\"image/png\" Extension=\"png\"/>\n" +
                                "  <Default ContentType=\"image/gif\" Extension=\"gif\"/>\n" +
                                "  <Default ContentType=\"image/bmp\" Extension=\"bmp\"/>\n" +
                                "  <Default ContentType=\"image/tiff\" Extension=\"tiff\"/>\n" +
                                "  <Default ContentType=\"image/tiff\" Extension=\"tif\"/>\n";
                            
                            newContent.insert(insertPos, imageTypes);
                            zos.write(newContent.toString().getBytes("UTF-8"));
                        } else {
                            // 如果没有找到</Types>标签，直接写入原内容
                            zos.write(buffer.toByteArray());
                        }
                    } else {
                        // 如果已存在图片类型声明，直接写入原内容
                        zos.write(buffer.toByteArray());
                    }
                } else {
                    // 直接复制其他文件
                    byte[] buffer = new byte[8192];
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                }
                
                zis.closeEntry();
                zos.closeEntry();
            }
        }
        
        // 替换文件
        Files.deleteIfExists(Paths.get(filePath));
        Files.move(Paths.get(filePath + ".tmp"), Paths.get(filePath));
    }
    
    /**
     * 更新word/_rels/document.xml.rels文件，添加图片关系
     * 
     * @param filePath 文件路径
     * @throws Exception IO异常
     */
    private void updateDocumentRels(String filePath) throws Exception {
        // 读取并更新word/_rels/document.xml.rels
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(filePath));
             ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(filePath + ".tmp"))) {
            
            ZipEntry entry;
            int imageIdCounter = 9; // 从rId9开始
            while ((entry = zis.getNextEntry()) != null) {
                ZipEntry newEntry = new ZipEntry(entry.getName());
                zos.putNextEntry(newEntry);
                
                if ("word/_rels/document.xml.rels".equals(entry.getName())) {
                    // 处理word/_rels/document.xml.rels文件
                    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
                    byte[] data = new byte[8192];
                    int len;
                    while ((len = zis.read(data)) > 0) {
                        buffer.write(data, 0, len);
                    }
                    
                    String content = new String(buffer.toByteArray(), "UTF-8");
                    
                    // 删除现有的图片关系（对于第一个文档）
                    content = content.replaceAll("<Relationship[^>]*Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"[^>]*/>", "");
                    
                    // 在</Relationships>标签前插入新的图片关系
                    StringBuilder newContent = new StringBuilder(content);
                    int insertPos = newContent.lastIndexOf("</Relationships>");
                    if (insertPos != -1) {
                        StringBuilder imageRels = new StringBuilder();
                        // 添加图片关系
                        for (Map.Entry<String, String> imageRefEntry : imageReferences.entrySet()) {
                            String newName = imageRefEntry.getValue();
                            imageIdCounter++;
                            imageRels.append("  <Relationship Id=\"rId").append(imageIdCounter)
                                    .append("\" Target=\"media/").append(newName)
                                    .append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"/>\n");
                        }
                        
                        newContent.insert(insertPos, imageRels.toString());
                        zos.write(newContent.toString().getBytes("UTF-8"));
                    } else {
                        // 如果没有找到</Relationships>标签，直接写入原内容
                        zos.write(buffer.toByteArray());
                    }
                } else {
                    // 直接复制其他文件
                    byte[] buffer = new byte[8192];
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                }
                
                zis.closeEntry();
                zos.closeEntry();
            }
        }
        
        // 替换文件
        Files.deleteIfExists(Paths.get(filePath));
        Files.move(Paths.get(filePath + ".tmp"), Paths.get(filePath));
    }
    
    /**
     * 从源文档复制图片到目标ZIP输出流
     * 
     * @param sourceDocPath 源文档路径
     * @param originalPath 原始图片路径
     * @param newName 新图片名称
     * @param zos ZIP输出流
     * @throws Exception IO异常
     */
    private void copyImageFromSourceDoc(String sourceDocPath, String originalPath, String newName, ZipOutputStream zos) throws Exception {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(sourceDocPath))) {
            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (entry.getName().equals("word/" + originalPath)) {
                    // 找到图片文件，复制到目标ZIP流中
                    ZipEntry newEntry = new ZipEntry("word/media/" + newName);
                    zos.putNextEntry(newEntry);
                    
                    byte[] buffer = new byte[8192];
                    int len;
                    while ((len = zis.read(buffer)) > 0) {
                        zos.write(buffer, 0, len);
                    }
                    
                    zis.closeEntry();
                    zos.closeEntry();
                    break;
                }
                zis.closeEntry();
            }
        }
    }
}