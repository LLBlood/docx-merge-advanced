package cn.liulin.docx;

import cn.liulin.docx.util.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
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
    
    // 存储所有图片引用路径 (关系ID -> {原始路径, 新名称, 文档索引})
    private Map<String, Map<String, String>> imageReferences = new HashMap<>();
    
    // 图片计数器
    private int imageCounter = 0;

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
            
            // 在保存后复制图片到最终文档
//            copyImagesToFinalDocument(docPathList, outputPath);
            
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
        
        // 直接在输出文件上操作，添加图片
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(outputPath));
             ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(outputPath + ".tmp"))) {
            
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
            
            // 添加图片文件
            for (Map.Entry<String, Map<String, String>> imageRefEntry : imageReferences.entrySet()) {
                Map<String, String> imageInfo = imageRefEntry.getValue();
                String originalPath = imageInfo.get("originalPath");
                String newName = imageInfo.get("newName");
                int docIndex = Integer.parseInt(imageInfo.get("docIndex"));
                
                // 从原始文档中提取图片并添加到新文档中
                copyImageFromSourceDoc(docPathList.get(docIndex), originalPath, newName, zos);
            }
        }
        
        // 替换原文件
        Files.deleteIfExists(Paths.get(outputPath));
        Files.move(Paths.get(outputPath + ".tmp"), Paths.get(outputPath));
        
        logger.info("图片复制完成");
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