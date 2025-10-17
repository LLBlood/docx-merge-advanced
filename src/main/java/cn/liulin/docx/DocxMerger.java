package cn.liulin.docx;

import cn.liulin.docx.util.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;


/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class DocxMerger {
    private static final Logger logger = LoggerUtil.getLogger(DocxMerger.class);

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

            // 分批处理文档，每批处理10个文档
            WordprocessingMLPackage resultDoc = null;
            List<String> batch = new ArrayList<>();
            
            for (int i = 0; i < docPathList.size(); i++) {
                batch.add(docPathList.get(i));
                
                // 每10个文档处理一次，或者到达最后一个文档时处理
                if (batch.size() >= 10 || i == docPathList.size() - 1) {
                    if (resultDoc == null) {
                        // 第一批文档，创建基础文档
                        resultDoc = mergeBatch(batch, null);
                    } else {
                        // 后续批次，将结果合并到已有文档中
                        resultDoc = mergeBatch(batch, resultDoc);
                    }
                    batch.clear();
                    
                    // 提示当前进度
                    logger.info("已处理 {}/{} 个文档", Math.min(i + 1, docPathList.size()), docPathList.size());
                }
            }

            // 保存最终文档
            resultDoc.save(output);
            logger.info("文档已成功合并并保存到: {}", outputPath);
            
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

        // 加载当前批次数据
        List<WordprocessingMLPackage> docList = WordProcessingUtils.loadDocList(processedDocPathList);

        WordprocessingMLPackage resultDoc;
        if (baseDoc == null) {
            // 如果没有基础文档，使用第一个文档作为基础
            resultDoc = docList.get(0);
        } else {
            // 如果已有基础文档，将其作为第一个文档
            docList.add(0, baseDoc);
            resultDoc = baseDoc;
        }

        // 处理样式冲突
        StyleReMapperUtil.renameDocListStyles(docList);

        // 映射编号（避免列表编号混乱）
        NumberingMapperUtil.mapNumbering(docList);

        // 复制图片、表格等资源（处理关系）
        ResourceCopierUtil.copyImages(docList);

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
        
        logger.info("批次处理完成");
        return resultDoc;
    }
}