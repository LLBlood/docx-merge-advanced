package cn.liulin.docx;

import cn.liulin.docx.util.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.slf4j.Logger;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Paths;
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

            // 在docx4j加载文档之前，预处理原始文档，替换不兼容标签
            List<String> processedDocPathList = PreprocessDocumentUtil.preprocessDocumentList(docPathList);

            // 加载所有数据
            List<WordprocessingMLPackage> docList = WordProcessingUtils.loadDocList(processedDocPathList);

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

            // 保存第一个文档的节属性设置，以其为基础合并文档
            MainDocumentPart main1 = docList.get(0).getMainDocumentPart();

            // 移除文档网格设置
            WordProcessingUtils.removeDocumentGridSettingsList(docList);

            // 将合并doc 的所有内容追加到 doc1
            WordProcessingUtils.addDocListToBase(main1, docList);

            // 确保输出目录存在
            File output = new File(outputPath);
            if (!output.getParentFile().exists()) {
                output.getParentFile().mkdirs();
            }

            // 保存文档
            docList.get(0).save(output);
            logger.info("文档已成功合并并保存到: {}", outputPath);
            
            // 清理临时文件
            for (String s : processedDocPathList) {
                Files.deleteIfExists(Paths.get(s));
            }
            
            LoggerUtil.logMethodExit(logger, "mergeList", "合并完成");
        } catch (Exception e) {
            LoggerUtil.logMethodException(logger, "mergeList", e);
            throw e;
        }
    }
}