package cn.liulin.docx.util;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Document;
import org.slf4j.Logger;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:39
 */
public class ResourceCopierUtil {
    private static final Logger logger = LoggerUtil.getLogger(ResourceCopierUtil.class);

    /**
     * 复制文档列表中除第一个文档外的所有文档的图片资源到第一个文档中
     * 该方法会遍历文档列表，将每个文档中的图片复制到第一个文档，并更新图片引用关系
     * 
     * @param docPath 包含WordprocessingMLPackage对象的文档列表，第一个文档作为目标文档
     */
    public static void copyImages(List<WordprocessingMLPackage> docPath) {
        LoggerUtil.logMethodEntry(logger, "copyImages", docPath != null ? docPath.size() : 0);

        assert docPath != null;
        WordprocessingMLPackage doc1 = docPath.get(0);
        try {
            logger.info("开始复制图片资源...");
            // 遍历除第一个文档外的所有文档，复制其中的图片资源
            for (int i = 1; i < docPath.size(); i++) {
                WordprocessingMLPackage tempDoc = docPath.get(i);
                Map<String, String> imageRelMap = new HashMap<>();
                RelationshipsPart relPart2 = tempDoc.getMainDocumentPart().getRelationshipsPart();
                if (relPart2 == null) {
                    logger.warn("文档没有关系部分，跳过图片复制");
                    continue;
                }

                List<Relationship> relationships = relPart2.getRelationships().getRelationship();
                logger.debug("文档中共有 {} 个关系", relationships.size());

                // 复制图片部件从doc到doc1
                int copiedImages = 0;
                for (Relationship rel : relationships) {
                    logger.debug("处理关系: ID={}, Type={}, Target={}", rel.getId(), rel.getType(), rel.getTarget());

                    // 只处理图片关系
                    if (Namespaces.IMAGE.equals(rel.getType())) {
                        String target = rel.getTarget();
                        logger.info("发现图片关系: {}", target);

                        // 构造 PartName
                        PartName partName = new PartName("/" + target);
                        logger.debug("尝试通过PartName获取图片部件: {}", partName.getName());

                        // 从 doc2 获取图片部件
                        Part imgPart = tempDoc.getParts().get(partName);
                        if (imgPart == null) {
                            // 尝试通过关系获取图片部件
                            try {
                                logger.debug("通过关系获取图片部件...");
                                imgPart = relPart2.getPart(rel);
                            } catch (Exception e) {
                                logger.error("无法通过关系获取图片部件: {}, 错误: {}", target, e.getMessage());
                                continue;
                            }
                        }

                        if (imgPart == null) {
                            logger.error("图片部件不存在: {}", target);
                            continue;
                        }

                        logger.debug("_TypeInfo: {}", imgPart.getClass().getName());

                        if (!(imgPart instanceof BinaryPartAbstractImage)) {
                            logger.error("图片部件类型错误: {}, 实际类型: {}", target, imgPart.getClass().getName());
                            continue;
                        }

                        // 将图片部件添加到doc1中
                        logger.debug("正在复制图片: {}", target);
                        BinaryPartAbstractImage binaryImage = (BinaryPartAbstractImage) imgPart;
                        BinaryPartAbstractImage newImagePart = BinaryPartAbstractImage.createImagePart(
                                doc1,
                                doc1.getMainDocumentPart(),
                                binaryImage.getBytes()
                        );

                        // 获取新生成的关系 ID
                        String newId = newImagePart.getSourceRelationships().get(0).getId();
                        String oldId = rel.getId();

                        // 建立旧ID到新ID的映射
                        imageRelMap.put(oldId, newId);

                        copiedImages++;
                        logger.info("图片已复制: {}, 旧 relId: {}, 新 relId: {}", target, oldId, newId);
                    } else {
                        logger.debug("跳过非图片关系: {}", rel.getType());
                    }
                }
                logger.info("文档{}图片复制完成，共复制 {} 张图片", (i + 1), copiedImages);
                updateImageReferences(tempDoc, imageRelMap);
            }

        } catch (Exception e) {
            LoggerUtil.logMethodException(logger, "copyImages", e);
        }
        
        LoggerUtil.logMethodExit(logger, "copyImages", "图片复制完成");
    }


    /**
     * 更新文档中的图片引用关系
     */
    private static void updateImageReferences(WordprocessingMLPackage doc2Package, Map<String, String> imageRelMap) {
        if (imageRelMap.isEmpty()) {
            logger.warn("没有图片关系需要更新");
            return;
        }

        logger.info("开始更新图片引用关系，共 {} 个关系需要更新", imageRelMap.size());

        try {
            // 获取文档的XML内容
            String xmlContent = XmlUtils.marshaltoString(doc2Package.getMainDocumentPart().getJaxbElement(), true, true);

            logger.debug("原始XML内容长度: {}", xmlContent.length());

            // 创建临时映射，避免替换冲突
            String updatedXmlContent = xmlContent;

            // 使用临时标识符避免链式替换问题
            String tempPrefix = "TEMP_REPLACE_";
            int counter = 0;

            // 第一步：将所有旧ID替换为临时ID
            for (Map.Entry<String, String> entry : imageRelMap.entrySet()) {
                String oldRelId = entry.getKey();
                String tempId = tempPrefix + (counter++);

                // 检查是否存在该引用
                if (updatedXmlContent.contains("r:embed=\"" + oldRelId + "\"")) {
                    logger.debug("第一步替换: {} -> {}", oldRelId, tempId);
                    updatedXmlContent = updatedXmlContent.replace(
                            "r:embed=\"" + oldRelId + "\"",
                            "r:embed=\"" + tempId + "\""
                    );
                }
            }

            // 第二步：将临时ID替换为新ID
            counter = 0;
            for (Map.Entry<String, String> entry : imageRelMap.entrySet()) {
                String newRelId = entry.getValue();
                String tempId = tempPrefix + counter++;

                if (updatedXmlContent.contains("r:embed=\"" + tempId + "\"")) {
                    logger.debug("第二步替换: {} -> {}", tempId, newRelId);
                    updatedXmlContent = updatedXmlContent.replace(
                            "r:embed=\"" + tempId + "\"",
                            "r:embed=\"" + newRelId + "\""
                    );
                }
            }

            // 检查是否真的有更新
            if (!updatedXmlContent.equals(xmlContent)) {
                logger.debug("XML内容已更新");
                // 将更新后的XML内容重新设置到文档中
                Document document = (Document)
                        XmlUtils.unmarshalString(updatedXmlContent);
                doc2Package.getMainDocumentPart().setJaxbElement(document);
            } else {
                logger.debug("XML内容未发生变化");
            }

            logger.info("图片引用关系更新完成");
        } catch (Exception e) {
            logger.error("更新图片引用关系失败: {}", e.getMessage(), e);
        }
    }
}