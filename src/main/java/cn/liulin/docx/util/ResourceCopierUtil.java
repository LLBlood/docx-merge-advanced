package cn.liulin.docx.util;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
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
     * 保存图片引用路径，不直接复制图片
     * 
     * @param docPath 包含WordprocessingMLPackage对象的文档列表
     * @param imageReferences 图片引用映射集合
     * @param imageCounter 图片计数器
     * @return 更新后的图片计数器
     */
    public static int saveImageReferences(List<WordprocessingMLPackage> docPath, 
                                          Map<String, Map<String, String>> imageReferences, 
                                          int imageCounter) {
        LoggerUtil.logMethodEntry(logger, "saveImageReferences", docPath != null ? docPath.size() : 0);

        int updatedImageCounter = imageCounter;
        Map<String, String> imageRelMap = new HashMap<>(100);
        try {
            logger.info("开始保存图片引用路径...");
            
            // 遍历所有文档，保存图片引用路径
            for (int i = 0; i < docPath.size(); i++) {
                WordprocessingMLPackage tempDoc = docPath.get(i);
                RelationshipsPart relPart = tempDoc.getMainDocumentPart().getRelationshipsPart();
                if (relPart == null) {
                    logger.warn("文档没有关系部分，跳过图片引用保存");
                    continue;
                }

                List<Relationship> relationships = relPart.getRelationships().getRelationship();
                logger.debug("文档中共有 {} 个关系", relationships.size());

                // 保存图片引用路径
                for (Relationship rel : relationships) {
                    logger.debug("处理关系: ID={}, Type={}, Target={}", rel.getId(), rel.getType(), rel.getTarget());

                    // 只处理图片关系
                    if (Namespaces.IMAGE.equals(rel.getType())) {
                        String target = rel.getTarget();
                        logger.info("发现图片关系: {}", target);
                        
                        // 重命名图片
                        updatedImageCounter++;
                        String extension = "";
                        int lastDotIndex = target.lastIndexOf('.');
                        if (lastDotIndex > 0) {
                            extension = target.substring(lastDotIndex);
                        }
                        String newName = String.format("image_%05d%s", updatedImageCounter, extension);
                        
                        // 保存图片引用路径映射（包含文档信息）
                        Map<String, String> docImageRef = new HashMap<>();
                        docImageRef.put("originalPath", target);
                        docImageRef.put("newName", newName);
                        docImageRef.put("docIndex", String.valueOf(i));
                        imageReferences.put(rel.getId(), docImageRef);
                        
                        // 保存关系ID映射，用于后续更新引用
                        imageRelMap.put(rel.getId(), "rId" + updatedImageCounter);
                        
                        logger.info("图片引用已保存: {} -> {}", target, newName);
                    } else {
                        logger.debug("跳过非图片关系: {}", rel.getType());
                    }
                }
                
                // 更新当前文档的图片引用关系
                updateImageReferences(tempDoc, imageRelMap);
            }
            
            logger.info("图片引用路径保存完成");

        } catch (Exception e) {
            LoggerUtil.logMethodException(logger, "saveImageReferences", e);
        }
        
        LoggerUtil.logMethodExit(logger, "saveImageReferences", "图片引用保存完成");
        return updatedImageCounter;
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