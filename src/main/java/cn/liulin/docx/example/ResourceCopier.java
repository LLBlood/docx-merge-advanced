package cn.liulin.docx.example;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:39
 */
public class ResourceCopier {

    public static Map<String, String> copyImages(WordprocessingMLPackage doc1, WordprocessingMLPackage doc2) {
        Map<String, String> imageRelMap = new HashMap<>();
        try {
            System.out.println("🔍 开始复制图片资源...");
            
            RelationshipsPart relPart2 = doc2.getMainDocumentPart().getRelationshipsPart();
            if (relPart2 == null) {
                System.out.println("⚠️ 文档2没有关系部分，跳过图片复制");
                return imageRelMap;
            }
            
            List<Relationship> relationships = relPart2.getRelationships().getRelationship();
            System.out.println("📄 文档2中共有 " + relationships.size() + " 个关系");
            
            // 复制图片部件从doc2到doc1
            int copiedImages = 0;
            for (Relationship rel : relationships) {
                System.out.println("🔗 处理关系: ID=" + rel.getId() + ", Type=" + rel.getType() + ", Target=" + rel.getTarget());
                
                // 只处理图片关系
                if (Namespaces.IMAGE.equals(rel.getType())) {
                    String target = rel.getTarget(); // e.g., "media/image1.png"
                    System.out.println("📎 发现图片关系: " + target);

                    // 构造 PartName
                    PartName partName = new PartName("/" + target);
                    System.out.println("📂 尝试通过PartName获取图片部件: " + partName.getName());

                    // 从 doc2 获取图片部件
                    Part imgPart = doc2.getParts().get(partName);
                    if (imgPart == null) {
                        // 尝试通过关系获取图片部件
                        try {
                            System.out.println("🔄 通过关系获取图片部件...");
                            imgPart = relPart2.getPart(rel);
                        } catch (Exception e) {
                            System.err.println("❌ 无法通过关系获取图片部件: " + target + ", 错误: " + e.getMessage());
                            continue;
                        }
                    }
                    
                    if (imgPart == null) {
                        System.err.println("❌ 图片部件不存在: " + target);
                        continue;
                    }
                    
                    System.out.println("_TypeInfo: " + imgPart.getClass().getName());
                    
                    if (!(imgPart instanceof BinaryPartAbstractImage)) {
                        System.err.println("❌ 图片部件类型错误: " + target + ", 实际类型: " + imgPart.getClass().getName());
                        continue;
                    }

                    // 将图片部件添加到doc1中
                    System.out.println("📥 正在复制图片: " + target);
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
                    System.out.println("✅ 图片已复制: " + target + ", 旧 relId: " + oldId + ", 新 relId: " + newId);
                } else {
                    System.out.println("➡️ 跳过非图片关系: " + rel.getType());
                }
            }
            
            System.out.println("✅ 图片复制完成，共复制 " + copiedImages + " 张图片");

        } catch (Exception e) {
            System.err.println("❌ 复制图片失败: " + e.getMessage());
            e.printStackTrace();
        }
        
        return imageRelMap;
    }
}