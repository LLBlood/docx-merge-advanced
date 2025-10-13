package cn.liulin.docx.example;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class DocxMerger {

    public void merge(String doc1Path, String doc2Path, String outputPath) throws Exception {
        System.out.println("🔄 开始合并文档...");

        // 加载两个文档
        WordprocessingMLPackage doc1 = WordprocessingMLPackage.load(new File(doc1Path));
        WordprocessingMLPackage doc2 = WordprocessingMLPackage.load(new File(doc2Path));

        MainDocumentPart main1 = doc1.getMainDocumentPart();
        MainDocumentPart main2 = doc2.getMainDocumentPart();

        // ✅ 1. 处理样式冲突（重命名 doc2 的样式）
        StyleRemapper.renameStyles(doc2, "_DOC2");

        // ✅ 2. 映射编号（避免列表编号混乱）
        NumberingMapper.mapNumbering(doc1, doc2);

        // ✅ 3. 复制图片、表格等资源（处理关系）
        Map<String, String> imageRelMap = ResourceCopier.copyImages(doc1, doc2);

        // ✅ 4. 更新图片引用关系
        if (!imageRelMap.isEmpty()) {
            updateImageReferences(main2, imageRelMap);
        }

        // ✅ 5. 保存第一个文档的节属性设置
        SectPr firstDocSectPr = getPgSzSettings(main1);
        
        // ✅ 6. 在合并前添加分节符，保持文档页面设置独立
        addSectionBreak(main1);

        // ✅ 7. 将 doc2 的所有内容追加到 doc1
        // 使用 addObject() 以触发样式/字体发现
        System.out.println("📄 开始合并文档内容，doc2内容项数: " + main2.getContent().size());
        int objectCount = 0;
        for (Object obj : main2.getContent()) {
            objectCount++;
            System.out.println("📑 正在添加第 " + objectCount + " 个内容项: " + obj.getClass().getSimpleName());
            main1.addObject(obj);
        }
        System.out.println("✅ 文档内容合并完成，共添加 " + objectCount + " 个内容项");

        // 修复对齐元素，确保符合Open XML规范
        fixJustificationElements(doc1);

        // ✅ 8. 获取 doc2 的最后一个节属性（SectPr）
        SectPr lastSectPr = findLastSectPr(main2);
        
        // 如果找不到最后一个节属性，则尝试获取文档默认的节属性
        if (lastSectPr == null) {
            lastSectPr = getPgSzSettings(main2);
        }

        // ✅ 9. 如果 doc2 有节结束（SectPr），则在合并后添加一个新节段落
        if (lastSectPr != null) {
            ObjectFactory factory = Context.getWmlObjectFactory();  // ✅ 正确方式
            P newSection = factory.createP();

            PPr pPr = factory.createPPr();
            // 深拷贝 sectPr，避免引用共享
            pPr.setSectPr((SectPr) org.docx4j.XmlUtils.deepCopy(lastSectPr));
            newSection.setPPr(pPr);

            // 使用 addObject() 添加，触发样式/字体等处理
            main1.addObject(newSection);
        } else if (firstDocSectPr != null) {
            // 如果 doc2 没有节属性，但第一个文档有，则使用第一个文档的节属性
            ObjectFactory factory = Context.getWmlObjectFactory();
            P newSection = factory.createP();
            PPr pPr = factory.createPPr();
            // 深拷贝 sectPr，避免引用共享
            pPr.setSectPr((SectPr) org.docx4j.XmlUtils.deepCopy(firstDocSectPr));
            newSection.setPPr(pPr);
            main1.addObject(newSection);
        } else {
            // 如果都没有节属性，则添加一个默认的节属性来保持页面设置
            ObjectFactory factory = Context.getWmlObjectFactory();
            P newSection = factory.createP();
            PPr pPr = factory.createPPr();
            SectPr sectPr = factory.createSectPr();
            pPr.setSectPr(sectPr);
            newSection.setPPr(pPr);
            main1.addObject(newSection);
        }

        // ✅ 10. 确保输出目录存在
        File output = new File(outputPath);
        if (!output.getParentFile().exists()) {
            output.getParentFile().mkdirs();
        }

        // ✅ 11. 保存文档
        doc1.save(output);
        System.out.println("✅ 文档已成功合并并保存到: " + outputPath);
    }

    /**
     * 修复对齐元素，确保所有 jc 元素都有 val 属性
     */
    private void fixJustificationElements(WordprocessingMLPackage doc) {
        try {
            // 获取文档的XML内容
            String xmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);
            
            // 修复所有缺失val属性的jc元素
            xmlContent = fixMissingValAttributes(xmlContent);
            
            // 修复重复的ID问题
            xmlContent = fixDuplicateIdsInXml(xmlContent);
            
            // 将更新后的XML内容重新设置到文档中
            org.docx4j.wml.Document document = (org.docx4j.wml.Document) 
                XmlUtils.unmarshalString(xmlContent);
            doc.getMainDocumentPart().setJaxbElement(document);
            
            System.out.println("✅ 对齐元素和ID修复完成");
        } catch (Exception e) {
            System.err.println("⚠️ 修复对齐元素时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 修复XML中缺失val属性的jc元素
     */
    private String fixMissingValAttributes(String xmlContent) {
        // 修复自闭合的jc标签缺失val属性的问题
        xmlContent = xmlContent.replaceAll(
            "<w:jc\\s*/>", 
            "<w:jc w:val=\"center\"/>");
            
        // 修复带有属性但缺少val属性的jc开始标签
        xmlContent = xmlContent.replaceAll(
            "<w:jc((?![^>]*\\bw:val\\b)[^>]*/?)>", 
            "<w:jc w:val=\"center\"$1>");
            
        return xmlContent;
    }
    
    /**
     * 修复XML中的重复ID问题
     */
    private String fixDuplicateIdsInXml(String xmlContent) {
        // 使用正则表达式查找并修复重复的ID
        // 这里我们简单地为所有bookmarkStart和bookmarkEnd元素生成新的唯一ID
        java.util.regex.Pattern bookmarkStartPattern = java.util.regex.Pattern.compile(
            "<w:bookmarkStart[^>]*w:id\\s*=\\s*\"([^\"]*)\"[^>]*/>");
        java.util.regex.Matcher matcher = bookmarkStartPattern.matcher(xmlContent);
        
        java.util.Set<String> usedIds = new java.util.HashSet<>();
        java.util.Map<String, String> idReplacements = new java.util.HashMap<>();
        
        // 收集所有现有的ID
        while (matcher.find()) {
            String id = matcher.group(1);
            if (usedIds.contains(id)) {
                // 生成新的唯一ID
                String newId = generateUniqueID(usedIds);
                idReplacements.put(id, newId);
                usedIds.add(newId);
            } else {
                usedIds.add(id);
            }
        }
        
        // 也检查bookmarkEnd元素
        java.util.regex.Pattern bookmarkEndPattern = java.util.regex.Pattern.compile(
            "<w:bookmarkEnd[^>]*w:id\\s*=\\s*\"([^\"]*)\"[^>]*/>");
        matcher = bookmarkEndPattern.matcher(xmlContent);
        
        while (matcher.find()) {
            String id = matcher.group(1);
            if (usedIds.contains(id)) {
                // 生成新的唯一ID
                String newId = generateUniqueID(usedIds);
                idReplacements.put(id, newId);
                usedIds.add(newId);
            } else {
                usedIds.add(id);
            }
        }
        
        // 替换重复的ID
        for (java.util.Map.Entry<String, String> entry : idReplacements.entrySet()) {
            xmlContent = xmlContent.replaceAll(
                "w:id\\s*=\\s*\"" + java.util.regex.Pattern.quote(entry.getKey()) + "\"",
                "w:id=\"" + entry.getValue() + "\"");
        }
        
        return xmlContent;
    }
    
    /**
     * 生成唯一ID
     */
    private String generateUniqueID(java.util.Set<String> existingIds) {
        String newId;
        do {
            newId = String.valueOf(System.currentTimeMillis() % 1000000 + Math.round(Math.random() * 1000));
        } while (existingIds.contains(newId));
        return newId;
    }

    /**
     * 在第一个文档末尾添加分节符，确保第二个文档保持其原始页面设置
     */
    private void addSectionBreak(MainDocumentPart main1) {
        try {
            ObjectFactory factory = Context.getWmlObjectFactory();
            P sectionBreakParagraph = factory.createP();
            PPr pPr = factory.createPPr();
            
            // 创建分节符
            SectPr sectPr = factory.createSectPr();
            
            // 设置分节符类型为下一页（NEXT_PAGE）
            // 这样可以确保第二个文档从新的一页开始，并保持其原始页面设置
            SectPr.Type sectType = factory.createSectPrType();
            sectType.setVal("nextPage"); // 下一页分节符
            sectPr.setType(sectType);
            
            // 保留第一个文档的页面设置
            SectPr firstDocSectPr = getPgSzSettings(main1);
            if (firstDocSectPr != null && firstDocSectPr.getPgSz() != null) {
                // 复制第一页的页面大小设置
                sectPr.setPgSz(XmlUtils.deepCopy(firstDocSectPr.getPgSz()));
            }
            
            pPr.setSectPr(sectPr);
            sectionBreakParagraph.setPPr(pPr);
            
            // 添加分节符段落到第一个文档末尾
            main1.addObject(sectionBreakParagraph);
            System.out.println("✅ 已添加分节符以保持页面设置独立");
        } catch (Exception e) {
            System.err.println("⚠️ 添加分节符时出错: " + e.getMessage());
        }
    }

    /**
     * 更新文档中的图片引用关系
     */
    private void updateImageReferences(MainDocumentPart doc2Part, Map<String, String> imageRelMap) {
        if (imageRelMap.isEmpty()) {
            System.out.println("⚠️ 没有图片关系需要更新");
            return;
        }
        
        System.out.println("🔄 开始更新图片引用关系，共 " + imageRelMap.size() + " 个关系需要更新");
        
        try {
            // 获取文档的XML内容
            String xmlContent = XmlUtils.marshaltoString(doc2Part.getJaxbElement(), true, true);
            
            System.out.println("📄 原始XML内容长度: " + xmlContent.length());
            
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
                    System.out.println("🔄 第一步替换: " + oldRelId + " -> " + tempId);
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
                    System.out.println("🔄 第二步替换: " + tempId + " -> " + newRelId);
                    updatedXmlContent = updatedXmlContent.replace(
                        "r:embed=\"" + tempId + "\"", 
                        "r:embed=\"" + newRelId + "\""
                    );
                }
            }
            
            // 检查是否真的有更新
            if (!updatedXmlContent.equals(xmlContent)) {
                System.out.println("✅ XML内容已更新");
                // 将更新后的XML内容重新设置到文档中
                org.docx4j.wml.Document document = (org.docx4j.wml.Document) 
                    XmlUtils.unmarshalString(updatedXmlContent);
                doc2Part.setJaxbElement(document);
            } else {
                System.out.println("ℹ️ XML内容未发生变化");
            }
            
            System.out.println("✅ 图片引用关系更新完成");
        } catch (Exception e) {
            System.err.println("❌ 更新图片引用关系失败: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 查找 MainDocumentPart 中最后一个带有节属性的段落
     */
    private SectPr findLastSectPr(MainDocumentPart part) {
        List<Object> content = part.getContent();
        // 从后往前找
        for (int i = content.size() - 1; i >= 0; i--) {
            Object obj = content.get(i);
            if (obj instanceof P) {
                P p = (P) obj;
                PPr ppr = p.getPPr();
                if (ppr != null && ppr.getSectPr() != null) {
                    return ppr.getSectPr();
                }
            }
        }
        return null;
    }
    
    /**
     * 获取文档的页面设置（页面大小和方向）
     */
    private SectPr getPgSzSettings(MainDocumentPart part) {
        // 获取文档的body元素
        org.docx4j.wml.Document wmlDocument = part.getJaxbElement();
        if (wmlDocument != null && wmlDocument.getBody() != null) {
            return wmlDocument.getBody().getSectPr();
        }
        return null;
    }
}