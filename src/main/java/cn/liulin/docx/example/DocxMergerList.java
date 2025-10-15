package cn.liulin.docx.example;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.*;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;


/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class DocxMergerList {

    public void mergeList(List<String> docPathList, String outputPath) throws Exception {
        System.out.println("🔄 开始合并文档...");

        // 在docx4j加载文档之前，预处理原始文档，替换不兼容标签
        List<String> processedDocPathList = new ArrayList<>();
        for (String docPath : docPathList) {
            String outPath = preprocessDocument(docPath);
            processedDocPathList.add(outPath);
        }

        List<WordprocessingMLPackage> docList = new ArrayList<>();
        for (String processedDocPath : processedDocPathList) {
            WordprocessingMLPackage load = WordprocessingMLPackage.load(new File(processedDocPath));
            docList.add(load);
        }

        // ✅ 1. 处理样式冲突（重命名 doc1 和 doc2 的样式）
        for (int i = 0; i < docList.size(); i++) {
            StyleRemapper.renameStyles(docList.get(i), "_DOC" + (i + 1));
        }

        // ✅ 2. 合并样式定义（在重命名之后合并样式）
        mergeStyles(docList);
        
        // ✅ 3. 映射编号（避免列表编号混乱）
        NumberingMapper.mapNumbering(docList);

        // ✅ 4. 复制图片、表格等资源（处理关系）
        ResourceCopier.copyImages(docList);

        // 保存两个文档的格式信息（暂时保留但不处理表格边框）
        System.out.println("💾 开始保存文档的格式信息...");
        Map<String, String> formatProperties = TableFormatPreserver.saveDocumentFormat(docList);
        System.out.println("💾 格式信息保存完成，共保存 " + formatProperties.size() + " 个属性");

        // 在合并前应用默认字体大小
        System.out.println("📐 开始在合并前应用默认字体大小...");
        for (int i = 0; i < docList.size(); i++) {
            applyDefaultFontSizesBeforeMerge(docList.get(i), formatProperties, "doc" + (i + 1));
        }
        System.out.println("✅ 合并前默认字体大小应用完成");

        // ✅ 6. 保存第一个文档的节属性设置
        MainDocumentPart main1 = docList.get(0).getMainDocumentPart();
        SectPr firstDocSectPr = getPgSzSettings(main1);
        
        // ✅ 7. 移除文档网格设置
        for (WordprocessingMLPackage doc : docList) {
            removeDocumentGridSettings(doc);
        }

        // ✅ 8. 在合并前添加分节符，保持文档页面设置独立
        addSectionBreak(main1);

        // ✅ 9. 将 doc2 的所有内容追加到 doc1
        // 使用 addObject() 以触发样式/字体发现
        for (int i = 1; i < docList.size(); i++) {
            MainDocumentPart tempMain = docList.get(i).getMainDocumentPart();
            System.out.println("📄 开始合并文档内容，doc内容项数: " + tempMain.getContent().size());
            int objectCount = 0;
            for (Object obj : tempMain.getContent()) {
                objectCount++;
                System.out.println("📑 正在添加第 " + objectCount + " 个内容项: " + obj.getClass().getSimpleName());
                main1.addObject(obj);
            }
            System.out.println("✅ 文档内容合并完成，共添加 " + objectCount + " 个内容项");
        }


        // 修复对齐元素，确保符合Open XML规范（不处理表格边框）
        System.out.println("🔧 开始修复对齐元素...");
        fixJustificationElements(docList.get(0));
        System.out.println("🔧 对齐元素修复完成");

        // ✅ 10. 获取 doc2 的最后一个节属性（SectPr）
        SectPr lastSectPr = findLastSctPr(docList.get(docList.size() - 1).getMainDocumentPart());
        
        // 如果找不到最后一个节属性，则尝试获取文档默认的节属性
        if (lastSectPr == null) {
            lastSectPr = getPgSzSettings(docList.get(docList.size() - 1).getMainDocumentPart());
        }

        // ✅ 11. 如果 doc2 有节结束（SectPr），则在合并后添加一个新节段落
        if (lastSectPr != null) {
            ObjectFactory factory = Context.getWmlObjectFactory();  // ✅ 正确方式
            P newSection = factory.createP();

            PPr pPr = factory.createPPr();
            // 深拷贝 sectPr，避免引用共享
            SectPr sectPrCopy = (SectPr) XmlUtils.deepCopy(lastSectPr);
            pPr.setSectPr(sectPrCopy);
            newSection.setPPr(pPr);

            // 使用 addObject() 添加，触发样式/字体等处理
            main1.addObject(newSection);
            System.out.println("✅ 已添加doc2的节属性设置");
        } else if (firstDocSectPr != null) {
            // 如果 doc2 没有节属性，但第一个文档有，则使用第一个文档的节属性
            ObjectFactory factory = Context.getWmlObjectFactory();
            P newSection = factory.createP();
            PPr pPr = factory.createPPr();
            // 深拷贝 sectPr，避免引用共享
            SectPr sectPrCopy = (SectPr) XmlUtils.deepCopy(firstDocSectPr);
            pPr.setSectPr(sectPrCopy);
            newSection.setPPr(pPr);
            main1.addObject(newSection);
            System.out.println("✅ 已添加第一个文档的节属性设置");
        } else {
            // 如果都没有节属性，则添加一个默认的节属性来保持页面设置
            ObjectFactory factory = Context.getWmlObjectFactory();
            P newSection = factory.createP();
            PPr pPr = factory.createPPr();
            SectPr sectPr = factory.createSectPr();
            pPr.setSectPr(sectPr);
            newSection.setPPr(pPr);
            main1.addObject(newSection);
            System.out.println("✅ 已添加默认节属性设置");
        }

        // ✅ 12. 确保输出目录存在
        File output = new File(outputPath);
        if (!output.getParentFile().exists()) {
            output.getParentFile().mkdirs();
        }

        // ✅ 13. 保存文档
        docList.get(0).save(output);
        System.out.println("✅ 文档已成功合并并保存到: " + outputPath);
        
        // 清理临时文件
        for (String s : processedDocPathList) {
            Files.deleteIfExists(Paths.get(s));
        }
    }

    /**
     * 在docx4j加载前预处理文档，替换不兼容的标签
     * 
     * @param docPath 原始文档路径
     * @return 处理后的文档路径
     * @throws Exception 处理异常
     */
    private String preprocessDocument(String docPath) throws Exception {
        Path originalDoc = Paths.get(docPath);
        Path processedDoc = Files.createTempFile("processed_", ".docx");
        
        // 复制原始文档到临时文件
        Files.copy(originalDoc, processedDoc, StandardCopyOption.REPLACE_EXISTING);
        
        // 创建一个新的临时文件用于输出
        Path outputDoc = Files.createTempFile("output_", ".docx");
        
        try (ZipFile zipFile = new ZipFile(processedDoc.toFile());
             ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(outputDoc.toFile()))) {
            
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
                        xmlContent = xmlContent.replaceAll("</w:start>", "</w:left>");
                        xmlContent = xmlContent.replaceAll("<w:end\\b", "<w:right");
                        xmlContent = xmlContent.replaceAll("</w:end>", "</w:right>");
                        
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

    /**
     * 合并两个文档的样式定义
     * 保留第一个文档的样式，添加第二个文档中独有的样式
     * 
     * @param docPath 文档
     */
    private void mergeStyles(List<WordprocessingMLPackage> docPath) {
        try {
            StyleDefinitionsPart stylePart1 = docPath.get(0).getMainDocumentPart().getStyleDefinitionsPart();
            Styles styles1 = stylePart1.getJaxbElement();
            // 创建一个映射来跟踪已存在的样式ID
            Map<String, Style> existingStyles = new HashMap<>();
            if (styles1.getStyle() != null) {
                for (Style style : styles1.getStyle()) {
                    if (style.getStyleId() != null) {
                        existingStyles.put(style.getStyleId(), style);
                    }
                }
            }

            for (int i = 1; i < docPath.size(); i++) {
                StyleDefinitionsPart tempStylePart = docPath.get(i).getMainDocumentPart().getStyleDefinitionsPart();
                Styles tempStyles = tempStylePart.getJaxbElement();
                // 遍历接下来的文档的样式
                for (Style tempStyle : tempStyles.getStyle()) {
                    String styleId = tempStyle.getStyleId();
                    if (styleId != null) {
                        // 检查样式是否已存在
                        if (!existingStyles.containsKey(styleId)) {
                            // 样式不存在，添加到第一个文档中
                            styles1.getStyle().add(tempStyle);
                            System.out.println("➕ 添加样式: " + styleId);
                        } else {
                            // 样式已存在，我们需要检查是否是重命名的样式
                            // 如果是重命名的样式（包含_DOC2后缀），则替换原始样式
                            if (styleId.contains("_DOC")) {
                                // 找到对应的原始样式ID
                                String originalStyleId = styleId.substring(0, styleId.indexOf("_DOC")); // 移除"_DOC"后缀

                                if (existingStyles.containsKey(originalStyleId)) {
                                    // 替换原始样式
                                    Style originalStyle = existingStyles.get(originalStyleId);
                                    int index = styles1.getStyle().indexOf(originalStyle);
                                    if (index >= 0) {
                                        styles1.getStyle().set(index, tempStyle);
                                        System.out.println("🔄 替换样式: " + originalStyleId + " -> " + styleId);
                                    }
                                }
                            } else {
                                // 保留第一个文档的样式定义
                                System.out.println("🔁 保留已存在的样式: " + styleId);
                            }
                        }
                    }
                }
            }
            
            System.out.println("✅ 样式合并完成");
        } catch (Exception e) {
            System.err.println("⚠️ 合并样式时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 修复对齐元素，确保所有 jc 元素都有 val 属性
     */
    private void fixJustificationElements(WordprocessingMLPackage doc) {
        try {
            // 直接使用docx4j API获取XML内容，不再需要通过ZIP方式读取
            String xmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);
            System.out.println("📄 docx4j读取的主文档XML内容长度: " + xmlContent.length());
            
            // 使用replace方法修复所有缺失val属性的jc标签
            xmlContent = fixMissingValAttributes(xmlContent);
            
            // 将更新后的XML内容重新设置到文档对象中
            Document document = (Document)
                XmlUtils.unmarshalString(xmlContent);
            doc.getMainDocumentPart().setJaxbElement(document);
            
            // 修复样式文档中的对齐元素
            StyleDefinitionsPart stylePart = doc.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylePart != null) {
                String styleXmlContent = XmlUtils.marshaltoString(stylePart.getJaxbElement(), true, true);
                System.out.println("📄 原始样式XML内容长度: " + styleXmlContent.length());
                
                // 使用replace方法修复所有缺失val属性的jc标签
                styleXmlContent = fixMissingValAttributes(styleXmlContent);
                
                // 将更新后的XML内容重新设置到样式部分中
                Styles styles = (Styles) XmlUtils.unmarshalString(styleXmlContent);
                stylePart.setJaxbElement(styles);
            }
            
            System.out.println("✅ 对齐元素修复完成");
        } catch (Exception e) {
            System.err.println("⚠️ 修复对齐元素时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 修复XML中缺失val属性的jc元素
     */
    private String fixMissingValAttributes(String xmlContent) {
        System.out.println("🔗 开始修复缺失val属性的对齐元素");
        
        // 使用replace方法修复所有缺失val属性的jc标签
        int beforeFix = xmlContent.length();
        xmlContent = xmlContent.replace("<w:jc/>", "<w:jc w:val=\"left\"/>");
        int afterFix = xmlContent.length();
        System.out.println("🔗 修复缺失val属性的jc标签: " + (afterFix - beforeFix) + " 字符变化");
            
        return xmlContent;
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
            if (firstDocSectPr != null) {
                // 复制第一页的页面大小设置
                if (firstDocSectPr.getPgSz() != null) {
                    sectPr.setPgSz(XmlUtils.deepCopy(firstDocSectPr.getPgSz()));
                }
                
                // 复制第一页的页边距设置
                if (firstDocSectPr.getPgMar() != null) {
                    sectPr.setPgMar(XmlUtils.deepCopy(firstDocSectPr.getPgMar()));
                }
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
     * 查找 MainDocumentPart 中最后一个带有节属性的段落
     */
    private SectPr findLastSctPr(MainDocumentPart part) {
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
        Document wmlDocument = part.getJaxbElement();
        if (wmlDocument != null && wmlDocument.getBody() != null) {
            return wmlDocument.getBody().getSectPr();
        }
        return null;
    }
    
    /**
     * 移除段落中的对齐到网络设置
     * 
     * @param xmlContent XML内容
     * @return 修复后的XML内容
     */
    private String removeParagraphSnapToGridSettings(String xmlContent) {
        System.out.println("📐 开始移除段落中的对齐到网络设置（合并前处理）");
        
        // 移除段落属性中的snapToGrid设置
        int beforeRemoval = xmlContent.length();
        
        // 处理段落属性标签内包含snapToGrid属性的情况
        xmlContent = xmlContent.replaceAll(
            "(<w:pPr[^>]*?)\\s+w:snapToGrid\\s*=\\s*\"[^\"]*\"([^>]*?>)", 
            "$1$2");
        
        // 如果pPr标签因此变为空标签，则简化它
        xmlContent = xmlContent.replaceAll(
            "<w:pPr\\s*>\\s*</w:pPr>", 
            "<w:pPr/>");
            
        // 处理自闭合的包含snapToGrid的pPr标签
        xmlContent = xmlContent.replaceAll(
            "<w:pPr\\s+[^>]*w:snapToGrid\\s*=\\s*\"[^\"]*\"[^>]*/>", 
            "<w:pPr/>");
            
        // 处理文档网格中的snapToGrid设置
        xmlContent = xmlContent.replaceAll(
            "<w:docGrid\\s+[^>]*w:snapToGrid\\s*=\\s*\"[^\"]*\"[^>]*/?>", 
            "<w:docGrid/>");
            
        // 处理独立的docGrid标签
        xmlContent = xmlContent.replaceAll(
            "<w:docGrid\\s*/>", 
            "");
            
        // 移除空的docGrid标签
        xmlContent = xmlContent.replaceAll(
            "<w:docGrid\\s*>\\s*</w:docGrid>", 
            "");
        
        int afterRemoval = xmlContent.length();
        System.out.println("📐 移除对齐到网络设置: " + (beforeRemoval - afterRemoval) + " 字符变化");
        
        return xmlContent;
    }
    
    /**
     * 移除文档网格设置
     * 
     * @param doc Word文档
     */
    private void removeDocumentGridSettings(WordprocessingMLPackage doc) {
        try {
            System.out.println("📐 开始移除文档网格设置");
            
            // 获取文档的body元素
            Document wmlDocument = doc.getMainDocumentPart().getJaxbElement();
            if (wmlDocument != null && wmlDocument.getBody() != null) {
                SectPr sectPr = wmlDocument.getBody().getSectPr();
                if (sectPr != null) {
                    // 移除文档网格设置
                    sectPr.setDocGrid(null);
                    System.out.println("✅ 文档网格设置已移除");
                } else {
                    System.out.println("⚠️ 未找到节属性设置");
                }
            } else {
                System.out.println("⚠️ 未找到文档主体");
            }
        } catch (Exception e) {
            System.err.println("⚠️ 移除文档网格设置时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 在合并前处理表格中的默认字体大小
     * 只有当单元格中没有<w:sz>或<w:szCs>时才添加默认字体大小
     */
    private void applyDefaultFontSizesBeforeMerge(WordprocessingMLPackage doc, Map<String, String> formatProperties, String docPrefix) {
        try {
            System.out.println("📏 开始为" + docPrefix + "应用默认字体大小...");
            
            // 获取文档的XML内容
            String xmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);
            
            // 查找默认字体大小
            String defaultSize = formatProperties.get(docPrefix + "_default_sz");
            String defaultStyleSize = formatProperties.get(docPrefix + "_default_style_sz");
            String defaultStyleSizeCs = formatProperties.get(docPrefix + "_default_style_szCs");
            
            String effectiveSize = defaultSize != null ? defaultSize : defaultStyleSize;
            String effectiveSizeCs = defaultStyleSizeCs; // 只有在使用默认样式时才有
            
            System.out.println("📏 " + docPrefix + "默认字体大小: " + effectiveSize + 
                (effectiveSizeCs != null ? " (szCs: " + effectiveSizeCs + ")" : ""));
            
            if (effectiveSize == null) {
                System.out.println("⚠️ " + docPrefix + "没有找到默认字体大小，跳过处理");
                return;
            }
            
            // 处理表格单元格中的<w:r>元素，在<w:rPr>中添加字体大小
            Pattern rPattern = Pattern.compile("(<w:r[^>]*>\\s*<w:rPr[^>]*>)(.*?)(</w:rPr>)", Pattern.DOTALL);
            Matcher matcher = rPattern.matcher(xmlContent);
            
            StringBuffer sb = new StringBuffer();
            
            while (matcher.find()) {
                String rStart = matcher.group(1);
                String rPrContent = matcher.group(2);
                String rPrEnd = matcher.group(3);
                
                // 只有在<w:rPr>中没有<w:sz>和<w:szCs>时才添加默认字体大小
                if (!rPrContent.contains("<w:sz ")) {
                    // 构建字体大小定义
                    StringBuilder fontSizeDefinition = new StringBuilder();
                    fontSizeDefinition.append("<w:sz w:val=\"").append(effectiveSize).append("\"/>");
                    
                    if (effectiveSizeCs != null && !rPrContent.contains("<w:szCs ")) {
                        fontSizeDefinition.append("<w:szCs w:val=\"").append(effectiveSizeCs).append("\"/>");
                    }
                    
                    // 在<w:rPr>中插入字体大小定义
                    String modifiedRPrContent = rPrContent + fontSizeDefinition.toString();
                    matcher.appendReplacement(sb, rStart + modifiedRPrContent + rPrEnd);
                    System.out.println("📏 为" + docPrefix + "运行元素添加默认字体大小: " + effectiveSize +
                        (effectiveSizeCs != null ? " (szCs: " + effectiveSizeCs + ")" : ""));
                } else {
                    matcher.appendReplacement(sb, matcher.group(0));
                }
            }
            
            matcher.appendTail(sb);
            String result = sb.toString();
            
            // 将更新后的内容重新设置到文档中
            Document document = (Document) XmlUtils.unmarshalString(result);
            doc.getMainDocumentPart().setJaxbElement(document);
            
            System.out.println("✅ " + docPrefix + "默认字体大小应用完成");
        } catch (Exception e) {
            System.err.println("⚠️ 为" + docPrefix + "应用默认字体大小时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
}