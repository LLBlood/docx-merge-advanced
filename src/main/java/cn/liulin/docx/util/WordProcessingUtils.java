package cn.liulin.docx.util;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/15 15:55
 */
public class WordProcessingUtils {
    private static final Logger logger = LoggerFactory.getLogger(WordProcessingUtils.class);

    /**
     * 根据处理后的文档路径列表加载Word文档
     * 该方法会遍历路径列表，将每个路径对应的文档加载为WordprocessingMLPackage对象
     * 
     * @param processedDocPathList 包含处理后文档路径的列表
     * @return 包含WordprocessingMLPackage对象的列表
     * @throws Docx4JException 如果加载文档过程中发生错误
     */
    public static List<WordprocessingMLPackage> loadDocList(List<String> processedDocPathList) throws Docx4JException {
        List<WordprocessingMLPackage> docList = new ArrayList<>();
        // 遍历处理后的文档路径列表，加载每个文档
        for (String processedDocPath : processedDocPathList) {
            WordprocessingMLPackage load = WordprocessingMLPackage.load(new File(processedDocPath));
            docList.add(load);
        }
        return docList;
    }

    /**
     * 移除文档列表中所有文档的网格设置
     * 该方法会遍历文档列表，为每个文档调用移除网格设置的方法
     * 
     * @param docList 包含WordprocessingMLPackage对象的文档列表
     */
    public static void removeDocumentGridSettingsList(List<WordprocessingMLPackage> docList) {
        // 遍历文档列表，为每个文档移除网格设置
        for (WordprocessingMLPackage doc : docList) {
            removeDocumentGridSettings(doc);
        }
    }

    /**
     * 移除文档网格设置
     *
     * @param doc Word文档
     */
    private static void removeDocumentGridSettings(WordprocessingMLPackage doc) {
        try {
            logger.info("开始移除文档网格设置");

            // 获取文档的body元素
            Document wmlDocument = doc.getMainDocumentPart().getJaxbElement();
            if (wmlDocument != null && wmlDocument.getBody() != null) {
                SectPr sectPr = wmlDocument.getBody().getSectPr();
                if (sectPr != null) {
                    // 移除文档网格设置
                    sectPr.setDocGrid(null);
                    logger.info("文档网格设置已移除");
                } else {
                    logger.warn("未找到节属性设置");
                }
            } else {
                logger.warn("未找到文档主体");
            }
        } catch (Exception e) {
            logger.error("移除文档网格设置时出错: {}", e.getMessage(), e);
        }
    }

    /**
     * 将文档列表中的内容添加到基础文档中
     * 该方法会遍历文档列表，将除第一个文档外的所有文档内容追加到第一个文档中，
     * 并在文档之间添加分节符以保持各自的页面设置
     * 
     * @param main1 基础文档的主要部分，其他文档的内容将被添加到此文档中
     * @param docList 包含WordprocessingMLPackage对象的文档列表
     */
    public static void addDocListToBase(MainDocumentPart main1, List<WordprocessingMLPackage> docList) {
        // 如果是第一个word，则获取word的body的SectPr属性
        // 删除第一个word的body的SectPr属性
        // 获取第一个word最后一个content的内容，将内容的分节属性设置为body的SectPr属性
        // 如果不是最后一个word，则获取word的body的SectPr属性
        // 获取word的最后一个content的内容，将内容的分节属性设置为body的SectPr属性
        // 如果是最后一个word，则获取word的body的SectPr属性
        // 将第一个word的body设置为最后一个word的body的SectPr属性
        for (int i = 0; i < docList.size(); i++) {
            WordprocessingMLPackage wordprocessingMLPackage = docList.get(i);
            MainDocumentPart mainDocumentPart = wordprocessingMLPackage.getMainDocumentPart();
            if (i == 0) {
                // 保存第一个文档的节设置
                SectPr firstDocSectPr = getPgSzSettings(mainDocumentPart);
                // 删除第一个word的body的SectPr属性
                Document wmlDocument = mainDocumentPart.getJaxbElement();
                if (wmlDocument != null && wmlDocument.getBody() != null) {
                    wmlDocument.getBody().setSectPr(null);
                    logger.debug("已移除文档的节属性设置");
                }
                // 获取第一个word最后一个content的内容，将内容的分节属性设置为body的SectPr属性
                List<Object> content = wmlDocument.getBody().getContent();
                Object o = content.get(content.size() - 1);
                if (o instanceof P) {
                    P p = (P) o;
                    PPr pPr = p.getPPr();
                    if (pPr == null) {
                        p.setPPr(createSectionPPr(firstDocSectPr));
                    } else {
                        p.getPPr().setSectPr(firstDocSectPr);
                    }
                } else {
                    P sectionParagraph = createSectionParagraph(firstDocSectPr);
                    content.add(sectionParagraph);
                }
            } else if (i == docList.size() - 1) {
                // 如果是最后一个word，则获取word的body的SectPr属性
                SectPr docSectPr = getPgSzSettings(mainDocumentPart);
                // 将第一个word的body设置为最后一个word的body的SectPr属性
                Body body = main1.getJaxbElement().getBody();
                body.setSectPr(docSectPr);
                for (Object o : mainDocumentPart.getJaxbElement().getContent()) {
                    main1.addObject(o);
                }
            } else {
                // 如果不是最后一个word，则获取word的body的SectPr属性
                SectPr docSectPr = getPgSzSettings(mainDocumentPart);
                // 获取word的最后一个content的内容，将内容的分节属性设置为body的SectPr属性
                List<Object> content = mainDocumentPart.getJaxbElement().getBody().getContent();
                Object lastContent = content.get(content.size() - 1);
                if (lastContent instanceof P) {
                    P p = (P) lastContent;
                    PPr pPr = p.getPPr();
                    if (pPr == null) {
                        p.setPPr(createSectionPPr(docSectPr));
                    } else {
                        p.getPPr().setSectPr(docSectPr);
                    }
                } else {
                    P sectionParagraph = createSectionParagraph(docSectPr);
                    content.add(sectionParagraph);
                }
                for (Object co : content) {
                    main1.addObject(co);
                }
            }
        }
        // 修复对齐元素，确保符合Open XML规范（不处理表格边框）
        logger.info("开始修复对齐元素...");
        fixJustificationElements(docList.get(0));
        logger.info("对齐元素修复完成");
    }

    /**
     * 创建带有指定节设置的段落
     * @param sectPr 节属性设置
     * @return 包含节设置的段落
     */
    private static P createSectionParagraph(SectPr sectPr) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P sectionParagraph = factory.createP();
        PPr pPr = factory.createPPr();
        
        // 深拷贝节属性，避免引用共享
        SectPr sectPrCopy = XmlUtils.deepCopy(sectPr);
        pPr.setSectPr(sectPrCopy);
        sectionParagraph.setPPr(pPr);
        
        logger.debug("创建了带有节设置的段落");
        return sectionParagraph;
    }

    /**
     * 创建带有指定节设置的段落
     * @param sectPr 节属性设置
     * @return 包含节设置的段落
     */
    private static PPr createSectionPPr(SectPr sectPr) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        PPr pPr = factory.createPPr();

        // 深拷贝节属性，避免引用共享
        SectPr sectPrCopy = XmlUtils.deepCopy(sectPr);
        pPr.setSectPr(sectPrCopy);

        logger.debug("创建了带有节设置的段落PPr");
        return pPr;
    }

    /**
     * 修复对齐元素，确保所有 jc 元素都有 val 属性
     */
    private static void fixJustificationElements(WordprocessingMLPackage doc) {
        try {
            // 直接使用docx4j API获取XML内容，不再需要通过ZIP方式读取
            String xmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);
            logger.debug("docx4j读取的主文档XML内容长度: {}", xmlContent.length());

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
                logger.debug("原始样式XML内容长度: {}", styleXmlContent.length());

                // 使用replace方法修复所有缺失val属性的jc标签
                styleXmlContent = fixMissingValAttributes(styleXmlContent);

                // 将更新后的XML内容重新设置到样式部分中
                Styles styles = (Styles) XmlUtils.unmarshalString(styleXmlContent);
                stylePart.setJaxbElement(styles);
            }

            logger.info("对齐元素修复完成");
        } catch (Exception e) {
            logger.error("修复对齐元素时出错: {}", e.getMessage(), e);
        }
    }

    /**
     * 修复XML中缺失val属性的jc元素
     */
    private static String fixMissingValAttributes(String xmlContent) {
        logger.debug("开始修复缺失val属性的对齐元素");

        // 使用replace方法修复所有缺失val属性的jc标签
        int beforeFix = xmlContent.length();
        xmlContent = xmlContent.replace("<w:jc/>", "<w:jc w:val=\"left\"/>");
        int afterFix = xmlContent.length();
        logger.debug("修复缺失val属性的jc标签: {} 字符变化", (afterFix - beforeFix));

        return xmlContent;
    }

    /**
     * 获取文档的页面设置（页面大小和方向）
     */
    private static SectPr getPgSzSettings(MainDocumentPart part) {
        // 获取文档的body元素
        Document wmlDocument = part.getJaxbElement();
        if (wmlDocument != null && wmlDocument.getBody() != null) {
            SectPr sectPr = wmlDocument.getBody().getSectPr();
            // 深拷贝节属性，避免引用共享
            if (sectPr != null) {
                return XmlUtils.deepCopy(sectPr);
            }
        }
        return null;
    }
}