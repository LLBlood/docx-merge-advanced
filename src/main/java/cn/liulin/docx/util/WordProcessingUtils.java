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
        // 使用 addObject() 以触发样式/字体发现
        // 遍历文档列表，将除第一个文档外的所有文档内容追加到基础文档中
        for (int i = 1; i < docList.size(); i++) {
            // ✅ 8. 在合并前添加分节符，保持文档页面设置独立
            P sectionBreak = getSectionBreak(docList.get(i - 1).getMainDocumentPart());
            main1.addObject(sectionBreak);
            MainDocumentPart tempMain = docList.get(i).getMainDocumentPart();
            logger.info("开始合并文档内容，doc内容项数: {}", tempMain.getContent().size());
            int objectCount = 0;
            for (Object obj : tempMain.getContent()) {
                objectCount++;
                logger.debug("正在添加第 {} 个内容项: {}", objectCount, obj.getClass().getSimpleName());
                main1.addObject(obj);
            }
            logger.info("文档内容合并完成，共添加 {} 个内容项", objectCount);
        }

        // ✅ 10. 获取 doc2 的最后一个节属性（SectPr）
        SectPr lastSectPr = getPgSzSettings(docList.get(docList.size() - 1).getMainDocumentPart());

        // ✅ 11. 如果 doc2 有节结束（SectPr），则在合并后添加一个新节段落
        ObjectFactory factory = Context.getWmlObjectFactory();
        P newSection = factory.createP();

        PPr pPr = factory.createPPr();
        // 深拷贝 sectPr，避免引用共享
        assert lastSectPr != null;
        SectPr sectPrCopy = XmlUtils.deepCopy(lastSectPr);
        pPr.setSectPr(sectPrCopy);
        newSection.setPPr(pPr);

        // 使用 addObject() 添加，触发样式/字体等处理
        main1.addObject(newSection);

        // 修复对齐元素，确保符合Open XML规范（不处理表格边框）
        logger.info("开始修复对齐元素...");
        fixJustificationElements(docList.get(0));
        logger.info("对齐元素修复完成");
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
     * 在第一个文档末尾添加分节符，确保第二个文档保持其原始页面设置
     */
    private static P getSectionBreak(MainDocumentPart documentPart) {
        try {
            ObjectFactory factory = Context.getWmlObjectFactory();
            P sectionBreakParagraph = factory.createP();
            PPr pPr = factory.createPPr();

            // 创建分节符
            SectPr sectPr = factory.createSectPr();

            // 设置分节符类型为下一页（NEXT_PAGE）
            // 这样可以确保第二个文档从新的一页开始，并保持其原始页面设置
            SectPr.Type sectType = factory.createSectPrType();
            // 下一页分节符
            sectType.setVal("nextPage");
            sectPr.setType(sectType);

            // 保留第一个文档的页面设置
            SectPr firstDocSectPr = getPgSzSettings(documentPart);
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

            return sectionBreakParagraph;
        } catch (Exception e) {
            logger.error("添加分节符时出错: {}", e.getMessage(), e);
        }
        return null;
    }

    /**
     * 获取文档的页面设置（页面大小和方向）
     */
    private static SectPr getPgSzSettings(MainDocumentPart part) {
        // 获取文档的body元素
        Document wmlDocument = part.getJaxbElement();
        if (wmlDocument != null && wmlDocument.getBody() != null) {
            return wmlDocument.getBody().getSectPr();
        }
        return null;
    }
}