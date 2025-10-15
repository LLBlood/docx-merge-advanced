package cn.liulin.docx.util;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/15 15:48
 */
public class FontMapperUtil {
    private static final Logger logger = LoggerFactory.getLogger(FontMapperUtil.class);

    /**
     * 为文档列表中的每个文档应用默认字体大小，确保在合并前所有文档具有一致的字体格式
     * 该方法会遍历文档列表，为每个文档调用字体大小应用方法
     * 
     * @param docList 包含WordprocessingMLPackage对象的文档列表
     * @param formatProperties 包含格式属性的映射，用于获取每个文档的默认字体大小信息
     */
    public static void applyDocListDefaultFontSizesBeforeMerge(List<WordprocessingMLPackage> docList, Map<String, String> formatProperties) {
        logger.info("开始在合并前应用默认字体大小...");
        // 遍历文档列表，为每个文档应用默认字体大小
        for (int i = 0; i < docList.size(); i++) {
            applyDefaultFontSizesBeforeMerge(docList.get(i), formatProperties, "doc" + (i + 1));
        }
        logger.info("合并前默认字体大小应用完成");
    }

    /**
     * 在合并前处理表格中的默认字体大小
     * 只有当单元格中没有<w:sz>或<w:szCs>时才添加默认字体大小
     */
    private static void applyDefaultFontSizesBeforeMerge(WordprocessingMLPackage doc, Map<String, String> formatProperties, String docPrefix) {
        try {
            logger.info("开始为{}应用默认字体大小...", docPrefix);

            // 获取文档的XML内容
            String xmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);

            // 查找默认字体大小
            String defaultSize = formatProperties.get(docPrefix + "_default_sz");
            String defaultStyleSize = formatProperties.get(docPrefix + "_default_style_sz");
            String defaultStyleSizeCs = formatProperties.get(docPrefix + "_default_style_szCs");

            String effectiveSize = defaultSize != null ? defaultSize : defaultStyleSize;
            // 只有在使用默认样式时才有

            logger.debug("{}默认字体大小: {} {}", docPrefix, effectiveSize,
                    (defaultStyleSizeCs != null ? " (szCs: " + defaultStyleSizeCs + ")" : ""));

            if (effectiveSize == null) {
                logger.warn("{}没有找到默认字体大小，跳过处理", docPrefix);
                return;
            }

            // 处理表格单元格中的<w:r>元素，在<w:rPr>中添加字体大小
            Matcher matcher = PatternConst.R_PATTERN.matcher(xmlContent);

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

                    if (defaultStyleSizeCs != null && !rPrContent.contains("<w:szCs ")) {
                        fontSizeDefinition.append("<w:szCs w:val=\"").append(defaultStyleSizeCs).append("\"/>");
                    }

                    // 在<w:rPr>中插入字体大小定义
                    String modifiedRprContent = rPrContent + fontSizeDefinition;
                    matcher.appendReplacement(sb, rStart + modifiedRprContent + rPrEnd);
                    logger.debug("为{}运行元素添加默认字体大小: {} {}", docPrefix, effectiveSize,
                            (defaultStyleSizeCs != null ? " (szCs: " + defaultStyleSizeCs + ")" : ""));
                } else {
                    matcher.appendReplacement(sb, matcher.group(0));
                }
            }

            matcher.appendTail(sb);
            String result = sb.toString();

            // 将更新后的内容重新设置到文档中
            Document document = (Document) XmlUtils.unmarshalString(result);
            doc.getMainDocumentPart().setJaxbElement(document);

            logger.info("{}默认字体大小应用完成", docPrefix);
        } catch (Exception e) {
            logger.error("为{}应用默认字体大小时出错: {}", docPrefix, e.getMessage(), e);
        }
    }
}