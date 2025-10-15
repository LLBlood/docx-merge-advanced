package cn.liulin.docx.util;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.XmlUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.regex.Matcher;

/**
 * 文档格式保持器
 * 用于在文档合并过程中保持各种格式的一致性
 * 
 * @author liulin
 * @version 1.0
 */
public class TableFormatPreserverUtil {
    private static final Logger logger = LoggerFactory.getLogger(TableFormatPreserverUtil.class);

    /**
     * 在文档合并前保存两个文档的格式信息
     *
     * @param docPath doc1文档
     * @return 包含两个文档格式信息的映射
     */
    public static Map<String, String> saveDocumentFormat(List<WordprocessingMLPackage> docPath) {
        Map<String, String> formatProperties = new HashMap<>();

        try {
            logger.info("开始保存文档的格式信息...");
            for (int i = 0; i < docPath.size(); i++) {
                // 直接使用docx4j API获取XML内容
                WordprocessingMLPackage doc = docPath.get(i);
                String docXmlContent = XmlUtils.marshaltoString(doc.getMainDocumentPart().getJaxbElement(), true, true);
                logger.debug("开始保存doc格式信息，XML长度: {}", docXmlContent.length());
                // 获取样式XML内容
                String docStyleXmlContent = "";
                StyleDefinitionsPart stylePart = doc.getMainDocumentPart().getStyleDefinitionsPart();
                if (stylePart != null) {
                    docStyleXmlContent = XmlUtils.marshaltoString(stylePart.getJaxbElement(), true, true);
                    logger.debug("doc1样式XML内容长度: {}", docStyleXmlContent.length());
                }

                // 保存doc的所有trHeight元素属性
                Matcher matcher = PatternConst.TR_HEIGHT_PATTERN.matcher(docXmlContent);
                int docIndex = 0;
                while (matcher.find()) {
                    String heightValue = matcher.group(2);
                    formatProperties.put("doc" + (i + 1) + "_trHeight_" + docIndex, heightValue);
                    logger.debug("保存doc表格行高[{}]: {}", docIndex, heightValue);
                    docIndex++;
                }

                logger.info("doc表格行高信息保存完成，共保存 {} 个行高设置", docIndex);

                // 保存doc的所有tbl元素属性
                matcher = PatternConst.TBL_PATTERN.matcher(docXmlContent);
                int docTblIndex = 0;
                while (matcher.find()) {
                    // 包括<w:tbl>标签本身
                    String tblContent = matcher.group(0);
                    formatProperties.put("doc" + (i + 1) + "_tbl_" + docTblIndex, tblContent);
                    logger.debug("保存doc表格[{}]，长度: {}", docTblIndex, tblContent.length());
                    docTblIndex++;
                }
                logger.info("doc表格属性信息保存完成，共保存 {} 个表格", docTblIndex);

                // 保存doc的字体信息（从样式中获取默认字体）
                saveDefaultStyleInfo(docStyleXmlContent, "doc" + (i + 1) , formatProperties);

                // 保存doc的段落缩进信息（特别是表格内的段落）
                matcher = PatternConst.INDENT_PATTERN.matcher(docXmlContent);

                int docIndIndex = 0;
                while (matcher.find()) {
                    String indValue = matcher.group(2);
                    formatProperties.put("doc" + (i + 1) + "_ind_" + docIndIndex, indValue);
                    logger.debug("保存doc段落缩进[{}]: {}", docIndIndex, indValue);
                    docIndIndex++;
                }

                logger.debug("doc段落缩进信息保存完成，共保存 {} 个缩进设置", docIndIndex);
            }

            logger.info("格式信息保存完成，总共保存了 {} 个格式属性", formatProperties.size());

        } catch (Exception e) {
            logger.error("保存文档格式信息时出错: {}", e.getMessage(), e);
        }

        return formatProperties;
    }
    
    /**
     * 保存默认样式(Normal样式)的字体和字体大小信息
     * 
     * @param xmlContent XML内容
     * @param docPrefix 文档前缀
     * @param formatProperties 格式属性映射
     */
    private static void saveDefaultStyleInfo(String xmlContent, String docPrefix, Map<String, String> formatProperties) {
        try {
            // 查找默认段落样式(Normal样式)
            Matcher styleMatcher = PatternConst.STYLE_PATTERN.matcher(xmlContent);
            
            if (styleMatcher.find()) {
                String styleContent = styleMatcher.group(0);
                
                // 提取字体主题信息
                Matcher fontMatcher = PatternConst.FONT_PATTERN.matcher(styleContent);
                
                if (fontMatcher.find()) {
                    String asciiTheme = fontMatcher.group(2);
                    String hAnsiTheme = fontMatcher.group(3);
                    String eastAsiaTheme = fontMatcher.group(4);
                    
                    formatProperties.put(docPrefix + "_default_style_font_asciiTheme", asciiTheme);
                    formatProperties.put(docPrefix + "_default_style_font_hAnsiTheme", hAnsiTheme);
                    formatProperties.put(docPrefix + "_default_style_font_eastAsiaTheme", eastAsiaTheme);
                    
                    logger.debug("保存{}默认样式(Normal)字体主题: asciiTheme={}, hAnsiTheme={}, eastAsiaTheme={}",
                        docPrefix, asciiTheme, hAnsiTheme, eastAsiaTheme);
                }
                
                // 提取字体大小信息
                Matcher sizeMatcher = PatternConst.SIZE_PATTERN.matcher(styleContent);
                if (sizeMatcher.find()) {
                    String szValue = sizeMatcher.group(2);
                    formatProperties.put(docPrefix + "_default_style_sz", szValue);
                    logger.debug("保存{}默认样式(Normal)字体大小: {}", docPrefix, szValue);
                }
                
                // 提取复杂字体大小信息
                Matcher sizeCsMatcher = PatternConst.SIZE_CS_PATTERN.matcher(styleContent);
                
                if (sizeCsMatcher.find()) {
                    String szCsValue = sizeCsMatcher.group(2);
                    formatProperties.put(docPrefix + "_default_style_szCs", szCsValue);
                    logger.debug("保存{}默认样式(Normal)复杂字体大小: {}", docPrefix, szCsValue);
                }
            }
        } catch (Exception e) {
            logger.error("保存默认样式信息时出错: {}", e.getMessage(), e);
        }
    }
}