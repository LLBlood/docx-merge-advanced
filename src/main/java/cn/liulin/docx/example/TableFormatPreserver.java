package cn.liulin.docx.example;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.XmlUtils;

import java.io.File;
import java.util.Map;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 文档格式保持器
 * 用于在文档合并过程中保持各种格式的一致性
 * 
 * @author liulin
 * @version 1.0
 */
public class TableFormatPreserver {

    /**
     * 在文档合并前保存两个文档的格式信息
     * 
     * @param doc1 doc1文档
     * @param doc2 doc2文档
     * @return 包含两个文档格式信息的映射
     */
    public static Map<String, String> saveDocumentFormat(WordprocessingMLPackage doc1, WordprocessingMLPackage doc2) {
        Map<String, String> formatProperties = new HashMap<>();
        
        try {
            System.out.println("🔍 开始保存两个文档的格式信息...");
            
            // 直接使用docx4j API获取XML内容
            String doc1XmlContent = XmlUtils.marshaltoString(doc1.getMainDocumentPart().getJaxbElement(), true, true);
            System.out.println("🔍 开始保存doc1格式信息，XML长度: " + doc1XmlContent.length());
            
            // 直接使用docx4j API获取XML内容
            String doc2XmlContent = XmlUtils.marshaltoString(doc2.getMainDocumentPart().getJaxbElement(), true, true);
            System.out.println("🔍 开始保存doc2格式信息，XML长度: " + doc2XmlContent.length());
            
            // 获取样式XML内容
            String doc1StyleXmlContent = "";
            String doc2StyleXmlContent = "";
            
            StyleDefinitionsPart stylePart1 = doc1.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylePart1 != null) {
                doc1StyleXmlContent = XmlUtils.marshaltoString(stylePart1.getJaxbElement(), true, true);
                System.out.println("🎨 doc1样式XML内容长度: " + doc1StyleXmlContent.length());
            }
            
            StyleDefinitionsPart stylePart2 = doc2.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylePart2 != null) {
                doc2StyleXmlContent = XmlUtils.marshaltoString(stylePart2.getJaxbElement(), true, true);
                System.out.println("🎨 doc2样式XML内容长度: " + doc2StyleXmlContent.length());
            }
            
            // 保存doc1的所有trHeight元素属性
            Pattern trHeightPattern = Pattern.compile("<w:trHeight\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
            Matcher matcher = trHeightPattern.matcher(doc1XmlContent);
            
            int doc1Index = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String heightValue = matcher.group(2);
                formatProperties.put("doc1_trHeight_" + doc1Index, heightValue);
                System.out.println("📊 保存doc1表格行高[" + doc1Index + "]: " + heightValue);
                doc1Index++;
            }
            
            System.out.println("✅ doc1表格行高信息保存完成，共保存 " + doc1Index + " 个行高设置");
            
            // 提取doc2所有表格行高的信息
            matcher = trHeightPattern.matcher(doc2XmlContent);
            
            int doc2Index = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String heightValue = matcher.group(2);
                formatProperties.put("doc2_trHeight_" + doc2Index, heightValue);
                System.out.println("📊 保存doc2表格行高[" + doc2Index + "]: " + heightValue);
                doc2Index++;
            }
            
            System.out.println("✅ doc2表格行高信息保存完成，共保存 " + doc2Index + " 个行高设置");
            
            // 保存doc1的所有tbl元素属性
            Pattern tblPattern = Pattern.compile("<w:tbl(?:\\s[^>]*)?>(.*?)</w:tbl>", Pattern.DOTALL);
            matcher = tblPattern.matcher(doc1XmlContent);
            
            int doc1TblIndex = 0;
            while (matcher.find()) {
                String tblContent = matcher.group(0); // 包括<w:tbl>标签本身
                formatProperties.put("doc1_tbl_" + doc1TblIndex, tblContent);
                System.out.println("📋 保存doc1表格[" + doc1TblIndex + "]，长度: " + tblContent.length());
                doc1TblIndex++;
            }
            
            System.out.println("✅ doc1表格属性信息保存完成，共保存 " + doc1TblIndex + " 个表格");
            
            // 保存doc2的所有tbl元素属性
            matcher = tblPattern.matcher(doc2XmlContent);
            
            int doc2TblIndex = 0;
            while (matcher.find()) {
                String tblContent = matcher.group(0); // 包括<w:tbl>标签本身
                formatProperties.put("doc2_tbl_" + doc2TblIndex, tblContent);
                System.out.println("📋 保存doc2表格[" + doc2TblIndex + "]，长度: " + tblContent.length());
                doc2TblIndex++;
            }
            
            System.out.println("✅ doc2表格属性信息保存完成，共保存 " + doc2TblIndex + " 个表格");
            
            // 保存doc1的字体信息（从样式中获取默认字体）
            saveDefaultStyleInfo(doc1StyleXmlContent, "doc1", formatProperties);
            
            // 保存doc2的字体信息（从样式中获取默认字体）
            saveDefaultStyleInfo(doc2StyleXmlContent, "doc2", formatProperties);
            
            // 保存doc1的段落缩进信息（特别是表格内的段落）
            Pattern indentPattern = Pattern.compile("<w:ind\\s+([^>]+w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
            matcher = indentPattern.matcher(doc1XmlContent);
            
            int doc1IndIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String indValue = matcher.group(2);
                formatProperties.put("doc1_ind_" + doc1IndIndex, indValue);
                System.out.println("-indent- 保存doc1段落缩进[" + doc1IndIndex + "]: " + indValue);
                doc1IndIndex++;
            }
            
            System.out.println("✅ doc1段落缩进信息保存完成，共保存 " + doc1IndIndex + " 个缩进设置");
            
            // 保存doc2的段落缩进信息（特别是表格内的段落）
            matcher = indentPattern.matcher(doc2XmlContent);
            
            int doc2IndIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String indValue = matcher.group(2);
                formatProperties.put("doc2_ind_" + doc2IndIndex, indValue);
                System.out.println("-indent- 保存doc2段落缩进[" + doc2IndIndex + "]: " + indValue);
                doc2IndIndex++;
            }
            
            System.out.println("✅ doc2段落缩进信息保存完成，共保存 " + doc2IndIndex + " 个缩进设置");
            
            System.out.println("💾 格式信息保存完成，总共保存了 " + formatProperties.size() + " 个格式属性");
            
        } catch (Exception e) {
            System.err.println("⚠️ 保存文档格式信息时出错: " + e.getMessage());
            e.printStackTrace();
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
            Pattern stylePattern = Pattern.compile(
                "<w:style[^>]*w:type=\"paragraph\"[^>]*>.*?<w:name\\s+w:val=\"Normal\"\\s*/>.*?</w:style>", 
                Pattern.DOTALL);
            Matcher styleMatcher = stylePattern.matcher(xmlContent);
            
            if (styleMatcher.find()) {
                String styleContent = styleMatcher.group(0);
                
                // 提取字体主题信息
                Pattern fontPattern = Pattern.compile(
                    "<w:rFonts\\s+([^>]*w:asciiTheme\\s*=\\s*\"([^\"]+)\"[^>]*w:hAnsiTheme\\s*=\\s*\"([^\"]+)\"[^>]*w:eastAsiaTheme\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
                Matcher fontMatcher = fontPattern.matcher(styleContent);
                
                if (fontMatcher.find()) {
                    String fullAttrs = fontMatcher.group(1);
                    String asciiTheme = fontMatcher.group(2);
                    String hAnsiTheme = fontMatcher.group(3);
                    String eastAsiaTheme = fontMatcher.group(4);
                    
                    formatProperties.put(docPrefix + "_default_style_font_asciiTheme", asciiTheme);
                    formatProperties.put(docPrefix + "_default_style_font_hAnsiTheme", hAnsiTheme);
                    formatProperties.put(docPrefix + "_default_style_font_eastAsiaTheme", eastAsiaTheme);
                    
                    System.out.println("🔤 保存" + docPrefix + "默认样式(Normal)字体主题: asciiTheme=" + asciiTheme + 
                        ", hAnsiTheme=" + hAnsiTheme + ", eastAsiaTheme=" + eastAsiaTheme);
                }
                
                // 提取字体大小信息
                Pattern sizePattern = Pattern.compile("<w:sz\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
                Matcher sizeMatcher = sizePattern.matcher(styleContent);
                
                if (sizeMatcher.find()) {
                    String fullAttrs = sizeMatcher.group(1);
                    String szValue = sizeMatcher.group(2);
                    formatProperties.put(docPrefix + "_default_style_sz", szValue);
                    System.out.println("📏 保存" + docPrefix + "默认样式(Normal)字体大小: " + szValue);
                }
                
                // 提取复杂字体大小信息
                Pattern sizeCsPattern = Pattern.compile("<w:szCs\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
                Matcher sizeCsMatcher = sizeCsPattern.matcher(styleContent);
                
                if (sizeCsMatcher.find()) {
                    String fullAttrs = sizeCsMatcher.group(1);
                    String szCsValue = sizeCsMatcher.group(2);
                    formatProperties.put(docPrefix + "_default_style_szCs", szCsValue);
                    System.out.println("📏 保存" + docPrefix + "默认样式(Normal)复杂字体大小: " + szCsValue);
                }
            }
        } catch (Exception e) {
            System.err.println("⚠️ 保存默认样式信息时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 保存两个文档的样式信息
     * 
     * @param doc1 doc1文档
     * @param doc2 doc2文档
     * @param formatProperties 格式信息存储映射
     */
    private static void saveStyleInformation(WordprocessingMLPackage doc1, WordprocessingMLPackage doc2, Map<String, String> formatProperties) {
        try {
            // 保存doc1的样式信息
            StyleDefinitionsPart stylePart1 = doc1.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylePart1 != null) {
                String style1Xml = XmlUtils.marshaltoString(stylePart1.getJaxbElement(), true, true);
                formatProperties.put("doc1_styles", style1Xml);
                System.out.println("🎨 保存doc1样式信息，XML长度: " + style1Xml.length());
            }
            
            // 保存doc2的样式信息
            StyleDefinitionsPart stylePart2 = doc2.getMainDocumentPart().getStyleDefinitionsPart();
            if (stylePart2 != null) {
                String style2Xml = XmlUtils.marshaltoString(stylePart2.getJaxbElement(), true, true);
                formatProperties.put("doc2_styles", style2Xml);
                System.out.println("🎨 保存doc2样式信息，XML长度: " + style2Xml.length());
            }
        } catch (Exception e) {
            System.err.println("⚠️ 保存样式信息时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
}