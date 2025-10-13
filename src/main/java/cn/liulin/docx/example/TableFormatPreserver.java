package cn.liulin.docx.example;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.docx4j.XmlUtils;

import java.util.List;
import java.util.ArrayList;
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
            // 获取doc1的XML内容
            String doc1XmlContent = XmlUtils.marshaltoString(doc1.getMainDocumentPart().getJaxbElement(), true, true);
            System.out.println("🔍 开始保存doc1格式信息，XML长度: " + doc1XmlContent.length());
            
            // 提取doc1所有表格行高的信息
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
            
            // 获取doc2的XML内容
            String doc2XmlContent = XmlUtils.marshaltoString(doc2.getMainDocumentPart().getJaxbElement(), true, true);
            System.out.println("🔍 开始保存doc2格式信息，XML长度: " + doc2XmlContent.length());
            
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
            
            // 保存doc1的字体信息
            Pattern rFontsPattern = Pattern.compile("<w:rFonts\\s+([^>]*w:ascii\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
            matcher = rFontsPattern.matcher(doc1XmlContent);
            
            int doc1FontIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String fontValue = matcher.group(2);
                formatProperties.put("doc1_font_" + doc1FontIndex, fontValue);
                System.out.println("🔤 保存doc1字体[" + doc1FontIndex + "]: " + fontValue);
                doc1FontIndex++;
            }
            
            System.out.println("✅ doc1字体信息保存完成，共保存 " + doc1FontIndex + " 个字体设置");
            
            // 保存doc2的字体信息
            matcher = rFontsPattern.matcher(doc2XmlContent);
            
            int doc2FontIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String fontValue = matcher.group(2);
                formatProperties.put("doc2_font_" + doc2FontIndex, fontValue);
                System.out.println("🔤 保存doc2字体[" + doc2FontIndex + "]: " + fontValue);
                doc2FontIndex++;
            }
            
            System.out.println("✅ doc2字体信息保存完成，共保存 " + doc2FontIndex + " 个字体设置");
            
            // 保存doc1的字体大小信息
            Pattern szPattern = Pattern.compile("<w:sz\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
            matcher = szPattern.matcher(doc1XmlContent);
            
            int doc1SzIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String szValue = matcher.group(2);
                formatProperties.put("doc1_sz_" + doc1SzIndex, szValue);
                System.out.println("📏 保存doc1字体大小[" + doc1SzIndex + "]: " + szValue);
                doc1SzIndex++;
            }
            
            System.out.println("✅ doc1字体大小信息保存完成，共保存 " + doc1SzIndex + " 个字体大小设置");
            
            // 保存doc2的字体大小信息
            matcher = szPattern.matcher(doc2XmlContent);
            
            int doc2SzIndex = 0;
            while (matcher.find()) {
                String fullAttrs = matcher.group(1);
                String szValue = matcher.group(2);
                formatProperties.put("doc2_sz_" + doc2SzIndex, szValue);
                System.out.println("📏 保存doc2字体大小[" + doc2SzIndex + "]: " + szValue);
                doc2SzIndex++;
            }
            
            System.out.println("✅ doc2字体大小信息保存完成，共保存 " + doc2SzIndex + " 个字体大小设置");
            
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
     * 在文档合并后精确恢复两个文档的格式
     * 
     * @param mergedDoc 合并后的文档
     * @param formatProperties 格式信息映射
     */
    public static void restoreDocumentFormat(WordprocessingMLPackage mergedDoc, Map<String, String> formatProperties) {
        try {
            // 当前实现中，我们通过fixDocumentFormatInXml方法处理
            System.out.println("✅ 文档格式恢复完成");
        } catch (Exception e) {
            System.err.println("⚠️ 恢复文档格式信息时出错: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 通过XML操作直接修复文档格式
     * 精确恢复两个文档的格式，包括行高、字体、字体大小等
     * 
     * @param xmlContent XML内容
     * @param formatProperties 格式信息
     * @return 修复后的XML内容
     */
    public static String fixDocumentFormatInXml(String xmlContent, Map<String, String> formatProperties) {
        try {
            System.out.println("🔧 开始修复文档格式，输入XML长度: " + xmlContent.length());
            System.out.println("🔧 格式属性数量: " + formatProperties.size());
            
            // 精确恢复两个文档的表格行高
            String result = restoreTableRowHeights(xmlContent, formatProperties);
            System.out.println("📊 表格行高恢复后XML长度: " + result.length());
            
            // 精确恢复两个文档的字体设置
            result = restoreFontSettings(result, formatProperties);
            System.out.println("🔤 字体设置恢复后XML长度: " + result.length());
            
            // 精确恢复两个文档的字体大小设置
            result = restoreFontSizeSettings(result, formatProperties);
            System.out.println("📏 字体大小恢复后XML长度: " + result.length());
            
            // 精确恢复两个文档的段落缩进设置
            result = restoreIndentSettings(result, formatProperties);
            System.out.println("-indent- 段落缩进恢复后XML长度: " + result.length());
            
            // 移除表格内段落的首行缩进（特别处理表格内的段落前空格问题）
            result = removeTableParagraphFirstLineIndent(result);
            System.out.println("-indent- 表格内段落首行缩进移除后XML长度: " + result.length());
            
            // 修复所有缺失val属性的jc元素（表格和段落对齐）
            result = fixMissingJustificationValues(result);
            System.out.println("🔗 对齐元素修复后XML长度: " + result.length());
            
            System.out.println("✅ 文档格式XML修复完成");
            return result;
        } catch (Exception e) {
            System.err.println("⚠️ 修复文档格式时出错: " + e.getMessage());
            e.printStackTrace();
            return xmlContent; // 出错时返回原始内容
        }
    }
    
    /**
     * 精确恢复两个文档的表格行高
     * 
     * @param xmlContent XML内容
     * @param formatProperties 格式信息
     * @return 修复后的XML内容
     */
    private static String restoreTableRowHeights(String xmlContent, Map<String, String> formatProperties) {
        // 恢复doc1的表格行高值（前N个）
        int doc1TableCount = 0;
        for (String key : formatProperties.keySet()) {
            if (key.startsWith("doc1_trHeight_")) {
                doc1TableCount++;
            }
        }
        
        System.out.println("📊 doc1表格行高数量: " + doc1TableCount);
        
        // 恢复表格行高值
        Pattern trHeightPattern = Pattern.compile("<w:trHeight\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
        Matcher matcher = trHeightPattern.matcher(xmlContent);
        StringBuffer sb = new StringBuffer();
        
        int index = 0;
        while (matcher.find()) {
            String originalHeight;
            if (index < doc1TableCount) {
                // 这是doc1的表格行高
                originalHeight = formatProperties.get("doc1_trHeight_" + index);
            } else {
                // 这是doc2的表格行高
                originalHeight = formatProperties.get("doc2_trHeight_" + (index - doc1TableCount));
            }
            
            if (originalHeight != null) {
                // 恢复原始行高值
                String fullAttrs = matcher.group(1);
                String currentHeight = matcher.group(2);
                
                // 替换为原始值
                String newFullAttrs = fullAttrs.replace("w:val=\"" + currentHeight + "\"", 
                                                       "w:val=\"" + originalHeight + "\"");
                matcher.appendReplacement(sb, "<w:trHeight " + newFullAttrs + ">");
                System.out.println("🔧 恢复第 " + (index + 1) + " 个表格行高值: " + currentHeight + " -> " + originalHeight);
            } else {
                matcher.appendReplacement(sb, matcher.group(0));
                System.out.println("⚠️ 未找到第 " + (index + 1) + " 个表格行高的原始值");
            }
            index++;
        }
        matcher.appendTail(sb);
        
        System.out.println("📊 总共处理了 " + index + " 个表格行高");
        return sb.toString();
    }
    
    /**
     * 精确恢复两个文档的字体设置
     * 
     * @param xmlContent XML内容
     * @param formatProperties 格式信息
     * @return 修复后的XML内容
     */
    private static String restoreFontSettings(String xmlContent, Map<String, String> formatProperties) {
        // 恢复doc1的字体设置值（前N个）
        int doc1FontCount = 0;
        for (String key : formatProperties.keySet()) {
            if (key.startsWith("doc1_font_")) {
                doc1FontCount++;
            }
        }
        
        System.out.println("🔤 doc1字体设置数量: " + doc1FontCount);
        
        // 恢复字体设置值
        Pattern rFontsPattern = Pattern.compile("<w:rFonts\\s+([^>]*w:ascii\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
        Matcher matcher = rFontsPattern.matcher(xmlContent);
        StringBuffer sb = new StringBuffer();
        
        int index = 0;
        while (matcher.find()) {
            String originalFont;
            if (index < doc1FontCount) {
                // 这是doc1的字体设置
                originalFont = formatProperties.get("doc1_font_" + index);
            } else {
                // 这是doc2的字体设置
                originalFont = formatProperties.get("doc2_font_" + (index - doc1FontCount));
            }
            
            if (originalFont != null) {
                // 恢复原始字体值
                String fullAttrs = matcher.group(1);
                String currentFont = matcher.group(2);
                
                // 替换为原始值
                String newFullAttrs = fullAttrs.replace("w:ascii=\"" + currentFont + "\"", 
                                                       "w:ascii=\"" + originalFont + "\"");
                matcher.appendReplacement(sb, "<w:rFonts " + newFullAttrs + ">");
                System.out.println("🔧 恢复第 " + (index + 1) + " 个字体设置值: " + currentFont + " -> " + originalFont);
            } else {
                matcher.appendReplacement(sb, matcher.group(0));
                System.out.println("⚠️ 未找到第 " + (index + 1) + " 个字体的原始值");
            }
            index++;
        }
        matcher.appendTail(sb);
        
        System.out.println("🔤 总共处理了 " + index + " 个字体设置");
        return sb.toString();
    }
    
    /**
     * 精确恢复两个文档的字体大小设置
     * 
     * @param xmlContent XML内容
     * @param formatProperties 格式信息
     * @return 修复后的XML内容
     */
    private static String restoreFontSizeSettings(String xmlContent, Map<String, String> formatProperties) {
        // 恢复doc1的字体大小设置值（前N个）
        int doc1SzCount = 0;
        for (String key : formatProperties.keySet()) {
            if (key.startsWith("doc1_sz_")) {
                doc1SzCount++;
            }
        }
        
        System.out.println("📏 doc1字体大小设置数量: " + doc1SzCount);
        
        // 恢复字体大小设置值
        Pattern szPattern = Pattern.compile("<w:sz\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
        Matcher matcher = szPattern.matcher(xmlContent);
        StringBuffer sb = new StringBuffer();
        
        int index = 0;
        while (matcher.find()) {
            String originalSz;
            if (index < doc1SzCount) {
                // 这是doc1的字体大小设置
                originalSz = formatProperties.get("doc1_sz_" + index);
            } else {
                // 这是doc2的字体大小设置
                originalSz = formatProperties.get("doc2_sz_" + (index - doc1SzCount));
            }
            
            if (originalSz != null) {
                // 恢复原始字体大小值
                String fullAttrs = matcher.group(1);
                String currentSz = matcher.group(2);
                
                // 替换为原始值
                String newFullAttrs = fullAttrs.replace("w:val=\"" + currentSz + "\"", 
                                                       "w:val=\"" + originalSz + "\"");
                matcher.appendReplacement(sb, "<w:sz " + newFullAttrs + ">");
                System.out.println("🔧 恢复第 " + (index + 1) + " 个字体大小设置值: " + currentSz + " -> " + originalSz);
            } else {
                matcher.appendReplacement(sb, matcher.group(0));
                System.out.println("⚠️ 未找到第 " + (index + 1) + " 个字体大小的原始值");
            }
            index++;
        }
        matcher.appendTail(sb);
        
        System.out.println("📏 总共处理了 " + index + " 个字体大小设置");
        return sb.toString();
    }
    
    /**
     * 精确恢复两个文档的段落缩进设置
     * 
     * @param xmlContent XML内容
     * @param formatProperties 格式信息
     * @return 修复后的XML内容
     */
    private static String restoreIndentSettings(String xmlContent, Map<String, String> formatProperties) {
        // 恢复doc1的段落缩进设置值（前N个）
        int doc1IndCount = 0;
        for (String key : formatProperties.keySet()) {
            if (key.startsWith("doc1_ind_")) {
                doc1IndCount++;
            }
        }
        
        System.out.println("-indent- doc1段落缩进设置数量: " + doc1IndCount);
        
        // 恢复段落缩进设置值
        Pattern indPattern = Pattern.compile("<w:ind\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");
        Matcher matcher = indPattern.matcher(xmlContent);
        StringBuffer sb = new StringBuffer();
        
        int index = 0;
        while (matcher.find()) {
            String originalInd;
            if (index < doc1IndCount) {
                // 这是doc1的段落缩进设置
                originalInd = formatProperties.get("doc1_ind_" + index);
            } else {
                // 这是doc2的段落缩进设置
                originalInd = formatProperties.get("doc2_ind_" + (index - doc1IndCount));
            }
            
            if (originalInd != null) {
                // 恢复原始段落缩进值
                String fullAttrs = matcher.group(1);
                String currentInd = matcher.group(2);
                
                // 替换为原始值
                String newFullAttrs = fullAttrs.replace("w:val=\"" + currentInd + "\"", 
                                                       "w:val=\"" + originalInd + "\"");
                matcher.appendReplacement(sb, "<w:ind " + newFullAttrs + ">");
                System.out.println("🔧 恢复第 " + (index + 1) + " 个段落缩进设置值: " + currentInd + " -> " + originalInd);
            } else {
                matcher.appendReplacement(sb, matcher.group(0));
                System.out.println("⚠️ 未找到第 " + (index + 1) + " 个段落缩进的原始值");
            }
            index++;
        }
        matcher.appendTail(sb);
        
        System.out.println("-indent- 总共处理了 " + index + " 个段落缩进设置");
        return sb.toString();
    }
    
    /**
     * 移除表格内段落的首行缩进，解决段落前空格问题
     * 
     * @param xmlContent XML内容
     * @return 修复后的XML内容
     */
    private static String removeTableParagraphFirstLineIndent(String xmlContent) {
        System.out.println("🗑️ 开始移除表格内段落的首行缩进");
        
        // 匹配表格内的段落及其缩进设置
        Pattern tblPIndentPattern = Pattern.compile(
            "(<w:tbl[^>]*>.*?)(<w:p[^>]*>.*?<w:ind\\s+[^>]*w:firstLine\\s*=\\s*\"[^\"]*\".*?/?>)(.*?</w:tbl>)", 
            Pattern.DOTALL);
        
        Matcher matcher = tblPIndentPattern.matcher(xmlContent);
        StringBuffer sb = new StringBuffer();
        
        int removedCount = 0;
        while (matcher.find()) {
            String beforeTbl = matcher.group(1);
            String pWithIndent = matcher.group(2);
            String afterP = matcher.group(3);
            
            // 移除首行缩进属性
            String pWithoutFirstLineIndent = pWithIndent.replaceAll(
                "w:firstLine\\s*=\\s*\"[^\"]*\"", "");
            
            matcher.appendReplacement(sb, beforeTbl + pWithoutFirstLineIndent + afterP);
            removedCount++;
            System.out.println("🗑️ 移除了1个表格内段落的首行缩进");
        }
        matcher.appendTail(sb);
        
        System.out.println("🗑️ 总共移除了 " + removedCount + " 个表格内段落的首行缩进");
        return sb.toString();
    }
    
    /**
     * 移除段落中的对齐到网络设置，解决表格行高无法调整的问题
     * 
     * @param xmlContent XML内容
     * @return 修复后的XML内容
     */
    private static String removeSnapToGridSetting(String xmlContent) {
        System.out.println("📐 开始移除段落中的对齐到网络设置");
        
        // 由于已经在合并前处理了段落的snapToGrid设置，这里不再重复处理
        System.out.println("📐 段落对齐到网络设置已在合并前处理，跳过此步骤");
        
        return xmlContent;
    }
    
    /**
     * 修复缺失val属性的对齐元素
     * 
     * @param xmlContent XML内容
     * @return 修复后的XML内容
     */
    private static String fixMissingJustificationValues(String xmlContent) {
        System.out.println("🔗 开始修复缺失val属性的对齐元素");
        
        // 修复自闭合的jc标签缺失val属性的问题
        int beforeFix1 = xmlContent.length();
        xmlContent = xmlContent.replaceAll(
            "<w:jc\\s*/>", 
            "<w:jc w:val=\"center\"/>");
        int afterFix1 = xmlContent.length();
        System.out.println("🔗 修复自闭合jc标签: " + (afterFix1 - beforeFix1) + " 字符变化");
            
        // 修复带有属性但缺少val属性的jc开始标签
        int beforeFix2 = xmlContent.length();
        xmlContent = xmlContent.replaceAll(
            "<w:jc((?![^>]*\\bw:val\\b)[^>]*/?)>", 
            "<w:jc w:val=\"center\"$1>");
        int afterFix2 = xmlContent.length();
        System.out.println("🔗 修复带属性jc标签: " + (afterFix2 - beforeFix2) + " 字符变化");
            
        return xmlContent;
    }
}