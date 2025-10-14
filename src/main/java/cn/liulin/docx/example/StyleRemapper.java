package cn.liulin.docx.example;

import org.docx4j.TraversalUtil;
import org.docx4j.finders.ClassFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class StyleRemapper {

    public static void renameStyles(WordprocessingMLPackage doc, String suffix) {
        Styles styles = doc.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement();
        // 样式ID映射
        Map<String, String> styleIdMap = new HashMap<>();
        // 样式名称映射
        Map<String, String> styleNameMap = new HashMap<>();
        // 样式ID到样式名称的映射（用于处理表格引用的是样式ID而非名称的情况）
        Map<String, String> styleIdToNameMap = new HashMap<>();

        if (styles != null && styles.getStyle() != null) {
            for (Style style : styles.getStyle()) {
                String origId = style.getStyleId();
                // 获取样式的名称
                String styleName = null;
                if (style.getName() != null) {
                    styleName = style.getName().getVal();
                }
                
                if (origId != null) {
                    String newId = origId + suffix;
                    style.setStyleId(newId);
                    styleIdMap.put(origId, newId);
                    
                    // 如果有样式名称，创建名称映射并更新样式名称
                    if (styleName != null) {
                        String newStyleName = styleName + suffix;
                        style.getName().setVal(newStyleName);
                        styleNameMap.put(styleName, newStyleName);
                        System.out.println("🔄 样式名称映射: " + styleName + " -> " + newStyleName);
                        // 建立样式ID到样式名称的映射
                        styleIdToNameMap.put(origId, styleName);
                    }

                    // 更新基于的样式
                    if (style.getBasedOn() != null) {
                        String basedOn = style.getBasedOn().getVal();
                        if (basedOn != null) {
                            style.getBasedOn().setVal(basedOn + suffix);
                        }
                    }
                    // 更新链接的样式
                    if (style.getLink() != null) {
                        String link = style.getLink().getVal();
                        if (link != null) {
                            style.getLink().setVal(link + suffix);
                        }
                    }
                    // 更新表格样式引用中的样式名称
                    if (style.getTblPr() != null && style.getTblPr().getTblStyle() != null) {
                        CTTblPrBase.TblStyle tblStyle = style.getTblPr().getTblStyle();
                        String tblStyleVal = tblStyle.getVal();
                        if (tblStyleVal != null && styleNameMap.containsKey(tblStyleVal)) {
                            tblStyle.setVal(styleNameMap.get(tblStyleVal));
                        }
                    }
                }
            }

            // 更新文档中所有段落的样式引用（使用样式ID映射）
            List<Object> paragraphs = doc.getMainDocumentPart().getContent();
            for (Object obj : paragraphs) {
                if (obj instanceof P) {
                    P p = (P) obj;
                    PPr ppr = p.getPPr();
                    if (ppr != null && ppr.getPStyle() != null) {
                        PPrBase.PStyle pStyle = ppr.getPStyle();
                        String oldStyleId = pStyle.getVal();
                        // 段落样式引用的是样式ID
                        if (oldStyleId != null && styleIdMap.containsKey(oldStyleId)) {
                            String newStyleId = styleIdMap.get(oldStyleId);
                            pStyle.setVal(newStyleId);
                            System.out.println("🔄 更新段落样式引用: " + oldStyleId + " -> " + newStyleId);
                        }
                    }
                }
            }
            
            // 更新文档中所有表格的样式引用（使用样式名称映射）
            updateTableStyleReferences(doc.getMainDocumentPart().getContent(), styleNameMap, styleIdToNameMap, styleIdMap);
        }
    }
    
    /**
     * 更新文档中所有表格的样式引用
     * 
     * @param content 文档内容
     * @param styleNameMap 样式名称映射
     * @param styleIdToNameMap 样式ID到样式名称的映射
     * @param styleIdMap 样式ID映射
     */
    private static void updateTableStyleReferences(List<Object> content, Map<String, String> styleNameMap, 
            Map<String, String> styleIdToNameMap, Map<String, String> styleIdMap) {
        // 使用ClassFinder和TraversalUtil来查找所有表格对象
        ClassFinder finder = new ClassFinder(Tbl.class);
        new TraversalUtil(content, finder);
        
        // 遍历所有找到的表格对象
        for (Object obj : finder.results) {
            if (obj instanceof Tbl) {
                Tbl tbl = (Tbl) obj;
                if (tbl.getTblPr() != null && tbl.getTblPr().getTblStyle() != null) {
                    CTTblPrBase.TblStyle tblStyle = tbl.getTblPr().getTblStyle();
                    String oldStyleRef = tblStyle.getVal(); // 这可能是样式ID或样式名称
                    
                    // 首先尝试直接通过样式名称映射查找
                    if (oldStyleRef != null && styleNameMap.containsKey(oldStyleRef)) {
                        String newStyleName = styleNameMap.get(oldStyleRef);
                        tblStyle.setVal(newStyleName);
                        System.out.println("🔄 更新表格样式引用 (直接名称映射): " + oldStyleRef + " -> " + newStyleName);
                    } 
                    // 如果没有找到，尝试通过样式ID到名称的映射
                    else if (oldStyleRef != null && styleIdToNameMap.containsKey(oldStyleRef)) {
                        String styleName = styleIdToNameMap.get(oldStyleRef);
                        if (styleNameMap.containsKey(styleName)) {
                            String newStyleName = styleNameMap.get(styleName);
                            tblStyle.setVal(newStyleName);
                            System.out.println("🔄 更新表格样式引用 (ID->名称映射): " + oldStyleRef + " -> " + newStyleName);
                        }
                    }
                    // 最后尝试通过样式ID映射查找
                    else if (oldStyleRef != null && styleIdMap.containsKey(oldStyleRef)) {
                        String newStyleId = styleIdMap.get(oldStyleRef);
                        tblStyle.setVal(newStyleId);
                        System.out.println("🔄 更新表格样式引用 (ID映射): " + oldStyleRef + " -> " + newStyleId);
                    }
                }
            }
        }
    }
}