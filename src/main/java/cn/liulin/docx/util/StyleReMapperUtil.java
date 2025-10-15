package cn.liulin.docx.util;

import org.docx4j.TraversalUtil;
import org.docx4j.finders.ClassFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class StyleReMapperUtil {
    private static final Logger logger = LoggerFactory.getLogger(StyleReMapperUtil.class);

    private static final String SUFFIX = "_DOC";

    /**
     * 为文档列表中的每个文档重命名样式，以避免在合并时出现样式冲突
     * 该方法会遍历文档列表，为每个文档的样式添加唯一的后缀标识
     * 
     * @param docList 包含WordprocessingMLPackage对象的文档列表
     */
    public static void renameDocListStyles(List<WordprocessingMLPackage> docList) {
        // 遍历文档列表，为每个文档的样式添加唯一后缀
        for (int i = 0; i < docList.size(); i++) {
            renameStyles(docList.get(i), SUFFIX + (i + 1));
        }
    }

    /**
     * 重命名文档中的所有样式ID，以避免在文档合并时出现样式冲突
     * 该方法会遍历文档中的所有样式定义，并为每个样式ID添加指定的后缀，
     * 同时更新所有引用这些样式的段落和表格元素
     * 
     * @param doc 要处理的Word文档对象
     * @param suffix 要添加到样式ID后的后缀字符串
     */
    private static void renameStyles(WordprocessingMLPackage doc, String suffix) {
        Styles styles = doc.getMainDocumentPart().getStyleDefinitionsPart().getJaxbElement();
        // 样式ID映射
        Map<String, String> styleIdMap = new HashMap<>();

        if (styles != null && styles.getStyle() != null) {
            for (Style style : styles.getStyle()) {
                String origId = style.getStyleId();
                
                if (origId != null) {
                    String newId = origId + suffix;
                    style.setStyleId(newId);
                    styleIdMap.put(origId, newId);

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
                            logger.info("更新段落样式引用: {} -> {}", oldStyleId, newStyleId);
                        }
                    }
                }
            }
            
            // 更新文档中所有表格的样式引用（全部使用样式ID映射）
            updateTableStyleReferencesById(doc.getMainDocumentPart().getContent(), styleIdMap);
        }
    }
    
    /**
     * 更新文档中所有表格的样式引用（使用样式ID映射）
     * 
     * @param content 文档内容
     * @param styleIdMap 样式ID映射
     */
    private static void updateTableStyleReferencesById(List<Object> content, Map<String, String> styleIdMap) {
        // 使用ClassFinder和TraversalUtil来查找所有表格对象
        ClassFinder finder = new ClassFinder(Tbl.class);
        new TraversalUtil(content, finder);
        
        // 遍历所有找到的表格对象
        for (Object obj : finder.results) {
            if (obj instanceof Tbl) {
                Tbl tbl = (Tbl) obj;
                if (tbl.getTblPr() != null && tbl.getTblPr().getTblStyle() != null) {
                    CTTblPrBase.TblStyle tblStyle = tbl.getTblPr().getTblStyle();
                    String oldStyleId = tblStyle.getVal();
                    
                    // 如果表格有明确的样式ID，则更新引用
                    if (oldStyleId != null && styleIdMap.containsKey(oldStyleId)) {
                        String newStyleId = styleIdMap.get(oldStyleId);
                        tblStyle.setVal(newStyleId);
                        logger.info("更新表格样式引用 (ID映射): {} -> {}", oldStyleId, newStyleId);
                    }
                }
            }
        }
    }
}