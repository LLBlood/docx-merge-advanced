package cn.liulin.docx.example;

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
        Map<String, String> styleMap = new HashMap<>();

        if (styles != null && styles.getStyle() != null) {
            for (Style style : styles.getStyle()) {
                String origId = style.getStyleId();
                if (origId != null) {
                    String newId = origId + suffix;
                    style.setStyleId(newId);
                    styleMap.put(origId, newId);

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

            // 更新文档中所有段落的样式引用
            List<Object> paragraphs = doc.getMainDocumentPart().getContent();
            for (Object obj : paragraphs) {
                if (obj instanceof P) {
                    P p = (P) obj;
                    PPr ppr = p.getPPr();
                    if (ppr != null && ppr.getPStyle() != null) {
                        PPrBase.PStyle pStyle = ppr.getPStyle();
                        String oldStyle = pStyle.getVal();
                        if (styleMap.containsKey(oldStyle)) {
                            pStyle.setVal(styleMap.get(oldStyle));
                        }
                    }
                }
            }
        }
    }
}