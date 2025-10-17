package cn.liulin.docx.util;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class StyleReMapperUtil {
    private static final Logger logger = LoggerFactory.getLogger(StyleReMapperUtil.class);

    /**
     * 合并样式，以第一个文档的样式为基准
     * @param docList 文档列表
     */
    public static void mergeStyles(List<WordprocessingMLPackage> docList) {
        try {
            logger.info("开始合并样式，以第一个文档为基准");

            // 获取第一个文档的样式定义
            StyleDefinitionsPart baseStylePart = docList.get(0).getMainDocumentPart().getStyleDefinitionsPart();
            if (baseStylePart == null) {
                logger.warn("第一个文档没有样式定义部分");
                return;
            }

            Styles baseStyles = baseStylePart.getJaxbElement();
            if (baseStyles == null) {
                logger.warn("第一个文档没有样式定义");
                return;
            }

            // 遍历后续文档，合并新增样式
            for (int i = 1; i < docList.size(); i++) {
                StyleDefinitionsPart currentStylePart = docList.get(i).getMainDocumentPart().getStyleDefinitionsPart();
                if (currentStylePart == null) {
                    logger.debug("文档 {} 没有样式定义部分", i);
                    continue;
                }

                Styles currentStyles = currentStylePart.getJaxbElement();
                if (currentStyles == null) {
                    logger.debug("文档 {} 没有样式定义", i);
                    continue;
                }

                // 获取当前文档的样式列表
                List<Style> currentStyleList = currentStyles.getStyle();
                if (currentStyleList == null || currentStyleList.isEmpty()) {
                    logger.debug("文档 {} 没有样式列表", i);
                    continue;
                }

                // 遍历当前文档的样式
                for (Style currentStyle : currentStyleList) {
                    // 检查该样式是否已存在于基础文档中
                    boolean styleExists = false;
                    String currentStyleId = currentStyle.getStyleId();
                    String currentStyleType = currentStyle.getType();
                    boolean currentStyleDefault = currentStyle.isDefault();

                    // 检查基础文档中是否已存在相同类型、ID和默认设置的样式
                    for (Style baseStyle : baseStyles.getStyle()) {
                        if (baseStyle.getType() != null && baseStyle.getType().equals(currentStyleType) &&
                                baseStyle.getStyleId() != null && baseStyle.getStyleId().equals(currentStyleId) &&
                                baseStyle.isDefault() == currentStyleDefault) {
                            styleExists = true;
                            break;
                        }
                    }

                    // 如果样式不存在，则添加到基础文档中
                    if (!styleExists) {
                        baseStyles.getStyle().add(currentStyle);
                        logger.debug("添加新样式: type={}, styleId={}, default={}",
                                currentStyleType, currentStyleId, currentStyleDefault);
                    } else {
                        logger.debug("样式已存在，跳过: type={}, styleId={}, default={}",
                                currentStyleType, currentStyleId, currentStyleDefault);
                    }
                }
            }

            logger.info("样式合并完成");
        } catch (Exception e) {
            logger.error("合并样式时出错: {}", e.getMessage(), e);
        }
    }
}