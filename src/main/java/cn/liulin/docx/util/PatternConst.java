package cn.liulin.docx.util;

import java.util.regex.Pattern;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/15 15:41
 */
public interface PatternConst {
    /**
     * trHeight元素属性
     */
    Pattern TR_HEIGHT_PATTERN = Pattern.compile("<w:trHeight\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");

    /**
     * tbl元素属性
     */
    Pattern TBL_PATTERN = Pattern.compile("<w:tbl(?:\\s[^>]*)?>(.*?)</w:tbl>", Pattern.DOTALL);

    /**
     * 段落缩进信息（特别是表格内的段落）
     */
    Pattern INDENT_PATTERN = Pattern.compile("<w:ind\\s+([^>]+w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");

    /**
     * 段落样式(Normal样式)
     */
    Pattern STYLE_PATTERN = Pattern.compile(
            "<w:style[^>]*w:type=\"paragraph\"[^>]*>.*?<w:name\\s+w:val=\"Normal\"\\s*/>.*?</w:style>",
            Pattern.DOTALL);

    /**
     * 字体主题信息
     */
    Pattern FONT_PATTERN = Pattern.compile(
            "<w:rFonts\\s+([^>]*w:asciiTheme\\s*=\\s*\"([^\"]+)\"[^>]*w:hAnsiTheme\\s*=\\s*\"([^\"]+)\"[^>]*w:eastAsiaTheme\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");

    /**
     * 字体大小信息
     */
    Pattern SIZE_PATTERN = Pattern.compile("<w:sz\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");

    /**
     * 复杂字体大小信息
     */
    Pattern SIZE_CS_PATTERN = Pattern.compile("<w:szCs\\s+([^>]*w:val\\s*=\\s*\"([^\"]+)\"[^>]*)/?>");

    /**
     * 表格单元格中的<w:r>元素
     */
    Pattern R_PATTERN = Pattern.compile("(<w:r[^>]*>\\s*<w:rPr[^>]*>)(.*?)(</w:rPr>)", Pattern.DOTALL);

}
