package cn.liulin.docx.util;

import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;
import java.util.HashMap;
import java.util.Map;
import java.util.List;

public class NumberingMapperUtil {
    private static final Logger logger = LoggerFactory.getLogger(NumberingMapperUtil.class);

    /**
     * 将 doc2 的编号定义合并到 doc1，并重映射 numId 避免冲突
     */
    public static void mapNumbering(List<WordprocessingMLPackage> docPath) {
        try {
            NumberingDefinitionsPart ndp1 = docPath.get(0).getMainDocumentPart().getNumberingDefinitionsPart();

            // 如果其中一个文档没有编号定义部分，则创建一个新的
            if (ndp1 == null) {
                logger.debug("为文档1创建编号定义部分");
                ndp1 = new NumberingDefinitionsPart();
                ndp1.setJaxbElement(new Numbering());
                docPath.get(0).getMainDocumentPart().addTargetPart(ndp1);
            }

            Numbering numbering1 = ndp1.getJaxbElement();
            // 1. 找出 doc1 中最大的 numId
            BigInteger maxNumId = BigInteger.ZERO;
            List<Numbering.Num> nums1 = numbering1.getNum();
            for (Numbering.Num num : nums1) {
                BigInteger id = num.getNumId();
                if (id != null && id.compareTo(maxNumId) > 0) {
                    maxNumId = id;
                }
            }

            // 2. 映射表：旧 numId -> 新 numId
            Map<BigInteger, BigInteger> numIdMap = new HashMap<>();

            for (int i = 1; i < docPath.size(); i++) {
                NumberingDefinitionsPart tempNdp = docPath.get(i).getMainDocumentPart().getNumberingDefinitionsPart();
                if (tempNdp == null) {
                    logger.warn("合并文档缺少编号定义部分，跳过编号映射");
                    continue;
                }
                Numbering tempNumbering = tempNdp.getJaxbElement();
                // 3. 遍历 doc2 的编号，重映射 numId 并合并到 doc1
                List<Numbering.Num> tempNums = tempNumbering.getNum();
                for (Numbering.Num num : tempNums) {
                    BigInteger oldId = num.getNumId();
                    if (oldId == null) {
                        continue;
                    }

                    // 新 ID = max + 1 + oldId（确保唯一）
                    BigInteger newId = maxNumId.add(BigInteger.ONE).add(oldId);

                    // 修改 numId
                    num.setNumId(newId);

                    // 记录映射
                    numIdMap.put(oldId, newId);

                    // 添加到 doc1 的编号定义中
                    numbering1.getNum().add(num);
                }

                // 4. 更新 合并 doc 内容中的编号引用（段落）
                List<Object> content = docPath.get(i).getMainDocumentPart().getContent();
                for (Object obj : content) {
                    if (obj instanceof P) {
                        P p = (P) obj;
                        PPr ppr = p.getPPr();
                        if (ppr != null && ppr.getNumPr() != null && ppr.getNumPr().getNumId() != null) {
                            BigInteger ref = ppr.getNumPr().getNumId().getVal();
                            if (ref != null && numIdMap.containsKey(ref)) {
                                ppr.getNumPr().getNumId().setVal(numIdMap.get(ref));
                                logger.debug("更新段落编号引用: {} -> {}", ref, numIdMap.get(ref));
                            }
                        }
                    }
                }
            }

            logger.debug("编号映射完成，共处理 {} 个编号", numIdMap.size());

        } catch (Exception e) {
            logger.error("编号映射失败：", e);
        }
    }
}