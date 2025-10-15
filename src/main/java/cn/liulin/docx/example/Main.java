package cn.liulin.docx.example;

import java.util.ArrayList;
import java.util.List;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class Main {
    public static void main(String[] args) {
        try {
            // 请确保这两个文件存在
            String doc1Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\2023年度安全生产费用使用管理台账.docx";
            String doc2Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\鲁CV5566.docx";
            String doc3Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\川AAB980-行车日志.docx";
            String doc4Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\鲁C653E挂.docx";
            String doc5Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\审验表.docx";
            String outputPath = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\3.docx";

//            DocxMerger merger = new DocxMerger();
//            merger.merge(doc1Path, doc2Path, outputPath);
            List<String> list = new ArrayList<>();
            list.add(doc1Path);
            list.add(doc2Path);
            list.add(doc3Path);
            list.add(doc4Path);
            list.add(doc5Path);
            DocxMergerList docxMergerList = new DocxMergerList();
            docxMergerList.mergeList(list, outputPath
            );

            System.out.println("🎉 合并成功！输出文件: " + outputPath);
        } catch (Exception e) {
            System.err.println("❌ 合并失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
