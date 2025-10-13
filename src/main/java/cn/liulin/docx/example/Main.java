package cn.liulin.docx.example;

/**
 * @author liulin
 * @version 1.0
 * @date 2025/10/11 14:38
 */
public class Main {
    public static void main(String[] args) {
        try {
            // 请确保这两个文件存在
            String doc1Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\测试.docx";
            String doc2Path = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\鲁CV5566.docx";
            String outputPath = "D:\\IdeaWorkSpace_Study\\docx-merge-advanced\\word\\3.docx";

            DocxMerger merger = new DocxMerger();
            merger.merge(doc1Path, doc2Path, outputPath);

            System.out.println("🎉 合并成功！输出文件: " + outputPath);
        } catch (Exception e) {
            System.err.println("❌ 合并失败: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
