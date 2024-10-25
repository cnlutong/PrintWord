import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        try {
            String name = "张三";
            String date1 = "2024-10-22";
            String date2 = "2024-10-23";
            String wordCount = "50";

            List<String> reviewDates = Arrays.asList(
                    "Z1", "Z2", "Z3", "Z3", "Z4", "Z5", "Z6", "Z7", "Z8", "Z9"
            );

            List<WordTemplateGenerator.Pair<String, WordTemplateGenerator.Pair<String, String>>> wordPairs = new ArrayList<>();
            // 添加示例单词
            wordPairs.add(new WordTemplateGenerator.Pair<>("ring",
                    new WordTemplateGenerator.Pair<>("[rɪŋ]", "v.（使）发出钟声，响起铃声")));

            // 添加更多单词对
            for (int i = 1; i < 36; i++) {
                wordPairs.add(new WordTemplateGenerator.Pair<>("word" + i,
                        new WordTemplateGenerator.Pair<>("[word" + i + "]", "meaning" + i)));
            }

            WordTemplateGenerator.generateDocument(
                    "output.docx",
                    name,
                    date1,
                    date2,
                    wordCount,
                    reviewDates,
                    wordPairs
            );

            System.out.println("文档生成成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}