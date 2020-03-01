package sentences;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class WordApp {
  
  private int num;
  
  public WordApp(int num) {
    this.num = num;
  }

  public void writea(List<String> text, List<String> line, List<String> pline, String filePath)
      throws Exception {
    List<String>[] sentences = cuta(text, pline);

    ComThread.InitSTA();
    ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
    wordApp.setProperty("Visible", new Variant(false));
    Dispatch document = wordApp.getProperty("Documents").toDispatch();
    //Dispatch doc = Dispatch.call(document, "Open", new Variant(filePath)).toDispatch();
    Dispatch doc = Dispatch.call(document, "Add").toDispatch();

    int p = text.size() / 9;
    int q = text.size() % 9;
    int[] k = new int[p + 1];
    for (int i = 0; i < p; i++) {
      k[i] = 9;
    }
    if (q <= 5) {
      k[p - 1] = 9 + q;
    } else {
      k[p] = q;
    }

    Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();
    Dispatch font = Dispatch.get(selection, "Font").toDispatch();

    Dispatch alignment = Dispatch.get(selection, "ParagraphFormat").toDispatch(); // 段落格式
    Dispatch.put(selection, "Text", "List " + num);
    Dispatch.put(font, "Bold", new Variant(true));
    Dispatch.put(font, "Size", new Variant(14));
    Dispatch.put(alignment, "Alignment", "1"); // (1:置中 2:靠右 3:靠左)
    Dispatch.call(selection, "EndKey", new Variant(6));
    Dispatch.put(font, "Bold", new Variant(false));
    Dispatch.put(font, "Size", new Variant(10.5));
    Dispatch.call(selection, "TypeParagraph"); // 插入一个空行
    Dispatch.put(alignment, "Alignment", "3");
    for (int i = 0; i < text.size(); i++) {
      if (i % 9 == 0) {
        if (i / 9 != p || q >= 6) {
          Dispatch.put(selection, "Text", "\r\n");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
        for (int j = 0; j < k[i / 9]; j++) {
          Dispatch.put(font, "Bold", new Variant(true));
          Dispatch.put(selection, "Text", line.get(i / 9 * 9 + j) + " ");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
        Dispatch.put(font, "Bold", new Variant(false));
        if (i / 9 != p || q >= 6) {
          Dispatch.put(selection, "Text", "\r\n\r\n");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
      }
      if (i / 9 != p || q >= 6) {
        Dispatch.put(selection, "Text", (i % 9 + 1) + "." + sentences[i].get(0));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(1));
        Dispatch.put(font, "Underline", new Variant(true));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(2) + "\r\n");
        Dispatch.put(font, "Underline", new Variant(false));
        Dispatch.call(selection, "EndKey", new Variant(6));
      } else {
        Dispatch.put(selection, "Text", (i % 9 + 10) + "." + sentences[i].get(0));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(1));
        Dispatch.put(font, "Underline", new Variant(true));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(2) + "\r\n");
        Dispatch.put(font, "Underline", new Variant(false));
        Dispatch.call(selection, "EndKey", new Variant(6));
      }
    }

    Dispatch.call(doc, "SaveAs", filePath);
    Dispatch.call(doc, "Close", new Variant(0));
    Dispatch.call(wordApp, "Quit");
    ComThread.Release();
  }

  @SuppressWarnings("unchecked")
  public List<String>[] cuta(List<String> text, List<String> pline) throws Exception {
    List<String>[] sentences = new List[text.size()];
    String regex0 = "(.*[,|.|!|\\s|?|:|-])("; // "([A-Za-z0-9|\\s|,|'|-|?|!|\"|—|+]*)("
    String regex2 = ")([,|.|!|\\s|?|:|-].*)"; // ")([A-Za-z0-9|\\s|,|'|-|?|!|\"|—|+]*.)"
    for (int i = 0; i < text.size(); i++) {
      Pattern pattern1 = Pattern.compile(regex0 + pline.get(i) + regex2, Pattern.CASE_INSENSITIVE);
      Pattern pattern0 = Pattern.compile("(" + pline.get(i) + regex2, Pattern.CASE_INSENSITIVE);
      Pattern pattern2 = Pattern.compile(regex0 + pline.get(i) + ")", Pattern.CASE_INSENSITIVE);
      Matcher matcher1 = pattern1.matcher(text.get(i));
      Matcher matcher0 = pattern0.matcher(text.get(i));
      Matcher matcher2 = pattern2.matcher(text.get(i));
      if (matcher1.find()) {
        if (matcher1.groupCount() != 3) {
          throw new Exception();
        }
        sentences[i] = new ArrayList<>();
        sentences[i].add(matcher1.group(1));
        sentences[i].add(matcher1.group(2));
        sentences[i].add(matcher1.group(3));
      } else if (matcher0.find()) {
        sentences[i] = new ArrayList<>();
        sentences[i].add("");
        sentences[i].add(matcher0.group(1));
        sentences[i].add(matcher0.group(2));
      } else if (matcher2.find()) {
        sentences[i] = new ArrayList<>();
        sentences[i].add(matcher2.group(1));
        sentences[i].add(matcher2.group(2));
        sentences[i].add("");
      } else {
        throw new Exception();
      }
    }
    return sentences;
  }

  public void writeq(List<String> text, List<String> line, List<String> pline, String filePath)
      throws Exception {
    List<String>[] sentences = cutq(text, pline);

    ComThread.InitSTA();
    ActiveXComponent wordApp = new ActiveXComponent("Word.Application");
    wordApp.setProperty("Visible", new Variant(false));
    Dispatch document = wordApp.getProperty("Documents").toDispatch();
    //Dispatch doc = Dispatch.call(document, "Open", new Variant(filePath)).toDispatch();
    Dispatch doc = Dispatch.call(document, "Add").toDispatch();

    int p = text.size() / 9;
    int q = text.size() % 9;
    int[] k = new int[p + 1];
    for (int i = 0; i < p; i++) {
      k[i] = 9;
    }
    if (q <= 5) {
      k[p - 1] = 9 + q;
    } else {
      k[p] = q;
    }

    Dispatch selection = Dispatch.get(wordApp, "Selection").toDispatch();
    Dispatch font = Dispatch.get(selection, "Font").toDispatch();

    Dispatch alignment = Dispatch.get(selection, "ParagraphFormat").toDispatch(); // 段落格式
    Dispatch.put(selection, "Text", "List " + num);
    Dispatch.put(font, "Bold", new Variant(true));
    Dispatch.put(font, "Size", new Variant(14));
    Dispatch.put(alignment, "Alignment", "1"); // (1:置中 2:靠右 3:靠左)
    Dispatch.call(selection, "EndKey", new Variant(6));
    Dispatch.put(font, "Bold", new Variant(false));
    Dispatch.put(font, "Size", new Variant(10.5));
    Dispatch.call(selection, "TypeParagraph"); // 插入一个空行
    Dispatch.put(alignment, "Alignment", "3");
    for (int i = 0; i < text.size(); i++) {
      if (i % 9 == 0) {
        if (i / 9 != p || q >= 6) {
          Dispatch.put(selection, "Text", "\r\n");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
        for (int j = 0; j < k[i / 9]; j++) {
          Dispatch.put(font, "Bold", new Variant(true));
          Dispatch.put(selection, "Text", line.get(i / 9 * 9 + j) + " ");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
        Dispatch.put(font, "Bold", new Variant(false));
        if (i / 9 != p || q >= 6) {
          Dispatch.put(selection, "Text", "\r\n\r\n");
          Dispatch.call(selection, "EndKey", new Variant(6));
        }
      }
      if (i / 9 != p || q >= 6) {
        Dispatch.put(selection, "Text", (i % 9 + 1) + "." + sentences[i].get(0));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(1));
        Dispatch.put(font, "Underline", new Variant(true));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(2) + "\r\n");
        Dispatch.put(font, "Underline", new Variant(false));
        Dispatch.call(selection, "EndKey", new Variant(6));
      } else {
        Dispatch.put(selection, "Text", (i % 9 + 10) + "." + sentences[i].get(0));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(1));
        Dispatch.put(font, "Underline", new Variant(true));
        Dispatch.call(selection, "EndKey", new Variant(6));
        Dispatch.put(selection, "Text", sentences[i].get(2) + "\r\n");
        Dispatch.put(font, "Underline", new Variant(false));
        Dispatch.call(selection, "EndKey", new Variant(6));
      }
    }

    Dispatch.call(doc, "SaveAs", filePath);
    Dispatch.call(doc, "Close", new Variant(0));
    Dispatch.call(wordApp, "Quit");
    ComThread.Release();
  }

  @SuppressWarnings("unchecked")
  public List<String>[] cutq(List<String> text, List<String> pline) throws Exception {
    List<String>[] sentences = new List[text.size()];
    String regex0 = "(.*[,|.|!|\\s|?|:|-])(";
    String regex2 = ")([,|.|!|\\s|?|:|-].*)";
    for (int i = 0; i < text.size(); i++) {
      Pattern pattern1 = Pattern.compile(regex0 + pline.get(i) + regex2, Pattern.CASE_INSENSITIVE);
      Pattern pattern0 = Pattern.compile("(" + pline.get(i) + regex2, Pattern.CASE_INSENSITIVE);
      Pattern pattern2 = Pattern.compile(regex0 + pline.get(i) + ")", Pattern.CASE_INSENSITIVE);
      Matcher matcher1 = pattern1.matcher(text.get(i));
      Matcher matcher0 = pattern0.matcher(text.get(i));
      Matcher matcher2 = pattern2.matcher(text.get(i));
      if (matcher1.find()) {
        if (matcher1.groupCount() != 3) {
          throw new Exception();
        }
        sentences[i] = new ArrayList<>();
        sentences[i].add(matcher1.group(1));
        sentences[i].add(space(matcher1.group(2)));
        sentences[i].add(matcher1.group(3));
      } else if (matcher0.find()) {
        sentences[i] = new ArrayList<>();
        sentences[i].add("");
        sentences[i].add(space(matcher0.group(1)));
        sentences[i].add(matcher0.group(2));
      } else if (matcher2.find()) {
        sentences[i] = new ArrayList<>();
        sentences[i].add(matcher2.group(1));
        sentences[i].add(space(matcher2.group(2)));
        sentences[i].add("");
      } else {
        throw new Exception();
      }
    }
    return sentences;
  }

  public String space(String string) {
    return "           ";
  }
}
