package sentences;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class Work {

  private int textsize;
  private List<String> easywords;
  
  public Work(List<String> easywords) {
    this.easywords = easywords;
  }

  public List<String> read(File file) throws IOException {
    BufferedReader reader = new BufferedReader(new FileReader(file));
    List<String> ptext = new ArrayList<>();
    // List<String> text = new ArrayList<>();
    String temp;
    while ((temp = reader.readLine()) != null) {
      ptext.add(temp);
    }
    reader.close();
    this.textsize = ptext.size();
    // System.out.println(ptext.size());
    // text = exchange(ptext, 0, ptext.size() - 1);

    return ptext;
  }

  /**
   * 合并若干个String类型的List为一个.
   */
  @SuppressWarnings("unchecked")
  public List<String> combine(List<String>... values) {
    List<String> list = new ArrayList<>();
    List<String>[] lists = values;
    for (int i = 0; i < lists.length; i++) {
      for (String j : lists[i]) {
        list.add(j);
      }
    }
    return list;
  }

  /**
   * 返回start到end之间随机打乱的部分.
   */
  public List<String> exchange(List<String> plist, int start, int end) {
    int n = end - start + 1;
    int randInt = 0;
    Random random = new Random();
    int[] rand = new int[n];
    boolean[] bool = new boolean[n];
    for (int i = 0; i < n; i++) {
      do {
        randInt = random.nextInt(n);
      } while (bool[randInt]);
      bool[randInt] = true;
      rand[i] = randInt;
    }
    String[] temp = new String[n];
    for (int i = start; i <= end; i++) {
      temp[i - start] = plist.get(start + rand[i - start]);
    }
    List<String> list = new ArrayList<String>();
    for (int i = 0; i < temp.length; i++) {
      list.add(temp[i]);
    }
    return list;
  }

  @SuppressWarnings("unchecked")
  public List<String> cutwords(List<String> pline) {
    List<String> line = new ArrayList<>();
    int p = textsize / 9;
    int q = textsize % 9;
    int k = p + (q > 5 ? 1 : 0);
    List<String>[] lists = new List[k];

    for (int i = 0; i < k - 1; i++) {
      lists[i] = exchange(pline, 9 * i, 9 * i + 8);
    }
    lists[k - 1] = exchange(pline, 9 * (k - 1), textsize - 1);

    line = combine(lists);

    return line;
  }

  public List<String> extract(List<String> text, String startword, String endword) {
    List<String> pline = new ArrayList<>();
    for (int j = 0; j < textsize; j++) {
      String sentence = text.get(j);
      sentence = sentence.toLowerCase();
      String[] strarry = sentence.split("[,|.|:|!|?|-|\\s+]");
      char[] word = strarry[strarry.length - 1].toCharArray();

      if (word[word.length - 1] == '.' || word[word.length - 1] == '!' || word[word.length - 1] == '?') {
        char[] nword = new char[word.length - 1];
        for (int i = 0; i < word.length - 1; i++) {
          nword[i] = word[i];
        }
        String end = new String(nword);
        strarry[strarry.length - 1] = end;
      }

      boolean flag = true;
      for (String s : strarry) {
        if (s.compareTo(startword) >= 0 && s.compareTo(endword) <= 0 && !easywords.contains(s)) {
          flag = false;
          pline.add(s);
          break;
        }
      }
      if (flag) {
        System.out.println("line: " + j + " not matched");
      }
    }

    return pline;
  }

  public void write(List<String> text, List<String> line, File file) throws IOException {
    FileWriter fr = new FileWriter(file);
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
    for (int i = 0; i < text.size(); i++) {
      if (i % 9 == 0) {
        if (i / 9 != p || q >= 6) {
          fr.write("\r\n");
        }
        for (int j = 0; j < k[i / 9]; j++) {
          fr.write(line.get(i / 9 * 9 + j) + " ");
        }
        if (i / 9 != p || q >= 6) {
          fr.write("\r\n\r\n");
        }
      }
      if (i / 9 != p || q >= 6) {
        fr.write((i % 9 + 1) + "." + text.get(i) + "\r\n");
      } else {
        fr.write((i % 9 + 10) + "." + text.get(i) + "\r\n");
      }
    }
    fr.close();
  }
}
