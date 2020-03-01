package sentences;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class Main {

  public static void main(String[] args) throws Exception {
    List<String> easywords = new ArrayList<>();
    String[] easy = {"is", "are", "was", "were", "be", "been", "an", "and", "ants", "birds",
        "broken", "by", "college", "come", "coming", "computers", "die", "easily", "during",
        "early", "doctor", "every", "fear", "from", "further", "fish", "foward", "had", "got",
        "her", "how", "have", "has", "his", "he", "in", "jazz", "mad", "made", "many", "man", "may",
        "more", "meet", "no", "moved", "most", "much", "music", "of", "our", "own", "or", "player",
        "period", "people\'s", "perfect", "places", "prize", "she\'ll", "stones", "stores", "take",
        "the", "they", "that", "these", "those", "to", "third", "than", "their", "travel", "time",
        "tiny", "tissues"};
    for (String s : easy) {
      easywords.add(s);
    }
    Work work = new Work(easywords);
    String[][] se = new String[36][2];
    se[1][0] = "abandon";
    se[1][1] = "alarm";
    se[2][0] = "alchemist";
    se[2][1] = "associate";
    se[3][0] = "assume";
    se[3][1] = "blip";
    se[4][0] = "blizzard";
    se[4][1] = "cascade";
    se[5][0] = "cascare";
    se[5][1] = "coil";
    se[6][0] = "coincide";
    se[6][1] = "configuration";
    se[7][0] = "confine";
    se[7][1] = "crack";
    se[8][0] = "craft";
    se[8][1] = "delicate";
    se[9][0] = "delicious";
    se[9][1] = "discrete";
    se[10][0] = "discrimination";
    se[10][1] = "eloquent";
    se[11][0] = "elusive";
    se[11][1] = "exceed";
    se[12][0] = "excess";
    se[12][1] = "fiction ";
    se[13][0] = "fidelity";
    se[13][1] = "gender";
    se[14][0] = "gene";
    se[14][1] = "harbor";
    se[15][0] = "hardware";
    se[15][1] = "illuminate";
    se[16][0] = "illusion";
    se[16][1] = "inhibit";
    se[17][0] = "initial";
    se[17][1] = "jeopardize";
    se[18][0] = "";
    se[18][1] = "";
    se[19][0] = "lithosphere";
    se[19][1] = "measure";
    se[20][0] = "mechanic";
    se[20][1] = "mosaic";
    se[21][0] = "mosquito";
    se[21][1] = "obstruct";
    se[22][0] = "obtain";
    se[22][1] = "patent";
    se[23][0] = "path";
    se[23][1] = "plunge";
    se[24][0] = "pointed";
    se[24][1] = "process";
    se[25][0] = "proclaim";
    se[25][1] = "rainfall";
    se[26][0] = "rally";
    se[26][1] = "renaissance";
    se[27][0] = "render";
    se[27][1] = "routine";
    se[28][0] = "rub";
    se[28][1] = "sequoia";
    se[29][0] = "series";
    se[29][1] = "soluble";
    se[30][0] = "solution";
    se[30][1] = "staple";
    se[31][0] = "starch";
    se[31][1] = "substitute";
    se[32][0] = "subtle";
    se[32][1] = "tedium";
    se[33][0] = "teem";
    se[33][1] = "tuition";
    se[34][0] = "";
    se[34][1] = "";
    se[35][0] = "";
    se[35][1] = "";
    for (int i = 33; i <= 33; i++) {
      if (se[i][0].equals("")) {
        continue;
      }
      List<String> ptext;
      List<String> pline;
      List<String> text;
      List<String> line;
      ptext = work.read(new File("files/List" + i + ".txt"));
      text = work.exchange(ptext, 0, ptext.size() - 1);
      pline = work.extract(text, se[i][0], se[i][1]);
      line = work.cutwords(pline);
      // work.write(text, line, new File("products/7.txt"));
      WordApp wordAppA = new WordApp(i);
      wordAppA.writea(text, line, pline,
          "C:\\Users\\lzlghvbn\\Desktop\\exercise\\List" + i + "-A.docx");
      WordApp wordAppQ = new WordApp(i);
      wordAppQ.writeq(text, line, pline,
          "C:\\Users\\lzlghvbn\\Desktop\\exercise\\List" + i + "-Q.docx");
      System.out.println(i + " Succeeded");
    }
  }

}
