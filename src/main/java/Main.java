import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

public class Main {

  @SuppressWarnings("unchecked")
  public static void main(String[] args) {
    File path = new File("");
    try {
      String filePath = path.getCanonicalPath();
      File[] files = new File(filePath + "/json/").listFiles();
      int i = 0;
      if (files != null) {
        for (File file : files) {
          String name = file.getName();
          String suffix = name.substring(name.lastIndexOf(".") + 1);
          if ("json".equals(suffix)) {
            String newFile = filePath + "/excel/" + name.replace(".json", ".xlsx");
            json2Excel(file, newFile);
            System.out.println("生成文件:" + newFile);
            i++;
          }
        }
      }
      System.out.println("处理完成, 生成文件" + i + "个");
    } catch (IOException e) {
      System.out.println("路径错误");
      e.printStackTrace();
    }
  }

  /**
   * 转化excel
   *
   * @param file json文件
   * @param outPath excel格式文件输出路径
   */
  @SuppressWarnings("unchecked")
  private static void json2Excel(File file, String outPath) {
    String jsonStr;
    try {
      FileReader fileReader = new FileReader(file);
      InputStreamReader reader =
          new InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8);
      int ch;
      StringBuilder sb = new StringBuilder();
      while ((ch = reader.read()) != -1) {
        sb.append((char) ch);
      }
      // 关闭流
      fileReader.close();
      reader.close();

      // 转换为对象
      jsonStr = sb.toString();
      JSONArray array = (JSONArray) JSON.parse(jsonStr);
      //
      Map<String, String> item = (Map<String, String>) array.get(0);
      Set<String> strings = item.keySet();

      XSSFWorkbook workbook = new XSSFWorkbook();

      XSSFSheet sheet = workbook.createSheet("来自于json数据");

      XSSFRow row = sheet.createRow(0);

      int i = 0;
      List<String> titles = new ArrayList<>();
      for (String title : strings) {
        XSSFCell cell = row.createCell(i);
        cell.setCellValue(title);
        titles.add(title);
        i++;
      }

      int j = 1;
      for (Object o : array) {
        Map<String, Object> map = (Map<String, Object>) o;
        row = sheet.createRow(j);
        int k = 0;
        for (String title : titles) {
          XSSFCell cell = row.createCell(k);
          Object value = map.get(title);
          if (value instanceof JSONArray) {
            if (((JSONArray) value).size() != 0) {
              for (Object e : ((JSONArray) value).toArray()) {
                cell.setCellValue(replaceBlank(removeBrace(format(e.toString()))));
              }
            } else {
              cell.setCellValue("");
            }
          }
          if (value instanceof JSONObject) {
            cell.setCellValue(replaceBlank(removeBrace(format(value.toString()))));
          }
          if (value instanceof String) {
            cell.setCellValue(replaceBlank(removeBrace((String) value)));
          }

          k++;
        }
        j++;
      }

      File outputFile = new File(outPath);
      FileOutputStream outputStream = new FileOutputStream(outputFile);

      workbook.write(outputStream);
    } catch (IOException e) {
      System.out.println("读取出错");
    }
  }

  /**
   * 去掉大括号
   *
   * @param content 原来字符串
   * @return 处理后字符串
   */
  private static String removeBrace(String content) {
    return content.replaceAll("}", "").replaceAll("\\{", "");
  }

  /**
   * 去掉所有空格\t、回车\n、换行符\r、制表符\t
   *
   * @param content 原来字符串
   * @return 处理后字符串
   */
  private static String replaceBlank(String content) {
    return Pattern.compile("\\s*|\t|\r|\n").matcher(content).replaceAll("");
  }

  /**
   * 美化json格式
   *
   * @param content 原来字符串
   * @return 处理后字符串
   */
  private static String format(String content) {
    Object jsonObject = JSONObject.parse(content);
    return JSON.toJSONString(jsonObject, true);
  }
}
