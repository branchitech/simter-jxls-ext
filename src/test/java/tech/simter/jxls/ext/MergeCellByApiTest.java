package tech.simter.jxls.ext;

import org.junit.Test;
import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

/**
 * The Jxls merge-cell test.
 *
 * @author RJ
 */
public class MergeCellByApiTest {
  @Test
  public void merge() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/merge-cell-by-api.xlsx");

    // output to
    File out = new File("target/merge-cell-by-api-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    // template data
    Map<String, Object> data = generateData();

    // render
    render(template, output, data);

    // verify
    assertThat(out.getTotalSpace() > 0, is(true));
  }

  @SuppressWarnings("unchecked")
  private void render(InputStream template, OutputStream output, Map<String, Object> data) throws IOException {
    Transformer transformer = TransformerFactory.createTransformer(template, output);

    // data
    Context context = convert2Context(data);

    // 用于查询并记录合并单元格行号的监听器：监控 A、C 列，A、B 列需要合并单元格
    MergeAreaListener mergeListener = new MergeAreaListener(transformer, 0, 2, new int[]{0, 1});
    mergeListener.setParentStartRow(2); // 数据行开始行的索引
    mergeListener.setParentCount(((List<Map<String, Object>>) context.getVar("rows")).size()); // 数据量

    // 1. row
    String rowRef = "A3:D3";
    XlsArea rowArea = new XlsArea(buildAreaRef(rowRef), transformer);
    rowArea.addAreaListener(mergeListener); // A 列
    EachCommand rowEachCommand = new EachCommand("row", "rows", rowArea);

    // 1.1. sub
    String subRef = "C3:D3";
    XlsArea subArea = new XlsArea(buildAreaRef(subRef), transformer);
    subArea.addAreaListener(mergeListener); // C 列
    EachCommand subEachCommand = new EachCommand("sub", "row.subs", subArea);
    rowArea.addCommand(new AreaRef(buildAreaRef(subRef)), subEachCommand);

    // 2. main
    XlsArea xlsArea = new XlsArea(buildAreaRef("A1:D3"), transformer);
    xlsArea.addCommand(new AreaRef(buildAreaRef(rowRef)), rowEachCommand);

    // render
    xlsArea.applyAt(new CellRef(buildAreaRef("A1")), context);
    xlsArea.processFormulas();
    transformer.write();
  }

  private String buildAreaRef(String name) {
    return "Sheet2!" + name;
  }

  public static Context convert2Context(Map<String, Object> data) {
    Context context = new Context();
    if (data != null) data.forEach(context::putVar);
    return context;
  }

  public static Map<String, Object> generateData() {
    Map<String, Object> data = new HashMap<>();
    data.put("subject", "JXLS merge cell test");

    List<Map<String, Object>> rows = new ArrayList<>();
    data.put("rows", rows);
    int rowNumber = 0;
    //rows.add(createRow(++rowNumber, 3));
    rows.add(createRow(++rowNumber, 3));
    rows.add(createRow(++rowNumber, 1));
    rows.add(createRow(++rowNumber, 2));
    //rows.add(createRow(++rowNumber, 0));

    return data;
  }

  private static Map<String, Object> createRow(int rowNumber, int subsCount) {
    Map<String, Object> row = new HashMap<>();
    row.put("sn", rowNumber);
    row.put("name", "row" + rowNumber);
    if (subsCount >= 0) row.put("subs", createSubs(rowNumber, subsCount));
    return row;
  }

  private static List<Map<String, Object>> createSubs(int rowNumber, int count) {
    List<Map<String, Object>> subs = new ArrayList<>();
    Map<String, Object> sub;
    for (int i = 1; i <= count; i++) {
      sub = new HashMap<>();
      subs.add(sub);
      sub.put("sn", rowNumber + "-" + i);
      sub.put("name", "row" + rowNumber + "sub" + +i);
    }
    return subs;
  }
}