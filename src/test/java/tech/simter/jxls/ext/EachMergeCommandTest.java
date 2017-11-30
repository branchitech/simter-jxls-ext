package tech.simter.jxls.ext;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.area.CommandData;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;
import static tech.simter.jxls.ext.MergeCellByApiTest.convert2Context;
import static tech.simter.jxls.ext.MergeCellByApiTest.generateData;

/**
 * The Jxls merge-cell test through comment.
 *
 * @author RJ
 */
public class EachMergeCommandTest {
  private static Logger logger = LoggerFactory.getLogger(EachMergeCommandTest.class);

  @Test
  @SuppressWarnings("unchecked")
  public void merge() throws Exception {
    // template
    InputStream template = getClass().getClassLoader().getResourceAsStream("templates/each-merge.xlsx");

    // output to
    File out = new File("target/each-merge-result.xlsx");
    if (out.exists()) out.delete();
    OutputStream output = new FileOutputStream(out);

    Transformer transformer = TransformerFactory.createTransformer(template, output);

    // generate template data
    Map<String, Object> data = generateData();
//    MergeAreaListener mergeListener = new MergeAreaListener(transformer, 0, 2, new int[]{0, 1});
//    mergeListener.setParentStartRow(2); // 数据行开始行的索引
//    mergeListener.setParentCount(((List<Map<String, Object>>) data.get("rows")).size()); // 数据量
//    data.put("mergeCellListener", mergeListener);

    // convert data to context
    Context context = convert2Context(data);

    // get comment area
    XlsCommentAreaBuilder.addCommandMapping(EachMergeCommand.COMMAND_NAME, EachMergeCommand.class);
    AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
    List<Area> xlsAreas = areaBuilder.build();
    logger.info("xlsAreas.size={}", xlsAreas.size());

    // test
    logger.info("test0");
    xlsAreas.forEach(area -> printArea(area, 1));

    // inject listener
    // 用于查询并记录合并单元格行号的监听器：监控 A、C 列，A、B 列需要合并单元格
    //logger.info("inject listener");
    //MergeAreaListener mergeListener = new MergeAreaListener(transformer, 0, 2, new int[]{0, 1});
    //mergeListener.setParentStartRow(2); // 数据行开始行的索引
    //mergeListener.setParentCount(((List<Map<String, Object>>) context.getVar("rows")).size()); // 数据量
    //xlsAreas.forEach(area -> injectListener(area, mergeListener));

    // test
    //logger.info("test1");
    //xlsAreas.forEach(area -> printArea(area, 1));

    // render
    for (Area xlsArea : xlsAreas) {
      xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
    }
    transformer.write();

    // verify
    assertThat(out.getTotalSpace() > 0, is(true));
  }

  CellRef A3 = new CellRef("Sheet1!A3");
  CellRef C3 = new CellRef("Sheet1!C3");

  private void injectListener(Area area, MergeAreaListener listener) {
    for (CommandData cd : area.getCommandDataList()) {
      CellRef firstCellRef = cd.getAreaRef().getFirstCellRef();
      if (firstCellRef.equals(A3) || firstCellRef.equals(C3)) {
        cd.getCommand().getAreaList().forEach(a -> a.addAreaListener(listener));
      }

      // recursive
      cd.getCommand().getAreaList().forEach(a -> injectListener(a, listener));
    }
  }

  private void printArea(Area area, int level) {
    logger.debug("{}area={}", String.join("", Collections.nCopies((level - 1) * 2, " ")), area.getAreaRef());
    for (CommandData cd : area.getCommandDataList()) {
      logger.debug("{}command={}, ref={}", String.join("", Collections.nCopies(level * 2, " ")),
        cd.getCommand().getName(), cd.getAreaRef());

      // recursive
      cd.getCommand().getAreaList().forEach(a -> printArea(a, level + 2));
    }
    for (AreaListener listener : area.getAreaListeners()) {
      logger.debug("{}listener={}", String.join("", Collections.nCopies(level * 2, " ")),
        listener.getClass().getSimpleName());
    }
  }
}