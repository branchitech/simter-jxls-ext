package tech.simter.jxls.ext;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Jxls {@link AreaListener} for monitor region to merge.
 *
 * @author RJ
 */
public class MergeAreaListener implements AreaListener {
  private static Logger logger = LoggerFactory.getLogger(MergeAreaListener.class);
  private final PoiTransformer transformer;
  private int parentColumn;
  private int childColumn;
  private final int[] mergeColumns;
  private final List<int[]> records = new ArrayList<>(); // 0-开始行的索引号、1-结束行的索引号
  private int parentStartRow;
  private int parentCount;

  private int childRow;
  private int parentProcessed;
  private String sheetName;

  public MergeAreaListener(Transformer transformer, int parentColumn, int childColumn, int[] mergeColumns) {
    this.transformer = (PoiTransformer) transformer;
    this.parentColumn = parentColumn;
    this.childColumn = childColumn;
    this.mergeColumns = mergeColumns;
  }

  public void setParentStartRow(int parentStartRow) {
    this.parentStartRow = parentStartRow;
  }

  public void setParentCount(int parentCount) {
    this.parentCount = parentCount;
  }

  @Override
  public void beforeApplyAtCell(CellRef cellRef, Context context) {
  }

  @Override
  public void afterApplyAtCell(CellRef cellRef, Context context) {
  }

  @Override
  public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
  }

  @Override
  public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
    if (parentProcessed == 0) this.sheetName = targetCell.getSheetName();
    // 利用子命令监听器先于主命令监听器执行这个特征

    if (targetCell.getCol() == childColumn) {         // 记录 childColumn 当前所在的行号
      this.childRow = targetCell.getRow();

      logger.debug("child: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
        targetCell.getRow(), targetCell.getCol());
    } else if (targetCell.getCol() == parentColumn) { // 记录合并区域
      this.parentProcessed++;

      logger.debug("parent: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
        targetCell.getRow(), targetCell.getCol());

      //-- 子集合多于 1 条时才需要合并
      if (targetCell.getRow() < this.childRow) this.records.add(new int[]{targetCell.getRow(), this.childRow});

      // 处理完最后那条数据时才执行合并单元格操作
      if (this.parentProcessed == this.parentCount) {
        Workbook workbook = transformer.getWorkbook();
        Sheet sheet = workbook.getSheet(sheetName);
        doMerge(sheet, this.records, this.mergeColumns);
      }
    }
  }

  private static void doMerge(Sheet sheet, List<int[]> records, int[] mergeColumns) {
    if (logger.isDebugEnabled()) {
      logger.debug("merge: sheetName={}, records={}", sheet.getSheetName(),
        records.stream().map(startEnd -> "[" + startEnd[0] + "," + startEnd[1] + "]")
          .collect(Collectors.joining(",")));
    }
    records.forEach(startEnd -> merge4Row(sheet, startEnd[0], startEnd[1], mergeColumns));
  }

  private static void merge4Row(Sheet sheet, int fromRow, int toRow, int[] mergeColumns) {
    if (fromRow >= toRow) {
      logger.warn("相同的起始和结束行号无需合并：fromRow={}, toRow={}", fromRow, toRow);
      return;
    }
    Cell firstCell;
    CellStyle firstCellStyle;
    CellRangeAddress region;
    for (int col : mergeColumns) {
      logger.debug("fromRow={}, toRow={}, col={}", fromRow, toRow, col);
      region = new CellRangeAddress(fromRow, toRow, col, col);
      sheet.addMergedRegion(region);

      firstCell = sheet.getRow(fromRow).getCell(col);
      if (firstCell == null) {
        logger.info("Missing cell: row={}, col={}", fromRow, col);
      }
      if (firstCell != null) {
        firstCellStyle = sheet.getRow(fromRow).getCell(col).getCellStyle();
        RegionUtil.setBorderTop(firstCellStyle.getBorderTopEnum(), region, sheet);
        RegionUtil.setBorderRight(firstCellStyle.getBorderRightEnum(), region, sheet);
        RegionUtil.setBorderBottom(firstCellStyle.getBorderBottomEnum(), region, sheet);
        RegionUtil.setBorderLeft(firstCellStyle.getBorderLeftEnum(), region, sheet);
      } else {

        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
      }
    }
  }
}