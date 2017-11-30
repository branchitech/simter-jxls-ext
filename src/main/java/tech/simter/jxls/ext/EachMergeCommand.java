package tech.simter.jxls.ext;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.jxls.area.Area;
import org.jxls.command.CellRefGenerator;
import org.jxls.command.EachCommand;
import org.jxls.common.*;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * The {@link EachCommand} extension class to support merge cells.
 *
 * @author RJ
 */
public class EachMergeCommand extends EachCommand {
  private static Logger logger = LoggerFactory.getLogger(EachMergeCommand.class);

  public static final String COMMAND_NAME = "each-merge";

  public EachMergeCommand() {
    super();
  }

  public EachMergeCommand(String var, String items, Direction direction) {
    super(var, items, direction);
  }

  public EachMergeCommand(String items, Area area) {
    super(items, area);
  }

  public EachMergeCommand(String var, String items, Area area) {
    super(var, items, area);
  }

  public EachMergeCommand(String var, String items, Area area, Direction direction) {
    super(var, items, area, direction);
  }

  public EachMergeCommand(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
    super(var, items, area, cellRefGenerator);
  }

  @Override
  @SuppressWarnings("unchecked")
  public Size applyAt(CellRef cellRef, Context context) {
    logger.info("applyAt in each-merge command: {}", cellRef);

    // 子命令区域收集
    List<Area> childAreas = this.getAreaList().stream()
      .flatMap(area1 -> area1.getCommandDataList().stream())
      .flatMap(commandData -> commandData.getCommand().getAreaList().stream())
      .collect(Collectors.toList());
    List<AreaRef> childAreaRefs = childAreas.stream()
      .map(Area::getAreaRef).collect(Collectors.toList());

    // 主区域注册监听器
    Area parentArea = this.getAreaList().get(0);
    MergeCellListener listener = new MergeCellListener(getTransformer(), parentArea.getAreaRef(), childAreaRefs,
      ((Collection) context.getVar(this.getItems())).size()); // 总数据量
    logger.info("register listener {} to {} from {}", listener, parentArea.getAreaRef(), cellRef);
    parentArea.addAreaListener(listener);

    // 所有子区域注册监听器
    childAreas.forEach(area -> {
      logger.info("register listener {} to {} by parent", listener, area.getAreaRef());
      area.addAreaListener(listener);
    });

    // standard dealing
    return super.applyAt(cellRef, context);
  }

  public static class MergeCellListener implements AreaListener {
    private final PoiTransformer transformer;
    private int parentStartColumn;    // 主命令区域的开始列号
    private int[] childStartColumns;  // 所有子命令区域的开始列号
    private final int[] mergeColumns; // 要合并单元格的列
    private final List<int[]> records = new ArrayList<>(); // 0-开始行的索引号、1-结束行的索引号
    private int parentCount;

    private int childRow;
    private int parentProcessed;
    private String sheetName;

    MergeCellListener(Transformer transformer, AreaRef parent, List<AreaRef> children, int parentCount) {
      this.transformer = (PoiTransformer) transformer;
      this.parentCount = parentCount;
      this.parentStartColumn = parent.getFirstCellRef().getCol();

      // 子命令区域的所有列号混杂在一起
      int[] childCols = children.stream()
        .flatMapToInt(ref -> IntStream.range(ref.getFirstCellRef().getCol(), ref.getLastCellRef().getCol() + 1))
        .distinct().sorted()
        .toArray();

      // 各子命令区域的开始列号
      this.childStartColumns = children.stream()
        .mapToInt(ref -> ref.getFirstCellRef().getCol())
        .distinct().sorted().toArray();

      // 过滤得到需要合并的单元格的列号
      this.mergeColumns = IntStream.range(parent.getFirstCellRef().getCol(), parent.getLastCellRef().getCol() + 1)
        .filter(parentCol -> IntStream.of(childCols).noneMatch(childCol -> childCol == parentCol))
        .toArray();

      logger.debug("parent={}", parent);
      logger.debug("parentStartColumn={}", parentStartColumn);
      logger.debug("childStartColumns={}", childStartColumns);
      logger.debug("mergeColumns={}", mergeColumns);
      logger.debug("childCols={}", childCols);
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

      if (targetCell.getCol() == parentStartColumn) { // 记录合并区域
        this.parentProcessed++;

        logger.debug("parent: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
          targetCell.getRow(), targetCell.getCol());

        //-- 子集合多于 1 条时才需要合并
        if (targetCell.getRow() < this.childRow) this.records.add(new int[]{targetCell.getRow(), this.childRow});

        // 处理完最后那条数据时才执行合并单元格操作
        if (this.parentProcessed == this.parentCount) {
          Workbook workbook = transformer.getWorkbook();
          Sheet sheet = workbook.getSheet(sheetName);
          doMerge(sheet, this.records, this.mergeColumns, srcCell);
        }
        this.childRow = 0;

        // 记录 childStartColumns 当前所在的行号
      } else if (IntStream.of(childStartColumns).anyMatch(col -> col == targetCell.getCol())) {
        this.childRow = Math.max(this.childRow, targetCell.getRow());

        logger.debug("child: srcCell={}, targetCell={} [{}, {}]", srcCell, targetCell,
          targetCell.getRow(), targetCell.getCol());
      }
    }

    private static void doMerge(Sheet sheet, List<int[]> records, int[] mergeColumns, CellRef srcCell) {
      if (logger.isDebugEnabled()) {
        logger.debug("merge: sheetName={}, records={}", sheet.getSheetName(),
          records.stream().map(startEnd -> "[" + startEnd[0] + "," + startEnd[1] + "]")
            .collect(Collectors.joining(",")));
      }
      records.forEach(startEnd -> merge4Row(sheet, startEnd[0], startEnd[1], mergeColumns, srcCell));
    }

    private static void merge4Row(Sheet sheet, int fromRow, int toRow, int[] mergeColumns, CellRef srcCell) {
      if (fromRow >= toRow) {
        logger.warn("No need to merge because same row：fromRow={}, toRow={}", fromRow, toRow);
        return;
      }
      Cell originCell;
      CellStyle originCellStyle;
      CellRangeAddress region;
      for (int col : mergeColumns) {
        logger.debug("fromRow={}, toRow={}, col={}", fromRow, toRow, col);
        region = new CellRangeAddress(fromRow, toRow, col, col);
        sheet.addMergedRegion(region);

        //firstCell = sheet.getRow(fromRow).getCell(col);
        originCell = sheet.getRow(srcCell.getRow()).getCell(col);
        if (originCell == null) {
          logger.info("Missing cell: row={}, col={}", fromRow, col);
        }
        if (originCell != null) {
          // copy originCell style to the merged cell
          originCellStyle = originCell.getCellStyle();
          RegionUtil.setBorderTop(originCellStyle.getBorderTopEnum(), region, sheet);
          RegionUtil.setBorderRight(originCellStyle.getBorderRightEnum(), region, sheet);
          RegionUtil.setBorderBottom(originCellStyle.getBorderBottomEnum(), region, sheet);
          RegionUtil.setBorderLeft(originCellStyle.getBorderLeftEnum(), region, sheet);
        } else {
          RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
          RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        }
      }
    }
  }
}