package tech.simter.jxls.ext;

import org.jxls.area.Area;
import org.jxls.command.CellRefGenerator;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * The {@link EachCommand} extension class to support config listener.
 *
 * @author RJ
 */
public class EachCommandWithAreaListener extends EachCommand {
  private static Logger logger = LoggerFactory.getLogger(EachCommandWithAreaListener.class);

  public EachCommandWithAreaListener() {
    super();
  }

  public EachCommandWithAreaListener(String var, String items, Direction direction) {
    super(var, items, direction);
  }

  public EachCommandWithAreaListener(String items, Area area) {
    super(items, area);
  }

  public EachCommandWithAreaListener(String var, String items, Area area) {
    super(var, items, area);
  }

  public EachCommandWithAreaListener(String var, String items, Area area, Direction direction) {
    super(var, items, area, direction);
  }

  public EachCommandWithAreaListener(String var, String items, Area area, CellRefGenerator cellRefGenerator) {
    super(var, items, area, cellRefGenerator);
  }

  private String listener;

  public String getListener() {
    return listener;
  }

  public void setListener(String listener) {
    this.listener = listener;
  }

  private boolean registered;

  @Override
  public Size applyAt(CellRef cellRef, Context context) {
    if (!registered) { // only register once
      AreaListener listener = createAreaListener(context);
      if (listener != null) this.getAreaList().forEach(area -> {
        area.addAreaListener(listener);
        logger.info("register listener {} to {} {}", listener, cellRef, area.getAreaRef());
      });
      registered = true;
    }
    return super.applyAt(cellRef, context);
  }

  private AreaListener createAreaListener(Context context) {
    if (listener == null || listener.isEmpty()) return null;
    Object value = context.getVar(listener);
    if (!(value instanceof AreaListener))
      throw new IllegalArgumentException("The listener attribute value '" + listener +
        "' should be set to the instance of AreaListener");
    else return (AreaListener) value;
  }
}
