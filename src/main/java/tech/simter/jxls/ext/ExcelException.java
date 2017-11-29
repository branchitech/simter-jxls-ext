package tech.simter.jxls.ext;

/**
 * The Excel Exception.
 *
 * @author RJ
 */
public class ExcelException extends RuntimeException {
  public ExcelException() {
    super();
  }

  public ExcelException(String message) {
    super(message);
  }

  public ExcelException(Throwable cause) {
    super(cause);
  }

  public ExcelException(String message, Throwable cause) {
    super(message, cause);
  }
}
