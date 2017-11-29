package tech.simter.jxls.ext;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import javax.ws.rs.core.Response;
import javax.ws.rs.core.StreamingOutput;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.time.OffsetDateTime;
import java.util.Map;

/**
 * The Excel Utils.
 *
 * @author RJ
 */
public class ExcelUtils {
  /**
   * Render the excel template with the specified data to the {@link OutputStream}.
   *
   * @param template the excel template, can be xlsx or xls format
   * @param target   the output target
   * @param data     the data
   * @throws ExcelException if has IOException inner
   */
  public static void renderTemplate(InputStream template, Map<String, Object> data, OutputStream target) {
    // Convert to jxls Context
    Context context = convert2Context(data);

    // Add default functions
    addDefault(context);

    // render
    try {
      JxlsHelper.getInstance().processTemplate(template, target, context);
    } catch (IOException e) {
      throw new ExcelException(e.getMessage(), e);
    }
  }

  private static Context convert2Context(Map<String, Object> data) {
    Context context = new Context();
    if (data != null) data.forEach(context::putVar);
    return context;
  }

  private static void addDefault(Context context) {
    // Current timestamp
    if (context.getVar("ts") == null) context.putVar("ts", OffsetDateTime.now());

    // Add default functions
    if (context.getVar("fn") == null) context.putVar("fn", Functions.getSingleton());
  }

  /**
   * Generate a {@link Response.ResponseBuilder} instance
   * and render the excel template with the specified data to its output stream.
   *
   * @param template the excel template, can be xlsx or xls format
   * @param data     the data
   * @return the instance of {@link Response.ResponseBuilder} with the excel data
   * @throws ExcelException if has IOException or UnsupportedEncodingException inner
   */
  public static Response.ResponseBuilder renderTemplate2Response(InputStream template, Map<String, Object> data) {
    return renderTemplate2Response(template, data, null);
  }

  /**
   * Generate a {@link Response.ResponseBuilder} instance
   * and render the excel template with the specified data to its output stream.
   *
   * @param template the excel template, can be xlsx or xls format
   * @param data     the data
   * @param filename the download filename of the response
   * @return the instance of {@link Response.ResponseBuilder} with the excel data
   * @throws ExcelException if has IOException or UnsupportedEncodingException inner
   */
  public static Response.ResponseBuilder renderTemplate2Response(InputStream template, Map<String, Object> data,
                                                                 String filename) {
    StreamingOutput stream = (OutputStream output) -> {
      // Convert to jxls Context
      Context context = convert2Context(data);

      // Add default functions
      addDefault(context);

      // render
      JxlsHelper.getInstance().processTemplate(template, output, context);
    };

    // create response
    Response.ResponseBuilder builder = Response.ok(stream);
    if (filename != null) {
      try {
        builder.header("Content-Disposition", "attachment; filename=\"" + URLEncoder.encode(filename, "UTF-8") + "\"");
      } catch (UnsupportedEncodingException e) {
        throw new ExcelException(e.getMessage(), e);
      }
    }
    return builder;
  }
}
