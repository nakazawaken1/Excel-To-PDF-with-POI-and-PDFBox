package app;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Iterator;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextBox;

/**
 * Excel to PDF entry point class
 * 
 * @author nakazawaken1
 */
public class Main {

    /**
     * テスト出力
     */
    private static final PrintStream debug = System.out;

    /**
     * get stream from iterator
     * 
     * @param iterator iterator
     * @param size count of item or 0(unknown size)
     * @return stream
     */
    public static <T> Stream<T> stream(Iterator<T> iterator, long size) {
        return StreamSupport.stream(size > 0 ? Spliterators.spliterator(iterator, size, Spliterator.ORDERED | Spliterator.NONNULL)
                : Spliterators.spliteratorUnknownSize(iterator, Spliterator.ORDERED | Spliterator.NONNULL), false);
    }

    /**
     * get cell value
     * 
     * @param cell cell
     * @return cell value
     */
    @SuppressWarnings("deprecation")
    public static Optional<Object> cellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
        case NUMERIC:
            return Optional.ofNullable(cell.getNumericCellValue());
        case STRING:
            return Optional.ofNullable(cell.getRichStringCellValue());
        case FORMULA:
            return Optional.ofNullable(cell.getCellFormula());
        case BOOLEAN:
            return Optional.ofNullable(cell.getBooleanCellValue());
        case ERROR:
            return Optional.ofNullable(cell.getErrorCellValue());
        default:
            return Optional.empty();
        }
    }

    /**
     * excel to pdf
     * 
     * @param book excel workbook
     * @param out output to pdf
     */
    public static void convert(Workbook book, OutputStream out) {
        book.sheetIterator().forEachRemaining(sheet -> {
            debug.println("sheet name: " + sheet.getSheetName());
            debug.println("rows: " + sheet.getLastRowNum());
            debug.println("columns: " + stream(sheet.rowIterator(), sheet.getPhysicalNumberOfRows()).mapToInt(Row::getLastCellNum).max().orElse(0));
            sheet.rowIterator().forEachRemaining(row -> {
                row.cellIterator().forEachRemaining(cell -> {
                    cellValue(cell).ifPresent(
                            value -> debug.println(CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1) + ": " + value));
                });
            });
            sheet.getCellComments().entrySet().forEach(entry -> {
                debug.println("[comment] " + entry.getKey() + ": " + entry.getValue());
            });
            if(sheet instanceof XSSFSheet) {
                ((XSSFSheet)sheet).getDrawingPatriarch().getShapes().iterator().forEachRemaining(shape -> {
                    if(shape instanceof XSSFSimpleShape) {
                        debug.println("[shape text] " + ((XSSFSimpleShape)shape).getText());
                    }
                });
            } else if(sheet instanceof HSSFSheet) {
                ((HSSFSheet)sheet).getDrawingPatriarch().getChildren().iterator().forEachRemaining(shape -> {
                    if(shape instanceof HSSFSimpleShape) {
                        debug.println("[shape text] " + ((HSSFSimpleShape)shape).getString());
                    }
                });
            }
        });
    }

    /**
     * change file extension
     * 
     * @param path file path
     * @param newExtension new extension
     * @return changed file path
     */
    public static String changeExtension(String path, String newExtension) {
        if (path == null) {
            return null;
        }
        int i = path.lastIndexOf('.');
        if (i < 0) {
            return path + newExtension;
        }
        return path.substring(0, i) + newExtension;
    }

    /**
     * entry point
     * 
     * @param args Excel files(.xls, .xlsx)
     */
    public static void main(String[] args) {
        debug.println("processed " + Stream.of(args).peek(path -> {
            String toPath = changeExtension(path, ".pdf");
            try (InputStream in = Files.newInputStream(Paths.get(path), StandardOpenOption.READ);
                    Workbook book = WorkbookFactory.create(in);
                    OutputStream out = Files.newOutputStream(Paths.get(toPath), StandardOpenOption.CREATE)) {
                debug.println("processing: " + path);
                convert(book, out);
                debug.println("converted: " + toPath);
                debug.println("--------");
            } catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
                e.printStackTrace();
            }
        }).count() + " files.");
    }

}
