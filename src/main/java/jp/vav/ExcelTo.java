/*
Copyright 2016 nakazawaken1

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
package jp.vav;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.logging.Logger;

import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

/**
 * Entry point class
 * 
 * @author nakazawaken1
 */
public class ExcelTo {

    /**
     * logger
     */
    private static final Logger logger = Logger.getLogger(ExcelTo.class.getCanonicalName());

    /**
     * excel to pdf
     * 
     * @param book excel workbook
     * @param out output to pdf
     * @param documentSetup document setup
     * @throws IOException I/O exception
     */
    public static void pdf(Workbook book, OutputStream out, Consumer<PDFPrinter> documentSetup) throws IOException {
        Objects.requireNonNull(book);
        Objects.requireNonNull(out);
        try (PDFPrinter printer = new PDFPrinter()) {
            printer.documentSetup = Optional.ofNullable(documentSetup);
            for (int i = 0, end = book.getNumberOfSheets(); i < end; i++) {
                Sheet sheet = book.getSheetAt(i);
                int rowCount = sheet.getPhysicalNumberOfRows();
                if (rowCount <= 0) {
                    logger.info(sheet.getSheetName() + ": empty");
                    continue; /* skip blank sheet */
                }
                logger.info(sheet.getSheetName() + ": " + rowCount + " rows");
                printer.println("sheet name: " + sheet.getSheetName());
                printer.println("max row index: " + sheet.getLastRowNum());
                printer.println("max column index: " + Tool.stream(sheet.rowIterator(), rowCount).mapToInt(Row::getLastCellNum).max().orElse(0));
                eachCell(sheet, (cell, range) -> Tool.cellValue(cell).ifPresent(
                        value -> printer.println('[' + (range == null ? new CellReference(cell).formatAsString() : range.formatAsString()) + "] " + value)));
                eachShape(sheet, shapeText(text -> printer.println("[shape text] " + text)));
                printer.newPage();
            }
            printer.getDocument().save(out);
        }
    }

    /**
     * get shape text
     * 
     * @param consumer text consumer
     * @return binary consumer
     */
    public static BiConsumer<XSSFSimpleShape, HSSFSimpleShape> shapeText(Consumer<String> consumer) {
        return (shapeX, shapeH) -> {
            try {
                consumer.accept(shapeX != null ? shapeX.getText() : shapeH.getString().getString());
            } catch (NullPointerException e) {
                /* NullPointerException occurs depending on shape type */
            }
        };
    }

    /**
     * excel to text
     * 
     * @param book excel workbook
     * @param out output to text
     */
    public static void text(Workbook book, OutputStream out) {
        Objects.requireNonNull(book);
        Objects.requireNonNull(out);
        try (PrintStream printer = Try.to(() -> new PrintStream(out, true, System.getProperty("file.encoding"))).get()) {
            for (int i = 0, end = book.getNumberOfSheets(); i < end; i++) {
                Sheet sheet = book.getSheetAt(i);
                int rowCount = sheet.getPhysicalNumberOfRows();
                if (rowCount <= 0) {
                    logger.info(sheet.getSheetName() + ": empty");
                    continue; /* skip blank sheet */
                }
                logger.info(sheet.getSheetName() + ": " + rowCount + " rows");
                printer.println("sheet name: " + sheet.getSheetName());
                printer.println("max row index: " + sheet.getLastRowNum());
                printer.println("max column index: " + Tool.stream(sheet.rowIterator(), rowCount).mapToInt(Row::getLastCellNum).max().orElse(0));
                eachCell(sheet, (cell, range) -> Tool.cellValue(cell).ifPresent(
                        value -> printer.println('[' + (range == null ? new CellReference(cell).formatAsString() : range.formatAsString()) + "] " + value)));
                sheet.getCellComments().entrySet().forEach(entry -> {
                    printer.println("[comment " + entry.getKey() + "] " + entry.getValue().getString());
                });
                eachShape(sheet, shapeText(text -> printer.println("[shape text] " + text)));
                printer.println("--------");
            }
        }
    }

    /**
     * traverse all cells
     * 
     * @param sheet sheet
     * @param consumer cell consumer
     */
    public static void eachCell(Sheet sheet, BiConsumer<Cell, CellRangeAddress> consumer) {
        sheet.rowIterator().forEachRemaining(row -> {
            row.cellIterator().forEachRemaining(cell -> {
                int rowIndex = cell.getRowIndex();
                int columnIndex = cell.getColumnIndex();
                boolean until = true;
                for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
                    if (mergedRegion.isInRange(rowIndex, columnIndex)) {
                        if (rowIndex == mergedRegion.getFirstRow() && columnIndex == mergedRegion.getFirstColumn()) {
                            consumer.accept(cell, mergedRegion);
                        }
                        until = false;
                        break;
                    }
                }
                if (until) {
                    consumer.accept(cell, null);
                }
            });
        });
    }

    /**
     * traverse all shape
     * 
     * @param sheet sheet
     * @param consumer shape consumer
     */
    public static void eachShape(Sheet sheet, BiConsumer<XSSFSimpleShape, HSSFSimpleShape> consumer) {
        if (sheet instanceof XSSFSheet) {
            Optional.ofNullable(((XSSFSheet) sheet).getDrawingPatriarch()).ifPresent(drawing -> drawing.getShapes().forEach(shape -> {
                if (shape instanceof XSSFSimpleShape) {
                    consumer.accept((XSSFSimpleShape) shape, null);
                }
            }));
        } else if (sheet instanceof HSSFSheet) {
            Optional.ofNullable(((HSSFSheet) sheet).getDrawingPatriarch()).ifPresent(drawing -> drawing.getChildren().forEach(shape -> {
                if (shape instanceof HSSFSimpleShape) {
                    consumer.accept(null, (HSSFSimpleShape) shape);
                }
            }));
        }
    }

    /**
     * entry point
     * 
     * @param args [-p password] [-m true|false(draw margin line if true)] Excel files(.xls, .xlsx, .xlsm)
     */
    public static void main(String[] args) {
        Objects.requireNonNull(args);
        int count = 0;
        boolean[] drawMarginLine = { false };
        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
            case "-m":/* set draw margin line */
                i++;
                drawMarginLine[0] = Boolean.parseBoolean(args[i]);
                break;
            case "-p":/* set password */
                i++;
                Biff8EncryptionKey.setCurrentUserPassword(args[i]);
                break;
            default:
                String path = Tool.trim(args[i], "\"", "\"");
                String toPath = Tool.changeExtension(path, ".pdf");
                String toTextPath = Tool.changeExtension(path, ".txt");
                try (InputStream in = Files.newInputStream(Paths.get(path));
                        Workbook book = WorkbookFactory.create(in);
                        OutputStream out = Files.newOutputStream(Paths.get(toPath));
                        OutputStream outText = Files.newOutputStream(Paths.get(toTextPath))) {
                    logger.info("processing: " + path);
                    pdf(book, out, printer -> {
                        printer.setPageSize(PDRectangle.A4, false);
                        printer.setFont(System.getenv("WINDIR") + "\\fonts\\msgothic.ttc", "MS-Gothic"); /* for windows */
                        printer.setFontSize(10.5f);
                        printer.setMargin(15);
                        printer.setLineSpace(5);
                        printer.setDrawMarginLine(drawMarginLine[0]);
                    });
                    text(book, outText);
                    logger.info("converted: " + toPath + ", " + toTextPath);
                    count++;
                } catch (InvalidFormatException e) {
                    throw new RuntimeException("Invalid file type: " + path);
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
                break;
            }
        }
        logger.info("processed " + count + " files.");
    }

}
