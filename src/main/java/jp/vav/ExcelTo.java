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
import java.util.function.Consumer;
import java.util.logging.Logger;

import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
                sheet.rowIterator().forEachRemaining(row -> {
                    row.cellIterator().forEachRemaining(cell -> {
                        Tool.cellValue(cell).ifPresent(value -> printer.println(new CellReference(cell).formatAsString() + ": " + value));
                    });
                });
                sheet.getCellComments().entrySet().forEach(entry -> {
                    printer.println("[comment] " + entry.getKey() + ": " + entry.getValue().getString());
                });
                if (sheet instanceof XSSFSheet) {
                    Optional.ofNullable(((XSSFSheet) sheet).getDrawingPatriarch())
                            .ifPresent(drawing -> drawing.getShapes().iterator().forEachRemaining(shape -> {
                                if (shape instanceof XSSFSimpleShape) {
                                    try {
                                        printer.println("[shape text] " + ((XSSFSimpleShape) shape).getText());
                                    } catch (NullPointerException e) {
                                        /* NullPointerException occurs depending on shape type */
                                    }
                                }
                            }));
                } else if (sheet instanceof HSSFSheet) {
                    Optional.ofNullable(((HSSFSheet) sheet).getDrawingPatriarch())
                            .ifPresent(drawing -> drawing.getChildren().iterator().forEachRemaining(shape -> {
                                if (shape instanceof HSSFSimpleShape) {
                                    try {
                                        printer.println("[shape text] " + ((HSSFSimpleShape) shape).getString());
                                    } catch (NullPointerException e) {
                                        /* NullPointerException occurs depending on shape type */
                                    }
                                }
                            }));
                }
                printer.newPage();
            }
            printer.getDocument().save(out);
        }
    }

    /**
     * Excel to text
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
                sheet.rowIterator().forEachRemaining(row -> {
                    row.cellIterator().forEachRemaining(cell -> {
                        Tool.cellValue(cell).ifPresent(value -> printer.println(new CellReference(cell).formatAsString() + ": " + value));
                    });
                });
                sheet.getCellComments().entrySet().forEach(entry -> {
                    printer.println("[comment] " + entry.getKey() + ": " + entry.getValue().getString());
                });
                if (sheet instanceof XSSFSheet) {
                    Optional.ofNullable(((XSSFSheet) sheet).getDrawingPatriarch())
                            .ifPresent(drawing -> drawing.getShapes().iterator().forEachRemaining(shape -> {
                                if (shape instanceof XSSFSimpleShape) {
                                    try {
                                        printer.println("[shape text] " + ((XSSFSimpleShape) shape).getText());
                                    } catch (NullPointerException e) {
                                        /* NullPointerException occurs depending on shape type */
                                    }
                                }
                            }));
                } else if (sheet instanceof HSSFSheet) {
                    Optional.ofNullable(((HSSFSheet) sheet).getDrawingPatriarch())
                            .ifPresent(drawing -> drawing.getChildren().iterator().forEachRemaining(shape -> {
                                if (shape instanceof HSSFSimpleShape) {
                                    try {
                                        printer.println("[shape text] " + ((HSSFSimpleShape) shape).getString());
                                    } catch (NullPointerException e) {
                                        /* NullPointerException occurs depending on shape type */
                                    }
                                }
                            }));
                }
                printer.println("--------");
            }
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
