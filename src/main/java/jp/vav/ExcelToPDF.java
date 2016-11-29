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

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.Supplier;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.fontbox.ttf.TTFParser;
import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import com.github.mygreen.cellformatter.POICellFormatter;

/**
 * Entry point class
 * 
 * @author nakazawaken1
 */
public class ExcelToPDF {

    /**
     * logger
     */
    private static final Logger logger = Logger.getLogger(ExcelToPDF.class.getCanonicalName());

    /**
     * throwsable runnable
     */
    public interface ThrowableRunnable {
        /**
         * @throws Exception Exception
         */
        void run() throws Exception;
    }

    /**
     * throwsable consumer
     * 
     * @param <T> consume type
     */
    public interface ThrowableConsumer<T> {
        /**
         * @param value consume value
         * @throws Exception Exception
         */
        void accept(T value) throws Exception;
    }

    /**
     * throwsable supplier
     * 
     * @param <T> supply type
     */
    public interface ThrowableSupplier<T> {
        /**
         * @return supply value
         * @throws Exception Exception
         */
        T get() throws Exception;
    }

    /**
     * throwsable function
     * 
     * @param <T> argument type
     * @param <R> result type
     */
    public interface ThrowableFunction<T, R> {
        /**
         * @param argument argument
         * @return supply value
         * @throws Exception Exception
         */
        R apply(T argument) throws Exception;
    }

    /**
     * throwable runnable
     * 
     * @param runnable throwable runnable
     * @return runnable
     */
    public static Runnable throwable(ThrowableRunnable runnable) {
        return () -> {
            try {
                runnable.run();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable supplier
     * 
     * @param consumer throwable consumer
     * @return consumer
     */
    public static <T> Consumer<T> throwable(ThrowableConsumer<T> consumer) {
        return value -> {
            try {
                consumer.accept(value);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable supplier
     * 
     * @param supplier throwable supplier
     * @return supplier
     */
    public static <T> Supplier<T> throwable(ThrowableSupplier<T> supplier) {
        return () -> {
            try {
                return supplier.get();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable function
     * 
     * @param function throwable function
     * @return function
     */
    public static <T, R> Function<T, R> throwable(ThrowableFunction<T, R> function) {
        return argument -> {
            try {
                return function.apply(argument);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

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
    public static Optional<String> cellValue(Cell cell) {
        return Optional.ofNullable(new POICellFormatter().formatAsString(cell, Locale.JAPAN)).filter(((Predicate<String>) String::isEmpty).negate());
    }

    /**
     * @param text target text
     * @param left trimming characters of left(no operation when null or empty)
     * @param right trimming characters of right(no operation when null or empty)
     * @return trimmed text
     */
    public static String trim(String text, String left, String right) {
        int begin = 0;
        int end = text.length();
        if (left != null && !left.isEmpty()) {
            while (begin < end && left.indexOf(text.charAt(begin)) >= 0) {
                begin++;
            }
        }
        if (right != null && !right.isEmpty()) {
            while (end > begin && right.indexOf(text.charAt(end - 1)) >= 0) {
                end--;
            }
        }
        return text.substring(begin, end);
    }

    /**
     * PDF printer
     */
    public static class Printer implements AutoCloseable {
        /**
         * output
         */
        protected OutputStream out;

        /**
         * PDF page size
         */
        protected PDRectangle pageSize = PDRectangle.A4;

        /**
         * Font size
         */
        protected float fontSize = 10.5f;

        /**
         * font
         */
        protected PDFont font = PDType1Font.HELVETICA;

        /**
         * top margin
         */
        protected float marginTop = 10.5f;

        /**
         * bottom margin
         */
        protected float marginButtom = 10.5f;

        /**
         * left margin
         */
        protected float marginLeft = 10.5f;

        /**
         * right margin
         */
        protected float marginRight = 10.5f;

        /**
         * space between lines
         */
        protected float lineSpace = fontSize / 2;

        /**
         * PDF document
         */
        protected Optional<PDDocument> document = Optional.empty();

        /**
         * PDF contents
         */
        protected Optional<PDPageContentStream> content = Optional.empty();

        /**
         * true type font collection
         */
        protected Optional<TrueTypeCollection> ttc = Optional.empty();

        /**
         * true type font
         */
        protected Optional<TrueTypeFont> ttf = Optional.empty();

        /**
         * setup at document created
         */
        protected Optional<Consumer<Printer>> documentSetup = Optional.empty();

        /**
         * setup at content created
         */
        protected Optional<Consumer<Printer>> contentSetup = Optional.empty();

        /**
         * current horizontal position
         */
        protected float currentX;

        /**
         * current vertical position
         */
        protected float currentY;

        /**
         * draw margin line if true
         */
        protected boolean drawMarginLine;

        /**
         * create printer
         * 
         * @param out output to PDF
         * @return printer
         */
        public static Printer use(OutputStream out) {
            Objects.requireNonNull(out);
            Printer printer = new Printer();
            logger.config("create Printer: " + printer.hashCode());
            printer.out = out;
            return printer;
        }

        /**
         * print text
         * 
         * @param text text
         */
        public void print(String text) {
            if (text == null || text.isEmpty()) {
                return;
            }
            try {
                final String lineBreak = "\n";
                int index = text.indexOf(lineBreak);
                if (index >= 0) {
                    println(text.substring(0, index));
                    print(text.substring(index + lineBreak.length()));
                    return;
                }
                PDPageContentStream content = getContent();
                float width = textWidth(text); /* font.getStringWidth must be after getContent */
                float max = pageSize.getWidth() - marginRight - currentX;
                if (width > max) { /* wrap */
                    int length = text.length();
                    int count = length / 2;
                    while (count > 1 && textWidth(text.substring(0, count)) > max) {
                        count /= 2;
                    }
                    while (count < length && textWidth(text.substring(0, count + 1)) <= max) {
                        count++;
                    }
                    content.showText(text.substring(0, count));
                    newLine();
                    print(text.substring(count));
                    return;
                }
                content.showText(text);
                currentX += width;
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        }

        /**
         * print newline
         */
        public void newLine() {
            currentX = marginLeft;
            currentY += fontSize + lineSpace;
            if (currentY + fontSize > pageSize.getHeight() - marginButtom) {
                newPage();
            } else {
                throwable(() -> getContent().newLineAtOffset(0, -fontSize - lineSpace)).run();
            }
        }

        /**
         * print text and newline
         * 
         * @param text text
         */
        public void println(String text) {
            print(text);
            newLine();
        }

        /**
         * break page
         */
        public void newPage() {
            content.ifPresent(i -> {
                try {
                    i.endText();
                    if(drawMarginLine) {
                        i.addRect(marginLeft, marginButtom, pageSize.getWidth() - marginLeft - marginRight, pageSize.getHeight() - marginTop - marginButtom);
                        i.setLineWidth(0.01f);
                        i.setLineDashPattern(new float[] { 3f, 1f }, 0);
                        i.stroke();
                    }
                    i.close();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            });
            content.ifPresent(i -> logger.config("destroy PDPageContentStream: " + i.hashCode()));
            content = Optional.empty();
        }

        /*
         * (non-Javadoc)
         * 
         * @see java.lang.AutoCloseable#close()
         */
        @Override
        public void close() {
            newPage();
            document.ifPresent(throwable(i -> {
                i.save(out);
                i.close();
                ttf.ifPresent(throwable(TrueTypeFont::close));
                ttc.ifPresent(throwable(TrueTypeCollection::close));
            }));
            document.ifPresent(i -> logger.config("destroy PDDocument: " + i.hashCode()));
            logger.config("destroy Printer: " + hashCode());
        }

        /**
         * @return PDF document
         */
        protected PDDocument getDocument() {
            return document.orElseGet(() -> {
                PDDocument document = new PDDocument();
                logger.config("add PDDocument: " + document.hashCode());
                this.document = Optional.of(document);
                documentSetup.ifPresent(i -> i.accept(this));
                return document;
            });
        }

        /**
         * @return PDF content
         */
        protected PDPageContentStream getContent() {
            return content.orElseGet(throwable(() -> {
                PDPage page = new PDPage(pageSize);
                logger.config("add PDPage: " + page.hashCode());
                getDocument().addPage(page);
                PDPageContentStream content = new PDPageContentStream(getDocument(), page);
                logger.config("create PDPageContentStream: " + content.hashCode());
                this.content = Optional.of(content);
                content.beginText();
                content.setFont(font, fontSize);
                content.newLineAtOffset(marginLeft, pageSize.getHeight() - fontSize - marginTop);
                currentY = marginTop;
                contentSetup.ifPresent(i -> i.accept(this));
                return content;
            }));
        }

        /**
         * @return descent
         */
        protected float descent() {
            return font.getFontDescriptor().getDescent() * fontSize / 1000f;
        }

        /**
         * @param text text
         * @return text width
         */
        protected float textWidth(String text) {
            return text == null ? 0 : throwable(() -> font.getStringWidth(text) * fontSize / 1000f).get();
        }

        /**
         * @param point top, bottom, left and right margin
         */
        public void setMargin(float point) {
            marginTop = marginButtom = marginLeft = marginRight = point;
        }

        /**
         * @param vertical top and bottom margin
         * @param horizontal left and right margin
         */
        public void setMargin(float vertical, float horizontal) {
            marginTop = marginButtom = vertical;
            marginLeft = marginRight = horizontal;
        }

        /**
         * @param top top margin
         * @param bottom bottom margin
         * @param left left margin
         * @param right right margin
         */
        public void setMargin(float top, float bottom, float left, float right) {
            marginTop = top;
            marginButtom = bottom;
            marginLeft = left;
            marginRight = right;
        }

        /**
         * @param page page size
         * @param isLandscape true: landscape, false: portrait
         */
        public void setPageSize(PDRectangle page, boolean isLandscape) {
            Objects.requireNonNull(page);
            pageSize = isLandscape ? new PDRectangle(page.getHeight(), page.getWidth()) : page;
        }

        /**
         * @param font font
         */
        public void setFont(PDFont font) {
            Objects.requireNonNull(font);
            this.font = font;
        }

        /**
         * @param path .ttf file path
         */
        public void setFont(String path) {
            Objects.requireNonNull(path);
            setFont(throwable(() -> {
                ttf = Optional.of(new TTFParser().parse(new File(path)));
                return PDType0Font.load(getDocument(), ttf.get(), true);
            }).get());
        }

        /**
         * @param path .ttc file path
         * @param name font name
         */
        public void setFont(String path, String name) {
            Objects.requireNonNull(path);
            Objects.requireNonNull(name);
            setFont(throwable(() -> {
                ttc = Optional.of(new TrueTypeCollection(new File(path)));
                ttf = Optional.of(ttc.get().getFontByName(name));
                return PDType0Font.load(getDocument(), ttf.get(), true);
            }).get());
        }

        /**
         * @param point font size
         */
        public void setFontSize(float point) {
            fontSize = point;
        }

        /**
         * @param point space between lines
         */
        public void setLineSpace(float point) {
            lineSpace = point;
        }

        /**
         * @param enabled draw margin line if true
         */
        public void setDrawMarginLine(boolean enabled) {
            drawMarginLine = enabled;
        }
    }

    /**
     * excel to pdf
     * 
     * @param book excel workbook
     * @param out output to pdf
     * @param documentSetup document setup
     */
    public static void convert(Workbook book, OutputStream out, Consumer<Printer> documentSetup) {
        Objects.requireNonNull(book);
        Objects.requireNonNull(out);
        try (Printer printer = Printer.use(out)) {
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
                printer.println("max column index: " + stream(sheet.rowIterator(), rowCount).mapToInt(Row::getLastCellNum).max().orElse(0));
                sheet.rowIterator().forEachRemaining(row -> {
                    row.cellIterator().forEachRemaining(cell -> {
                        cellValue(cell).ifPresent(value -> printer.println(new CellReference(cell).formatAsString() + ": " + value));
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
        }
    }

    /**
     * Excel to text
     * 
     * @param book excel workbook
     * @param out output to text
     */
    public static void toText(Workbook book, OutputStream out) {
        Objects.requireNonNull(book);
        Objects.requireNonNull(out);
        try (PrintStream printer = throwable(() -> new PrintStream(out, true, System.getProperty("file.encoding"))).get()) {
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
                printer.println("max column index: " + stream(sheet.rowIterator(), rowCount).mapToInt(Row::getLastCellNum).max().orElse(0));
                sheet.rowIterator().forEachRemaining(row -> {
                    row.cellIterator().forEachRemaining(cell -> {
                        cellValue(cell).ifPresent(value -> printer.println(new CellReference(cell).formatAsString() + ": " + value));
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
     * change file extension
     * 
     * @param path file path
     * @param newExtension new extension
     * @return changed file path
     */
    public static String changeExtension(String path, String newExtension) {
        Objects.requireNonNull(path);
        Objects.requireNonNull(newExtension);
        int i = path.lastIndexOf('.');
        if (i < 0) {
            return path + newExtension;
        }
        return path.substring(0, i) + newExtension;
    }

    /**
     * entry point
     * 
     * @param args [-l logLevel(INFO, CONFIG, ...)] [-p password] Excel files(.xls, .xlsx, .xlsm)
     */
    public static void main(String[] args) {
        Objects.requireNonNull(args);
        int count = 0;
        boolean[] drawMarginLine = {false};
        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
            case "-l":/* set log level */
                i++;
                Level level = Level.parse(args[i]);
                for (Handler handler : Logger.getLogger("").getHandlers()) {
                    handler.setLevel(level);
                }
                break;
            case "-m":/* set draw margin line */
                i++;
                drawMarginLine[0] = Boolean.parseBoolean(args[i]);
               break;
            case "-p":/* set password */
                i++;
                Biff8EncryptionKey.setCurrentUserPassword(args[i]);
                break;
            default:
                String path = trim(args[i], "\"", "\"");
                String toPath = changeExtension(path, ".pdf");
                String toTextPath = changeExtension(path, ".txt");
                try (InputStream in = Files.newInputStream(Paths.get(path));
                        Workbook book = WorkbookFactory.create(in);
                        OutputStream out = Files.newOutputStream(Paths.get(toPath));
                        OutputStream outText = Files.newOutputStream(Paths.get(toTextPath))) {
                    logger.info("processing: " + path);
                    convert(book, out, printer -> {
                        printer.setPageSize(PDRectangle.A4, false);
                        printer.setFont(System.getenv("WINDIR") + "\\fonts\\msgothic.ttc", "MS-Gothic"); /* for windows */
                        printer.setFontSize(10.5f);
                        printer.setMargin(15);
                        printer.setLineSpace(5);
                        printer.setDrawMarginLine(drawMarginLine[0]);
                    });
                    toText(book, outText);
                    logger.info("converted: " + toPath + ", " + toTextPath);
                    count++;
                } catch (InvalidFormatException e) {
                    throw new RuntimeException("Invalid file type.");
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
                break;
            }
        }
        logger.info("processed " + count + " files.");
    }

}
