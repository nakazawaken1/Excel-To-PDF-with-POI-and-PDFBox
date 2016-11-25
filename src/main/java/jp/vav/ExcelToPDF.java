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
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.chrono.JapaneseChronology;
import java.time.format.DateTimeFormatter;
import java.time.temporal.Temporal;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
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
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
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
    @SuppressWarnings("deprecation")
    public static Optional<Object> cellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
        case NUMERIC:
            int index = cell.getCellStyle().getDataFormat();
            Function<Temporal, String> formatter = formats.get(index);
            return Optional.ofNullable(formatter == null ? new DataFormatter(Locale.JAPAN).formatCellValue(cell)
                    : formatter.apply(LocalDateTime.ofInstant(cell.getDateCellValue().toInstant(), ZoneOffset.systemDefault())));
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
     * @param date Japanese date string
     * @return converted date string
     */
    static String eraKanjiToAlpha(String date) {
        return trim(date.replace("明治", "M").replace("大正", "T").replace("昭和", "S").replace("平成", "H").replaceAll("[年月日]", "."), null, ".");
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
     * date formatters for Japanese
     */
    @SuppressWarnings("serial")
    static final Map<Integer, Function<Temporal, String>> formats = new HashMap<Integer, Function<Temporal, String>>() {
        {
            DateTimeFormatter yyyymd = DateTimeFormatter.ofPattern("yyyy/M/d");
            DateTimeFormatter jyyyymd = DateTimeFormatter.ofPattern("yyyy年M月d");
            DateTimeFormatter yyyymdhmm = DateTimeFormatter.ofPattern("yyyy/M/d H:mm");
            DateTimeFormatter mdyy = DateTimeFormatter.ofPattern("M/d/yy");
            DateTimeFormatter jhmm = DateTimeFormatter.ofPattern("H'時'mm'分'");
            DateTimeFormatter jhmmss = DateTimeFormatter.ofPattern("H'時'mm'分'ss'秒'");
            DateTimeFormatter jym = DateTimeFormatter.ofPattern("yyyy'年'M'月'");
            DateTimeFormatter jmd = DateTimeFormatter.ofPattern("M'月'd'日'");
            DateTimeFormatter jymd = DateTimeFormatter.ofPattern("Gy年M月d日").withChronology(JapaneseChronology.INSTANCE);
            put(14, yyyymd::format);
            put(22, yyyymdhmm::format);
            put(27, ((Function<Temporal, String>) jymd::format).andThen(ExcelToPDF::eraKanjiToAlpha));
            put(28, jymd::format);
            put(29, jymd::format);
            put(30, mdyy::format);
            put(31, jyyyymd::format);
            put(32, jhmm::format);
            put(33, jhmmss::format);
            put(34, yyyymd::format);
            put(35, jmd::format);
            put(36, ((Function<Temporal, String>) jymd::format).andThen(ExcelToPDF::eraKanjiToAlpha));
            put(55, jym::format);
            put(56, jmd::format);
            put(57, ((Function<Temporal, String>) jymd::format).andThen(ExcelToPDF::eraKanjiToAlpha));
            put(58, jymd::format);
        }
    };

    /**
     * PDF page wrapper
     */
    /**
     * @author nk0126
     *
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
         * PDF page
         */
        protected Optional<PDPage> page = Optional.empty();

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
         * setup at page created
         */
        protected Optional<Consumer<Printer>> pageSetup = Optional.empty();

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
                float width = font.getStringWidth(text) * fontSize / 1000; /* font.getStringWidth must be after getContent */
                float max = getPage().getMediaBox().getWidth() - marginRight - currentX;
                if (width > max) { /* wrap */
                    int length = text.length();
                    int count = length / 2;
                    while (count > 1 && font.getStringWidth(text.substring(0, count)) * fontSize / 1000 > max) {
                        count /= 2;
                    }
                    while (count < length && font.getStringWidth(text.substring(0, count + 1)) * fontSize / 1000 <= max) {
                        count++;
                    }
                    content.showText(text.substring(0, count));
                    newLine();
                    print(text.substring(count));
                    return;
                }
                currentX += width;
                content.showText(text);
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
            if (currentY + fontSize > getPage().getMediaBox().getHeight() - marginButtom) {
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
            content.ifPresent(throwable(PDPageContentStream::endText).andThen(throwable(PDPageContentStream::close)));
            content.ifPresent(i -> logger.config("destroy PDPageContentStream: " + i.hashCode()));
            page = Optional.empty();
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
         * @return PDF page
         */
        protected PDPage getPage() {
            return page.orElseGet(() -> {
                PDPage page = new PDPage(pageSize);
                logger.config("add PDPage: " + page.hashCode());
                this.page = Optional.of(page);
                pageSetup.ifPresent(i -> i.accept(this));
                getDocument().addPage(page);
                return page;
            });
        }

        /**
         * @return PDF content
         */
        protected PDPageContentStream getContent() {
            return content.orElseGet(throwable(() -> {
                PDPageContentStream content = new PDPageContentStream(getDocument(), getPage());
                logger.config("create PDPageContentStream: " + content.hashCode());
                this.content = Optional.of(content);
                content.beginText();
                content.setFont(font, fontSize);
                content.newLineAtOffset(marginLeft, getPage().getMediaBox().getHeight() - fontSize - marginTop);
                currentY = marginTop;
                contentSetup.ifPresent(i -> i.accept(this));
                return content;
            }));
        }

        /**
         * @param margin top, bottom, left and right margin
         */
        public void setMargin(float margin) {
            marginTop = marginButtom = marginLeft = marginRight = margin;
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
                        cellValue(cell).ifPresent(
                                value -> printer.println(CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1) + ": " + value));
                    });
                });
                sheet.getCellComments().entrySet().forEach(entry -> {
                    printer.println("[comment] " + entry.getKey() + ": " + entry.getValue());
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
                        cellValue(cell).ifPresent(
                                value -> printer.println(CellReference.convertNumToColString(cell.getColumnIndex()) + (cell.getRowIndex() + 1) + ": " + value));
                    });
                });
                sheet.getCellComments().entrySet().forEach(entry -> {
                    printer.println("[comment] " + entry.getKey() + ": " + entry.getValue());
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
     * @param args Excel files(.xls, .xlsx, .xlsm)
     */
    public static void main(String[] args) {
        Objects.requireNonNull(args);
        logger.info("processed " + Stream.of(args).map(path -> trim(path, "\"", "\"")).peek(path -> {
            String toPath = changeExtension(path, ".pdf");
            String toTextPath = changeExtension(path, ".txt");
            try (InputStream in = Files.newInputStream(Paths.get(path));
                    Workbook book = throwable(() -> WorkbookFactory.create(in)).get();
                    OutputStream out = Files.newOutputStream(Paths.get(toPath));
                    OutputStream outText = Files.newOutputStream(Paths.get(toTextPath))) {
                logger.info("processing: " + path);
                convert(book, out, printer -> {
                    printer.setPageSize(PDRectangle.A4, false);
                    printer.setFont(System.getenv("WINDIR") + "\\fonts\\msgothic.ttc", "MS-Gothic"); /* for windows */
                    printer.setFontSize(10.5f);
                    printer.setMargin(15);
                    printer.setLineSpace(5);
                });
                toText(book, outText);
                logger.info("converted: " + toPath + ", " + toTextPath);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        }).count() + " files.");
    }

}
