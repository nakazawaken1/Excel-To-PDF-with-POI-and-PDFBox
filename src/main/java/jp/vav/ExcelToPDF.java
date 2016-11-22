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
import java.util.Iterator;
import java.util.Objects;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.EncryptedDocumentException;
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

/**
 * Entry point class
 * 
 * @author nakazawaken1
 */
public class ExcelToPDF {

    /**
     * テスト出力
     */
    private static final PrintStream debug = System.out;

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
        public PDRectangle pageSize = PDRectangle.A4;

        /**
         * Font size
         */
        public float fontSize = 10.5f;

        /**
         * マージン
         */
        protected float margin = 10.5f;

        /**
         * font
         */
        public PDFont font = PDType1Font.HELVETICA;

        /**
         * PDF document
         */
        public Optional<PDDocument> document = Optional.empty();

        /**
         * PDF page
         */
        protected Optional<PDPage> page = Optional.empty();

        /**
         * PDF contents
         */
        protected Optional<PDPageContentStream> content = Optional.empty();

        /**
         * setup at document created
         */
        protected Optional<BiConsumer<Printer, PDDocument>> documentSetup;

        /**
         * setup at page created
         */
        protected Optional<BiConsumer<Printer, PDPage>> pageSetup;

        /**
         * setup at content created
         */
        protected Optional<BiConsumer<Printer, PDPageContentStream>> contentSetup;

        /**
         * current vertical position
         */
        protected float currentY;

        /**
         * Add page to document
         * 
         * @param out output
         * @param documentSetup setup at document created
         * @param pageSetup setup at page created
         * @param contentSetup setup at content created
         * @return Page
         */
        public static Printer use(OutputStream out, BiConsumer<Printer, PDDocument> documentSetup, BiConsumer<Printer, PDPage> pageSetup,
                BiConsumer<Printer, PDPageContentStream> contentSetup) {
            Objects.requireNonNull(out);
            Printer printer = new Printer();
            debug.println("create Printer: " + printer.hashCode());
            printer.out = out;
            printer.documentSetup = Optional.ofNullable(documentSetup);
            printer.pageSetup = Optional.ofNullable(pageSetup);
            printer.contentSetup = Optional.ofNullable(contentSetup);
            return printer;
        }

        /**
         * print text
         * 
         * @param text text
         */
        public void print(String text) {
            throwable(() -> getContent().showText(text)).run();
        }

        /**
         * print newline
         */
        public void newline() {
            currentY += fontSize;
            if (currentY > getPage().getMediaBox().getHeight() - margin * 2) {
                newpage();
            } else {
                throwable(() -> getContent().newLineAtOffset(0, -fontSize)).run();
            }
        }

        /**
         * print text and newline
         * 
         * @param text text
         */
        public void println(String text) {
            print(text);
            newline();
        }

        /**
         * break page
         */
        public void newpage() {
            content.ifPresent(throwable(PDPageContentStream::endText).andThen(throwable(PDPageContentStream::close)));
            content.ifPresent(i -> debug.println("destroy PDPageContentStream: " + i.hashCode()));
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
            newpage();
            document.ifPresent(throwable(i -> {
                i.save(out);
                i.close();
            }));
            document.ifPresent(i -> debug.println("destroy PDDocument: " + i.hashCode()));
            debug.println("destroy Printer: " + hashCode());
        }

        /**
         * @return PDF document
         */
        public PDDocument getDocument() {
            return document.orElseGet(() -> {
                PDDocument document = new PDDocument();
                debug.println("add PDDocument: " + document.hashCode());
                this.document = Optional.of(document);
                documentSetup.ifPresent(i -> i.accept(this, document));
                return document;
            });
        }

        /**
         * @return PDF page
         */
        public PDPage getPage() {
            return page.orElseGet(() -> {
                PDPage page = new PDPage(pageSize);
                debug.println("add PDPage: " + page.hashCode());
                this.page = Optional.of(page);
                pageSetup.ifPresent(c -> c.accept(this, page));
                getDocument().addPage(page);
                return page;
            });
        }

        /**
         * @return PDF content
         */
        private PDPageContentStream getContent() {
            return content.orElseGet(throwable(() -> {
                PDPageContentStream content = new PDPageContentStream(getDocument(), getPage());
                debug.println("create PDPageContentStream: " + content.hashCode());
                this.content = Optional.of(content);
                content.beginText();
                content.setFont(font, fontSize);
                content.newLineAtOffset(margin, getPage().getMediaBox().getHeight() - fontSize - margin);
                currentY = margin;
                contentSetup.ifPresent(c -> c.accept(this, content));
                return content;
            }));
        }
    }

    /**
     * excel to pdf
     * 
     * @param book excel workbook
     * @param out output to pdf
     * @param documentSetup document setup
     * @throws IOException failed to PDF creation
     */
    public static void convert(Workbook book, OutputStream out, BiConsumer<Printer, PDDocument> documentSetup) throws IOException {
        try (Printer printer = Printer.use(out, documentSetup, null, null)) {
            for (int i = 0, end = book.getNumberOfSheets(); i < end; i++) {
                Sheet sheet = book.getSheetAt(i);
                int rowCount = sheet.getPhysicalNumberOfRows();
                if (rowCount <= 0) {
                    debug.println(sheet.getSheetName() + ": empty");
                    continue; /* skip blank sheet */
                }
                debug.println(sheet.getSheetName() + ": " + rowCount + " rows");
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
                printer.newpage();
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
                try (TrueTypeCollection fonts = new TrueTypeCollection(new File(System.getenv("WINDIR") + "\\fonts\\msgothic.ttc")); /* for windows */
                        TrueTypeFont font = fonts.getFontByName("MS-Gothic")) {
                    convert(book, out, (printer, document) -> {
                        printer.pageSize = new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth());
                        printer.font = throwable(() -> PDType0Font.load(document, font, true)).get();
                        printer.margin = 10;
                    });
                }
                debug.println("converted: " + toPath);
                debug.println("--------");
            } catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
                e.printStackTrace();
            }
        }).count() + " files.");
    }

}
