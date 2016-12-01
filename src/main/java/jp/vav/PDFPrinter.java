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
import java.io.OutputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.logging.Logger;

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

/**
 * PDF printer supported auto line and page break, printer center and right, change font and font size, change page size, change margin, change line spacing
 * 
 * @author nakazawaken1
 */
public class PDFPrinter implements AutoCloseable {

    /**
     * logger
     */
    private static final Logger logger = Logger.getLogger(PDFPrinter.class.getCanonicalName());

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
    protected float marginBottom = 10.5f;

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
    protected Optional<PDPageContentStream> page = Optional.empty();

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
    protected Optional<Consumer<PDFPrinter>> documentSetup = Optional.empty();

    /**
     * setup at content created
     */
    protected Optional<Consumer<PDFPrinter>> pageSetup = Optional.empty();

    /**
     * draw margin line if true
     */
    protected boolean drawMarginLine = false;

    /**
     * debug points
     */
    protected List<Float> debugPoints = new ArrayList<>();

    /**
     * draw debug point if true
     */
    private boolean drawDebugPoint = false;

    /**
     * current horizontal position
     */
    protected float currentX;

    /**
     * current vertical position
     */
    protected float currentY;

    /**
     * previous position
     */
    protected float currentX0;

    /**
     * add current x
     * 
     * @param offset offset
     */
    protected void addCurrentX(float offset) {
        currentX0 = currentX;
        currentX += offset;
        debugPoints.add(currentX);
    }

    /**
     * find index for fit in width
     * 
     * @param text text
     * @param width max width
     * @return index
     */
    protected int fitIndex(String text, float width) {
        if (width <= 0) {
            return 0;
        }
        int length = text.length();
        int count = length / 2;
        while (count > 1 && textWidth(text.substring(0, count)) > width) {
            count /= 2;
        }
        while (count < length && textWidth(text.substring(0, count + 1)) <= width) {
            count++;
        }
        return count;
    }

    /**
     * @return width of inner margin
     */
    protected float innerWidth() {
        return pageSize.getWidth() - marginLeft - marginRight;
    }

    /**
     * @return height of inner margin
     */
    protected float innerHeight() {
        return pageSize.getHeight() - marginTop - marginBottom;
    }

    /**
     * @return remains width
     */
    protected float remainingWidth() {
        getPage();
        return pageSize.getWidth() - marginRight - currentX;
    }

    /**
     * @return remains height
     */
    protected float remainingHeight() {
        getPage();
        return pageSize.getHeight() - marginBottom - currentY;
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
    protected PDPageContentStream getPage() {
        return page.orElseGet(Try.to(() -> {
            PDPage p = new PDPage(pageSize);
            logger.config("add PDPage: " + p.hashCode());
            getDocument().addPage(p);
            PDPageContentStream page = new PDPageContentStream(getDocument(), p);
            logger.config("create PDPageContentStream: " + page.hashCode());
            this.page = Optional.of(page);
            page.beginText();
            page.setFont(font, fontSize);
            page.newLineAtOffset(marginLeft, pageSize.getHeight() - fontSize - marginTop);
            currentX0 = currentX = marginLeft;
            currentY = marginTop;
            pageSetup.ifPresent(i -> i.accept(this));
            return page;
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
        return text == null || text.isEmpty() ? 0 : Try.to(() -> font.getStringWidth(text) * fontSize / 1000f).get();
    }

    /**
     * create printer
     */
    public PDFPrinter() {
        logger.config("create Printer: " + hashCode());
        /* for Japanese */
        String path = System.getenv("WINDIR") + "\\fonts\\msgothic.ttc";
        if(Files.exists(Paths.get(path))) {
            setFont(path, "MS-Gothic");
        }
    }

    /**
     * print text
     * 
     * @param text text
     * @return this
     */
    public PDFPrinter print(String text) {
        if (text == null || text.isEmpty()) {
            return this;
        }
        try {
            final String lineBreak = "\n";
            int index = text.indexOf(lineBreak);
            if (index >= 0) {
                println(text.substring(0, index));
                print(text.substring(index + lineBreak.length()));
                return this;
            }
            PDPageContentStream content = getPage();
            float width = textWidth(text); /* font.getStringWidth must be after getContent */
            float max = remainingWidth();
            if (width > max) { /* wrap */
                int fit = fitIndex(text, max);
                print(text.substring(0, fit));
                newLine();
                print(text.substring(fit));
                return this;
            }
            content.showText(text);
            addCurrentX(width);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
        return this;
    }

    /**
     * print text and newline
     * 
     * @param text text
     * @return this
     */
    public PDFPrinter println(String text) {
        return print(text).newLine();
    }

    /**
     * print text with centering
     * 
     * @param text text
     * @return this
     */
    public PDFPrinter printCenter(String text) {
        float max = remainingWidth() * 2 - innerWidth();
        float width = textWidth(text);
        if (width > max) {
            int fit = fitIndex(text, max);
            printCenter(text.substring(0, fit));
            newLine();
            printCenter(text.substring(fit));
            return this;
        }
        Try.<PDPageContentStream> to(c -> {
            float left = (max - width) / 2;
            c.newLineAtOffset(left + currentX - currentX0, 0);
            addCurrentX(left);
            c.showText(text);
            addCurrentX(width);
        }).accept(getPage());
        return this;
    }

    /**
     * print text with centering
     * 
     * @param text text
     * @return this
     */
    public PDFPrinter printRight(String text) {
        float max = remainingWidth();
        float width = textWidth(text);
        if (width > max) {
            int fit = fitIndex(text, max);
            printRight(text.substring(0, fit));
            newLine();
            printRight(text.substring(fit));
            return this;
        }
        Try.<PDPageContentStream> to(c -> {
            float left = max - width;
            c.newLineAtOffset(left + currentX - currentX0, 0);
            addCurrentX(left);
            c.showText(text);
            addCurrentX(width);
        }).accept(getPage());
        return this;
    }

    /**
     * print newline
     * 
     * @return this
     */
    public PDFPrinter newLine() {
        currentY += fontSize + lineSpace;
        if (fontSize > remainingHeight()) {
            newPage();
        } else {
            Try.to(() -> getPage().newLineAtOffset(marginLeft - currentX0, -fontSize - lineSpace)).run();
            currentX0 = currentX = marginLeft;
        }
        return this;
    }

    /**
     * break page
     * 
     * @return this
     */
    public PDFPrinter newPage() {
        page.ifPresent(i -> {
            try {
                i.endText();
                if (drawMarginLine) {
                    i.addRect(marginLeft, marginBottom, innerWidth(), innerHeight());
                    i.setLineWidth(0.01f);
                    i.setLineDashPattern(new float[] { 3f, 1f }, 0);
                    i.stroke();
                }
                if(drawDebugPoint) {
                    debugPoints.forEach(Try.to(p -> {
                        i.moveTo(p, pageSize.getHeight() - marginTop / 2);
                        i.lineTo(p, pageSize.getHeight());
                        i.stroke();
                    }));
                }
                i.close();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        });
        page.ifPresent(i -> logger.config("destroy PDPageContentStream: " + i.hashCode()));
        page = Optional.empty();
        return this;
    }

    /**
     * save and close
     * 
     * @param out output to PDF
     * @return this
     * @throws IOException I/O exception
     */
    public PDFPrinter saveAndClose(OutputStream out) throws IOException {
        try {
            newPage();
            getDocument().save(out);
        } finally {
            out.close();
        }
        return this;
    }

    /*
     * (non-Javadoc)
     * 
     * @see java.lang.AutoCloseable#close()
     */
    @Override
    public void close() {
        newPage();
        document.ifPresent(Try.to(i -> {
            i.close();
            ttf.ifPresent(Try.to(TrueTypeFont::close));
            ttc.ifPresent(Try.to(TrueTypeCollection::close));
        }));
        document.ifPresent(i -> logger.config("destroy PDDocument: " + i.hashCode()));
        logger.config("destroy Printer: " + hashCode());
    }

    /**
     * @param point top, bottom, left and right margin
     * @return this
     */
    public PDFPrinter setMargin(float point) {
        marginTop = marginBottom = marginLeft = marginRight = point;
        return this;
    }

    /**
     * @param vertical top and bottom margin
     * @param horizontal left and right margin
     * @return this
     */
    public PDFPrinter setMargin(float vertical, float horizontal) {
        marginTop = marginBottom = vertical;
        marginLeft = marginRight = horizontal;
        return this;
    }

    /**
     * @param top top margin
     * @param bottom bottom margin
     * @param left left margin
     * @param right right margin
     * @return this
     */
    public PDFPrinter setMargin(float top, float bottom, float left, float right) {
        marginTop = top;
        marginBottom = bottom;
        marginLeft = left;
        marginRight = right;
        return this;
    }

    /**
     * @param page page size
     * @param isLandscape true: landscape, false: portrait
     * @return this
     */
    public PDFPrinter setPageSize(PDRectangle page, boolean isLandscape) {
        Objects.requireNonNull(page);
        pageSize = isLandscape ? new PDRectangle(page.getHeight(), page.getWidth()) : page;
        return this;
    }

    /**
     * @param font font
     * @return this
     */
    public PDFPrinter setFont(PDFont font) {
        Objects.requireNonNull(font);
        this.font = font;
        return this;
    }

    /**
     * @param path .ttf file path
     * @return this
     */
    public PDFPrinter setFont(String path) {
        Objects.requireNonNull(path);
        setFont(Try.to(() -> {
            ttf = Optional.of(new TTFParser().parse(new File(path)));
            return PDType0Font.load(getDocument(), ttf.get(), true);
        }).get());
        return this;
    }

    /**
     * @param path .ttc file path
     * @param name font name
     * @return this
     */
    public PDFPrinter setFont(String path, String name) {
        Objects.requireNonNull(path);
        Objects.requireNonNull(name);
        setFont(Try.to(() -> {
            ttc = Optional.of(new TrueTypeCollection(new File(path)));
            ttf = Optional.of(ttc.get().getFontByName(name));
            return PDType0Font.load(getDocument(), ttf.get(), true);
        }).get());
        return this;
    }

    /**
     * @param point font size
     * @return this
     */
    public PDFPrinter setFontSize(float point) {
        fontSize = point;
        page.ifPresent(Try.<PDPageContentStream>to(i -> i.setFont(font, fontSize)));
        return this;
    }

    /**
     * @param point space between lines
     * @return this
     */
    public PDFPrinter setLineSpace(float point) {
        lineSpace = point;
        return this;
    }

    /**
     * @param enabled draw margin line if true
     * @return this
     */
    public PDFPrinter setDrawMarginLine(boolean enabled) {
        drawMarginLine = enabled;
        return this;
    }

    /**
     * @param enabled draw debug point if true
     * @return this
     */
    public PDFPrinter setDrawDebugPoint(boolean enabled) {
        drawDebugPoint = enabled;
        return this;
    }

    /**
     * example
     * 
     * @param args not use
     */
    public static void main(String[] args) {
        Tool.logSetup(null);
        try (PDFPrinter printer = new PDFPrinter()) {
            printer.setDrawMarginLine(true).setDrawDebugPoint(true);
            printer.printRight("right").newLine();
            printer.setFontSize(18f).printCenter("center large text").newLine().setFontSize(10.5f);
            printer.print("left").newLine();
            printer.newPage().setPageSize(PDRectangle.A5, true);
            printer.print("second page");
            String path = "\\temp\\PDFPrinter.pdf";
            printer.saveAndClose(Files.newOutputStream(Paths.get(path)));
            logger.info("sample to " + path);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
}
