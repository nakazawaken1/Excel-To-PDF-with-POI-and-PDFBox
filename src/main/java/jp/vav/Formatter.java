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
import java.io.UncheckedIOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.logging.Logger;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import org.apache.pdfbox.pdmodel.common.PDRectangle;

/**
 * word processor like document formatter
 * 
 * @author nakazawaken1
 */
public class Formatter {

    /**
     * logger
     */
    private static final Logger logger = Logger.getLogger(Formatter.class.getCanonicalName());

    /**
     * processing top,bottom,left,right parameters
     * 
     * @param parts parts
     * @param top top action
     * @param bottom bottom action
     * @param left left action
     * @param right right action
     */
    static void tblr(String[] parts, Consumer<Float> top, Consumer<Float> bottom, Consumer<Float> left, Consumer<Float> right) {
        Tool.real(parts[2]).ifPresent(value -> {
            if (parts.length > 2) {
                parts[1].chars().forEach(c -> {
                    Optional<String> position = Optional.empty();
                    switch (Character.toUpperCase(c)) {
                    case 'T':
                        top.accept(value);
                        position = Optional.of("top");
                        break;
                    case 'B':
                        bottom.accept(value);
                        position = Optional.of("bottom");
                        break;
                    case 'L':
                        left.accept(value);
                        position = Optional.of("left");
                        break;
                    case 'R':
                        right.accept(value);
                        position = Optional.of("right");
                        break;
                    }
                    position.ifPresent(i -> logger.info(parts[0] + " " + i + ": " + value));
                });
            } else {
                top.accept(value);
                bottom.accept(value);
                left.accept(value);
                right.accept(value);
                logger.info(parts[0] + ": " + value);
            }
        });
    }

    /**
     * format marked text to PDF
     * 
     * @param source marked text
     * @param printer PDF printer
     */
    public static void format(String source, PDFPrinter printer) {
        Parser parser = new Parser(source);
        try {
            for (;;) {
                parser.skip(' ', '\t', '\r', '\n');
                if (parser.eat(":")) {
                    if (parser.eat(":")) {
                        if (parser.eat(":")) {
                            parser.skip(' ', '\t');
                            String[] pair = parser.line().split(" +", 2);
                            String[] parts = pair[0].split(":");
                            for (String part : parts) {
                                String[] p = part.split("\\-");
                                switch (p[0]) {
                                case "page":
                                    PDRectangle pageSize = Try.to(() -> (PDRectangle) PDRectangle.class.getDeclaredField(p[1].toUpperCase()).get(null)).get();
                                    boolean isLandscape = p.length > 2 && p[2].toUpperCase().charAt(0) == 'H';
                                    logger.info("pageSize: " + p[1] + ", isLandscape: " + isLandscape);
                                    printer.newPage().setPageSize(pageSize, isLandscape);
                                    break;
                                case "margin":
                                    tblr(p, i -> printer.marginTop = i, i -> printer.marginBottom = i, i -> printer.marginLeft = i,
                                            i -> printer.marginRight = i);
                                }
                            }
                            continue;
                        }
                        String[] pair = parser.line().split(" +", 2);
                        String[] parts = pair[0].split(":");
                        switch (parts[0]) {
                        case "header":
                            logger.info("header");
                            break;
                        case "footer":
                            logger.info("footer");
                            break;
                        case "table":
                            logger.info("table");
                            break;
                        }
                        do {
                            printLine(printer, parser);
                        } while (!parser.eat("::"));
                        continue;
                    }
                    printLine(printer, parser);
                    continue;
                }
                printer.println(parser.line());
            }
        } catch (EndOfText e) {
        }
    }

    /**
     * print line
     * 
     * @param printer printer
     * @param parser parser
     * @throws EndOfText end of text
     */
    private static void printLine(PDFPrinter printer, Parser parser) throws EndOfText {
        String[] pair = parser.line().split(" +", 2);
        if (pair.length <= 1) {
            printer.println(pair[0]);
            return;
        }
        String[] parts = pair[0].split(":");
        Consumer<String> write = printer::print;
        Action before = Action.noop();
        Action after = Action.noop();
        for (String part : parts) {
            String[] p = part.split("\\-");
            switch (p[0]) {
            case "left":
                break;
            case "center":
                write = printer::printCenter;
                break;
            case "right":
                write = printer::printRight;
                break;
            default:
                if (p[0].endsWith("%")) {
                    float fontSize = printer.fontSize;
                    int zoom = Tool.integer(Tool.substr(p[0], 0, -1)).orElse(100);
                    if (zoom != 100) {
                        before = before.andThen(() -> printer.setFontSize(fontSize * zoom / 100));
                        after = after.andThen(() -> printer.setFontSize(fontSize));
                    }
                }
            }
        }
        before.run();
        write.accept(pair[1]);
        printer.newLine();
        after.run();
    }

    /**
     * command
     * 
     * @param args files(.gf)
     */
    public static void main(String[] args) {
        Objects.requireNonNull(args);
        long count = Stream.of(args).map(i -> Tool.trim(i, "\"", "\"")).peek(path -> {
            try (PDFPrinter printer = new PDFPrinter()) {
                Formatter.format(new String(Files.readAllBytes(Paths.get(path)), StandardCharsets.UTF_8), printer);
                printer.saveAndClose(Files.newOutputStream(Paths.get(Tool.changeExtension(path, ".pdf"))));
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        }).count();
        if(count <= 0) {
            StringBuilder sample = new StringBuilder(":center:120% title\n:right date\ncontents...");
            IntStream.range(0, 6000).map(i -> i % 26 + 'A').forEach(i -> sample.append((char)i));
            sample.append("\n:::page-A3-h\nnext page");
            try (PDFPrinter printer = new PDFPrinter()) {
                Formatter.format(sample.toString(), printer);
                String path = "\\temp\\formatter.pdf";
                printer.saveAndClose(Files.newOutputStream(Paths.get(path)));
                logger.info("sample pdf to " + path);
                count++;
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        }
        logger.info("processed " + count + " files.");
    }

}
