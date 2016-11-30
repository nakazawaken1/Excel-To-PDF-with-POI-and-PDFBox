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

import java.util.Iterator;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.function.Predicate;
import java.util.logging.ConsoleHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.ss.usermodel.Cell;

import com.github.mygreen.cellformatter.POICellFormatter;

/**
 * static methods
 * 
 * @author nakazawaken1
 */
public class Tool {

    /**
     * cell value formatter
     */
    public static final POICellFormatter formatter = new POICellFormatter();

    static {
        logSetup(System.getProperty("java.util.logging.level"));
    }

    /**
     * setup logging
     * 
     * @param levelText log level text
     */
    public static void logSetup(String levelText) {
        Logger.getGlobal().info("log level : " + levelText);
        if (levelText != null) {
            Level level = Level.parse(levelText);
            for (Handler handler : Logger.getLogger("").getHandlers()) {
                if (handler instanceof ConsoleHandler) {
                    handler.setLevel(level);
                }
            }
        }
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
        return Optional.ofNullable(formatter.formatAsString(cell, Locale.JAPAN)).filter(((Predicate<String>) String::isEmpty).negate());
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
}
