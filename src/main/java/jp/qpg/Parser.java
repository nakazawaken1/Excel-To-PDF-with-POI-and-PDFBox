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
package jp.qpg;

import java.util.stream.IntStream;

/**
 * general text parser
 * 
 * @author nakazawaken1
 */
public class Parser {

    /**
     * source text
     */
    StringBuilder source;

    /**
     * current index
     */
    int index;

    /**
     * last index
     */
    int lastIndex;

    /**
     * constructor
     *
     * @param text source text
     */
    Parser(String text) {
        source = new StringBuilder(text);
        index = 0;
        lastIndex = source.length();
    }

    /**
     * skip letters
     * 
     * @param letters letters
     * @throws EndOfText end of text
     */
    void skip(int... letters) throws EndOfText {
        for (;; index++) {
            if (index >= lastIndex) {
                throw EndOfText.instance;
            }
            int letter = source.charAt(index);
            for (int i : letters) {
                if (i != letter) {
                    return;
                }
            }
        }
    }

    /**
     * skip until any letters
     *
     * @param letters letters
     * @throws EndOfText end of text
     */
    void skipUntil(int... letters) throws EndOfText {
        for (;; index++) {
            if (index >= lastIndex) {
                throw EndOfText.instance;
            }
            int letter = source.charAt(index);
            for (int i : letters) {
                if (i == letter) {
                    return;
                }
            }
        }
    }

    /**
     * eat word
     *
     * @param word word
     * @return true if ate word
     * @throws EndOfText end of text
     */
    boolean eat(String word) throws EndOfText {
        int newIndex = index + word.length();
        if (newIndex >= lastIndex) {
            throw EndOfText.instance;
        }
        if (source.substring(index, newIndex).equals(word)) {
            index = newIndex;
            return true;
        }
        return false;
    }

    /**
     * read text from index to endIndex
     * 
     * @param endIndex end index
     * @return text
     * @throws EndOfText end of text
     */
    String substring(int endIndex) throws EndOfText {
        if (endIndex > lastIndex) {
            throw EndOfText.instance;
        }
        String result = source.substring(index, endIndex);
        index = endIndex;
        return result;
    }

    /**
     * compare previous word
     *
     * @param word word
     * @return true if match word
     */
    boolean prev(String word) {
        int length = word.length();
        return index - length >= 0 && source.substring(index - length, index).equals(word);
    }

    /**
     * get index of word
     *
     * @param word word
     * @return -1 if not found word, index if found
     */
    int indexOf(String word) {
        return source.indexOf(word, index);
    }

    /**
     * get index of nearest letter
     *
     * @param letters letters
     * @return -1 if not found word, index if found
     */
    int indexOf(int... letters) {
        return IntStream.of(letters).map(i -> source.indexOf(String.valueOf((char) i), index)).filter(i -> i >= 0).min().orElse(-1);
    }

    /**
     * get line text
     * 
     * @return line text
     * @throws EndOfText end of text
     */
    String line() throws EndOfText {
        int lineEnd = indexOf('\r', '\n');
        if (lineEnd < 0) {
            if(index < lastIndex) {
                return substring(lastIndex);
            }
            throw EndOfText.instance;
        }
        String result = substring(lineEnd);
        index++;
        eat("\n");
        return result;
    }
}
