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

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Supplier;

/**
 * lambda exception processor
 * 
 * @author nakazawaken1
 */
public class Try {

    /**
     * throwsable action
     * 
     * @author nakazawaken1
     */
    public interface TryAction {
        /**
         * @throws Exception Exception
         */
        void run() throws Exception;
    }

    /**
     * throwsable consumer
     * 
     * @author nakazawaken1
     * @param <T> consume type
     */
    public interface TryConsumer<T> {
        /**
         * @param value consume value
         * @throws Exception Exception
         */
        void accept(T value) throws Exception;
    }

    /**
     * throwsable supplier
     * 
     * @author nakazawaken1
     * @param <T> supply type
     */
    public interface TrySupplier<T> {
        /**
         * @return supply value
         * @throws Exception Exception
         */
        T get() throws Exception;
    }

    /**
     * throwsable function
     * 
     * @author nakazawaken1
     * @param <T> argument type
     * @param <R> result type
     */
    public interface TryFunction<T, R> {
        /**
         * @param argument argument
         * @return supply value
         * @throws Exception Exception
         */
        R apply(T argument) throws Exception;
    }

    /**
     * throwable action
     * 
     * @author nakazawaken1
     * @param action throwable action
     * @return action
     */
    public static Action to(TryAction action) {
        return () -> {
            try {
                action.run();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (RuntimeException e) {
                throw e;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable supplier
     * @param <T> Item type
     * @param consumer throwable consumer
     * @return consumer
     */
    public static <T> Consumer<T> to(TryConsumer<T> consumer) {
        return value -> {
            try {
                consumer.accept(value);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (RuntimeException e) {
                throw e;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable supplier
     * @param <T> Argument type
     * 
     * @param supplier throwable supplier
     * @return supplier
     */
    public static <T> Supplier<T> to(TrySupplier<T> supplier) {
        return () -> {
            try {
                return supplier.get();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (RuntimeException e) {
                throw e;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }

    /**
     * throwable function
     * @param <T> Argument type
     * @param <R> Result type
     * 
     * @param function throwable function
     * @return function
     */
    public static <T, R> Function<T, R> to(TryFunction<T, R> function) {
        return argument -> {
            try {
                return function.apply(argument);
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            } catch (RuntimeException e) {
                throw e;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        };
    }
}
