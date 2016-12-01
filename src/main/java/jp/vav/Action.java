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

import java.util.Objects;

/**
 * functional interface that add andThen and noop to Runnable
 * 
 * @author nakazawaken1
 */
public interface Action {

    /**
     * run
     */
    void run();

    /**
     * combine action
     * 
     * @param after after action
     * @return action
     */
    default Action andThen(Action after) {
        Objects.requireNonNull(after);
        return () -> {
            run();
            after.run();
        };
    }

    /**
     * @return no operation
     */
    static Action noop() {
        return () -> {
        };
    }
}
