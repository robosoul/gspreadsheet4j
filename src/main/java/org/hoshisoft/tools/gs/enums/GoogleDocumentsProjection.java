/**
 *   Copyright 2013 Luka Obradovic
 *
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */

package org.hoshisoft.tools.gs.enums;

public enum GoogleDocumentsProjection {
    FULL("full"), BASIC("basic");
    
    private final String value;
    
    private GoogleDocumentsProjection(final String value) {
        this.value = value;        
    }
    
    public String value() {
        return value;
    }
    
    public GoogleDocumentsProjection fromValue(final String v) {
        for (GoogleDocumentsProjection p : GoogleDocumentsProjection.values()) {
            if (p.value.equals(v)) {
                return p;
            }
        }
        
        throw new IllegalArgumentException(v);
    }
}
