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

package org.hoshisoft.tools.gs.formatters;

import com.google.gdata.data.spreadsheet.CustomElementCollection;
import com.google.gdata.data.spreadsheet.ListEntry;

public abstract class GSOutputFormatter {
    private String separator;
    
    public GSOutputFormatter(String separator) {
        this.separator = separator;
    }
    
    public String format(ListEntry entry, boolean isHeader) {
        StringBuilder sb = new StringBuilder();

        CustomElementCollection elements = entry.getCustomElements();

        boolean isFirst = true;
        for (String tag : elements.getTags()) {
            if (!isFirst) {
                sb.append(separator);
            } else {
                isFirst = false;
            }

            if (isHeader) {
                sb.append(tag);
            } else {
                sb.append(elements.getValue(tag));
            }
        }

        return sb.toString();
    }
}
