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
