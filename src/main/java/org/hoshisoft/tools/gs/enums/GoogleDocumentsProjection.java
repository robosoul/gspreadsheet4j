package org.hoshisoft.tools.gs.enums;

public enum GoogleDocumentsProjection {
    FULL("full"), BASIC("basic");
    
    private String value;
    
    private GoogleDocumentsProjection(String value) {
        this.value = value;        
    }
    
    public String value() {
        return value;
    }
    
    public GoogleDocumentsProjection fromValue(String v) {
        for (GoogleDocumentsProjection p : GoogleDocumentsProjection.values()) {
            if (p.value.equals(v)) {
                return p;
            }
        }
        
        throw new IllegalArgumentException(v);
    }
}
