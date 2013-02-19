package org.hoshisoft.tools.gs.enums;


public enum GoogleDocumentsVisibility {
    PUBLIC("public"), PRIVATE("private");
    
    String value;
    
    private GoogleDocumentsVisibility(String value) {
        this.value = value;
    }
    
    public String value() {
        return value;
    }
    
    public static GoogleDocumentsVisibility fromValue(final String v) {
        for (GoogleDocumentsVisibility t : GoogleDocumentsVisibility.values()) {
            if (t.value.equals(v)) {
                return t;
            }
        }
        
        throw new IllegalArgumentException(v);
    }
}
