package com.lancer;

import lombok.Data;

@Data
class TableFiled {
    private String field;
    private String type;
    private String length;
    private boolean isNull;
    private String key;
    private String defaultVal;
    private String extra;
    private String comment;
}
