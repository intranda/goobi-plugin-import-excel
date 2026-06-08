package de.intranda.goobi.plugins.util;

import java.util.ArrayList;
import java.util.List;

import lombok.Data;

@Data
public class MetadataMappingObject {

    private String rulesetName;
    private String propertyName;
    private Integer excelColumn;
    private String headerName;
    private String normdataHeaderName;
    private String docType;
    private String searchField;
    private boolean splittingAllowed;

    private boolean required;
    private String pattern;
    private List<String> validContent = new ArrayList<>();
    private String listErrorMessage;
    private String patternErrorMessage;
    private String requiredErrorMessage;
}
