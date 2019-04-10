package de.intranda.goobi.plugins.util;


import lombok.Data;

@Data
public class MetadataMappingObject {

    private String rulesetName;
    private String propertyName;
    private Integer excelColumn;
    private Integer identifierColumn;

    private String headerName;

    private String normdataHeaderName;

    private String docType ;
}
