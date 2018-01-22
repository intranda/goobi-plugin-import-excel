package de.intranda.goobi.plugins.util;

import lombok.Data;

@Data
public class PersonMappingObject {

    private String rulesetName;
    private Integer firstnameColumn;
    private Integer lastnameColumn;
    private Integer identifierColumn;
    
    
}
