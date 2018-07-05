package de.intranda.goobi.plugins.util;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;

import lombok.Data;

@Data
public class Config {

    private String publicationType;
    private String collection;
    private int firstLine;
    private int identifierColumn;
    private int conditionalColumn;
    private List<MetadataMappingObject> metadataList = new ArrayList<>();
    private List<PersonMappingObject> personList = new ArrayList<>();
    private List<GroupMappingObject> groupList = new ArrayList<>();

    /**
     * loads the &lt;config&gt; block from xml file
     * 
     * @param xmlConfig
     */
    
    public Config(SubnodeConfiguration xmlConfig) {

        publicationType = xmlConfig.getString("/publicationType", "Monograph");
        collection = xmlConfig.getString("/collection");
        firstLine = xmlConfig.getInt("/firstLine", 1);
        identifierColumn = xmlConfig.getInt("/identifierColumn", 1);
        conditionalColumn = xmlConfig.getInt("/conditionalColumn", identifierColumn);

        List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
        for (HierarchicalConfiguration md : mml) {
            metadataList.add(getMetadata(md));
        }

        List<HierarchicalConfiguration> pml = xmlConfig.configurationsAt("//person");
        for (HierarchicalConfiguration md : pml) {
            personList.add(getPersons(md));
        }

        List<HierarchicalConfiguration> gml = xmlConfig.configurationsAt("//group");
        for (HierarchicalConfiguration md : gml) {
            String rulesetName = md.getString("@ugh");
            GroupMappingObject grp = new GroupMappingObject();
            grp.setRulesetName(rulesetName);
            List<HierarchicalConfiguration> subList = md.configurationsAt("//person");
            for (HierarchicalConfiguration sub : subList) {
                PersonMappingObject pmo = getPersons(sub);
                grp.getPersonList().add(pmo);
            }

            subList = md.configurationsAt("//metadata");
            for (HierarchicalConfiguration sub : subList) {
                MetadataMappingObject pmo = getMetadata(sub);
                grp.getMetadataList().add(pmo);
            }

            groupList.add(grp);

        }
    }

    private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        String propertyName = md.getString("@name");
        Integer columnNumber = md.getInteger("@column", null);
        Integer identifierColumn = md.getInteger("@identifier", null);
        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setExcelColumn(columnNumber);
        mmo.setIdentifierColumn(identifierColumn);
        mmo.setPropertyName(propertyName);
        mmo.setRulesetName(rulesetName);
        return mmo;
    }

    private PersonMappingObject getPersons(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        Integer firstname = md.getInteger("firstname", null);
        Integer lastname = md.getInteger("lastname", null);
        Integer identifier = md.getInteger("identifier", null);
        PersonMappingObject pmo = new PersonMappingObject();
        pmo.setFirstnameColumn(firstname);
        pmo.setLastnameColumn(lastname);
        pmo.setIdentifierColumn(identifier);
        pmo.setRulesetName(rulesetName);
        return pmo;

    }

}
