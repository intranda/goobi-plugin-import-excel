package de.intranda.goobi.plugins.util;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.tree.ConfigurationNode;

import lombok.Data;

@Data
public class ExcelConfig {

    private String publicationType;
    private String collection;
    private int firstLine;
    private int identifierColumn;
    private int conditionalColumn;
    private int rowHeader;
    private int rowDataStart;
    private int rowDataEnd;
    private List<MetadataMappingObject> metadataList = new ArrayList<>();
    private List<PersonMappingObject> personList = new ArrayList<>();
    private List<PersonMappingObject> personWithRoleList = new ArrayList<>();    
    private List<GroupMappingObject> groupList = new ArrayList<>();
    private String roleField;
    private String identifierHeaderName;

    private boolean useOpac = false;
    private String opacName;
    private String opacHeader;
    private String searchField;

    private String processtitleRule;

    private String imageFolderPath;
    private String imageFolderHeaderName;
    private String listSplitChar;

    private boolean ignoreImages;
    private boolean moveImage;
    private boolean runAsGoobiScript;
    private List<MetadataMappingObject> rolesList = new ArrayList<>();

    private String imageHandlingStrategy; //copy, move or ignore
    private boolean failOnMissingImageFiles = false;

    /**
     * loads the &lt;config&gt; block from xml file
     * 
     * @param xmlConfig
     */

    public ExcelConfig(SubnodeConfiguration xmlConfig) {
        if (xmlConfig == null) {
            return;
        }

        publicationType = xmlConfig.getString("/publicationType", "Monograph");
        collection = xmlConfig.getString("/collection", "");
        firstLine = xmlConfig.getInt("/firstLine", 1);
        identifierColumn = xmlConfig.getInt("/identifierColumn", 1);
        conditionalColumn = xmlConfig.getInt("/conditionalColumn", identifierColumn);
        identifierHeaderName = xmlConfig.getString("/identifierHeaderName", null);
        rowHeader = xmlConfig.getInt("/rowHeader", 1);
        rowDataStart = xmlConfig.getInt("/rowDataStart", 2);
        rowDataEnd = xmlConfig.getInt("/rowDataEnd", 20000);

        processtitleRule = xmlConfig.getString("/processTitleRule", null);

        List<HierarchicalConfiguration> iml = xmlConfig.configurationsAt("//importImages");

        for (HierarchicalConfiguration md : iml) {

            List<ConfigurationNode> attr = md.getRootNode().getAttributes("failOnMissingImageFiles");

            if (attr != null && attr.size() > 0) {
                failOnMissingImageFiles = attr.get(0).getValue().toString().contentEquals("true");
            }

            imageFolderPath = md.getString("/imageFolderPath", null);
            imageFolderHeaderName = md.getString("/imageFolderHeaderName", null);
            imageHandlingStrategy = md.getString("/imageHandlingStrategy", "ignore");

            //only allow "copy"; "move" and "ignore" for imageHandlingStrategy:
            if (!imageHandlingStrategy.contentEquals("copy") && !imageHandlingStrategy.contentEquals("move")
                    && !imageHandlingStrategy.contentEquals("ignore")) {
                imageHandlingStrategy = "ignore";
            }
            moveImage = imageHandlingStrategy.contentEquals("move");
        }

        runAsGoobiScript = xmlConfig.getBoolean("/runAsGoobiScript", true);
        listSplitChar = xmlConfig.getString("/splitList", null);
        
        roleField = xmlConfig.getString("/roleField", null);
        
        List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
        for (HierarchicalConfiguration md : mml) {
            metadataList.add(getMetadata(md));
        }
        
        List<HierarchicalConfiguration> mmr = xmlConfig.configurationsAt("//role");
        for (HierarchicalConfiguration md : mmr) {
            rolesList.add(getRoleMetadata(md));
        }

       
        List<HierarchicalConfiguration> pml = xmlConfig.configurationsAt("//person");
        for (HierarchicalConfiguration md : pml) {
            personList.add(getPersons(md));
        }
        
        List<HierarchicalConfiguration> pmlr = xmlConfig.configurationsAt("//person-role");
        for (HierarchicalConfiguration md : pmlr) {
            personWithRoleList.add(getPersonsWithRoles(md));
        }

        List<HierarchicalConfiguration> gml = xmlConfig.configurationsAt("//group");
        for (HierarchicalConfiguration md : gml) {
            String rulesetName = md.getString("@ugh");
            GroupMappingObject grp = new GroupMappingObject();
            grp.setRulesetName(rulesetName);

            String docType = md.getString("@docType", "child");
            grp.setDocType(docType);
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
        useOpac = xmlConfig.getBoolean("/useOpac", false);
        if (useOpac) {
            opacName = xmlConfig.getString("/opacName", "ALMA WUW");
            opacHeader = xmlConfig.getString("/opacHeader", "");
            searchField = xmlConfig.getString("/searchField", "12");
        }
    }

    private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        String propertyName = md.getString("@property");
        Integer columnNumber = md.getInteger("@column", null);
        //        Integer identifierColumn = md.getInteger("@identifier", null);
        String headerName = md.getString("@headerName", null);
        String normdataHeaderName = md.getString("@normdataHeaderName", null);
        String docType = md.getString("@docType", "child");

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setExcelColumn(columnNumber);
        //        mmo.setIdentifierColumn(identifierColumn);
        mmo.setPropertyName(propertyName);
        mmo.setRulesetName(rulesetName);
        mmo.setHeaderName(headerName);
        mmo.setNormdataHeaderName(normdataHeaderName);
        mmo.setDocType(docType);

        mmo.setSearchField(md.getString("@opacSearchField", null));
        return mmo;
    }

    private MetadataMappingObject getRoleMetadata(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        String propertyName = md.getString("@roleName");
        String docType = md.getString("@docType", "child");

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setPropertyName(propertyName);
        mmo.setRulesetName(rulesetName);
        mmo.setDocType(docType);
        return mmo;
    }
    
    private PersonMappingObject getPersons(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        Integer firstname = md.getInteger("firstname", null);
        Integer lastname = md.getInteger("lastname", null);
        Integer identifier = md.getInteger("identifier", null);
        String headerName = md.getString("nameFieldHeader", null);
        String firstnameHeaderName = md.getString("firstnameFieldHeader", null);
        String lastnameHeaderName = md.getString("lastnameFieldHeader", null);
        String normdataHeaderName = md.getString("@normdataHeaderName", null);
        boolean splitName = md.getBoolean("splitName", false);
        String splitChar = md.getString("splitChar", " ");
        boolean firstNameIsFirstPart = md.getBoolean("splitName/@firstNameIsFirstPart", false);
        String splitList = md.getString("splitList", null);
        String gndIds = md.getString("gndIds", null);
        
        String docType = md.getString("@docType", "child");
        String useRoleField =  md.getString("useRoleField", null);

        PersonMappingObject pmo = new PersonMappingObject();
        pmo.setFirstnameColumn(firstname);
        pmo.setLastnameColumn(lastname);
        pmo.setIdentifierColumn(identifier);
        pmo.setRulesetName(rulesetName);
        pmo.setHeaderName(headerName);
        pmo.setNormdataHeaderName(normdataHeaderName);

        pmo.setFirstnameHeaderName(firstnameHeaderName);
        pmo.setLastnameHeaderName(lastnameHeaderName);
        pmo.setSplitChar(splitChar);
        pmo.setSplitName(splitName);
        pmo.setFirstNameIsFirst(firstNameIsFirstPart);
        pmo.setDocType(docType);

        pmo.setSplitList(splitList);
        pmo.setGndIds(gndIds);
        pmo.setUseRoleField(useRoleField);
        
        return pmo;

    }

    private PersonMappingObject getPersonsWithRoles(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        Integer firstname = md.getInteger("firstname", null);
        Integer lastname = md.getInteger("lastname", null);
        Integer identifier = md.getInteger("identifier", null);
        String headerName = md.getString("nameFieldHeader", null);
        String firstnameHeaderName = md.getString("firstnameFieldHeader", null);
        String lastnameHeaderName = md.getString("lastnameFieldHeader", null);
        String normdataHeaderName = md.getString("@normdataHeaderName", null);
        boolean splitName = md.getBoolean("splitName", false);
        String splitChar = md.getString("splitChar", " ");
        boolean firstNameIsFirstPart = md.getBoolean("splitName/@firstNameIsFirstPart", false);
        String splitList = md.getString("splitList", null);
        String splitRole = md.getString("splitRole", null);
        String gndIds = md.getString("gndIds", null);
        
        String docType = md.getString("@docType", "child");
        
        PersonMappingObject pmo = new PersonMappingObject();
        pmo.setFirstnameColumn(firstname);
        pmo.setLastnameColumn(lastname);
        pmo.setIdentifierColumn(identifier);
        pmo.setRulesetName(rulesetName);
        pmo.setHeaderName(headerName);
        pmo.setNormdataHeaderName(normdataHeaderName);

        pmo.setFirstnameHeaderName(firstnameHeaderName);
        pmo.setLastnameHeaderName(lastnameHeaderName);
        pmo.setSplitChar(splitChar);
        pmo.setSplitName(splitName);
        pmo.setSplitRole(splitRole);
        pmo.setFirstNameIsFirst(firstNameIsFirstPart);
        pmo.setDocType(docType);

        pmo.setSplitList(splitList);
        pmo.setGndIds(gndIds);
        
        return pmo;

    }
}
