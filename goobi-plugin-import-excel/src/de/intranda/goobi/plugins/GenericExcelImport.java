package de.intranda.goobi.plugins;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.goobi.beans.Processproperty;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.enums.ImportType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.importer.DocstructElement;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.interfaces.IImportPlugin;
import org.goobi.production.plugin.interfaces.IImportPluginVersion2;
import org.goobi.production.plugin.interfaces.IPlugin;
import org.goobi.production.properties.ImportProperty;

import de.intranda.goobi.plugins.util.Config;
import de.intranda.goobi.plugins.util.GroupMappingObject;
import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.goobi.plugins.util.PersonMappingObject;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import lombok.Data;
import lombok.extern.log4j.Log4j;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.DocStructType;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataGroup;
import ugh.dl.Person;
import ugh.dl.Prefs;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.TypeNotAllowedForParentException;
import ugh.exceptions.WriteException;
import ugh.fileformats.mets.MetsMods;

@Log4j
@Data
@PluginImplementation
public class GenericExcelImport implements IImportPlugin, IPlugin {

    private Prefs prefs;
    private Record data;
    private String importFolder;
    private String processTitle;
    private MassImportForm form;
    private List<ImportType> importTypes = new ArrayList<>();
    private String title = "generic_import_excel";
    private File importFile;
    private String workflowTitle;

    private Config config;

    public GenericExcelImport() {
        importTypes.add(ImportType.FILE);
    }

    public Config getConfig() {
        if (config == null) {
            config = loadConfig(workflowTitle);
        }
        return config;
    }

    @Override
    public Fileformat convertData() throws ImportPluginException {
        return null;
    }

    @Override
    public List<ImportObject> generateFiles(List<Record> records) {
        List<ImportObject> answer = new ArrayList<ImportObject>();

        for (Record record : records) {
            ImportObject io = new ImportObject();
            try {
                // generate a mets file
                Fileformat ff = new MetsMods(prefs);
                DigitalDocument digitalDocument = new DigitalDocument();
                ff.setDigitalDocument(digitalDocument);

                String publicationType = getConfig().getPublicationType();
                DocStructType logicalType = prefs.getDocStrctTypeByName(publicationType);
                DocStruct logical = digitalDocument.createDocStruct(logicalType);
                digitalDocument.setLogicalDocStruct(logical);
                DocStructType physicalType = prefs.getDocStrctTypeByName("BoundBook");
                DocStruct physical = digitalDocument.createDocStruct(physicalType);
                digitalDocument.setPhysicalDocStruct(physical);
                Metadata imagePath = new Metadata(prefs.getMetadataTypeByName("pathimagefiles"));
                imagePath.setValue("./images/");
                physical.addMetadata(imagePath);

                // add collections if configured
                String col = getConfig().getCollection();
                if (StringUtils.isNotBlank(col)) {
                    Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
                    mdColl.setValue(col);
                    logical.addMetadata(mdColl);
                }
                // and add all collections that where selected
                for (String colItem : form.getDigitalCollections()) {
                    if (!colItem.equals(col.trim())) {
                        Metadata mdColl = new Metadata(prefs.getMetadataTypeByName("singleDigCollection"));
                        mdColl.setValue(colItem);
                        logical.addMetadata(mdColl);
                    }
                }
                // create file name for mets file
                String fileName = null;

                // create importobject for massimport
                io.setProcessTitle(record.getId());
                io.setImportReturnValue(ImportReturnValue.ExportFinished);

                Object tempObject = record.getObject();
                Map<Integer, String> map = (Map<Integer, String>) tempObject;

                for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                    String value = map.get(mmo.getExcelColumn());
                    String identifier = null;
                    if (mmo.getIdentifierColumn() != null) {
                        identifier = map.get(mmo.getIdentifierColumn());
                    }
                    if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                        try {
                            Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            md.setValue(value);
                            if (identifier != null) {
                                md.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);

                            }
                            logical.addMetadata(md);
                        } catch (MetadataTypeNotAllowedException e) {
                            log.info(e);
                            // Metadata is not known or not allowed
                        }

                        if (mmo.getRulesetName().equalsIgnoreCase("CatalogIDDigital")) {
                            fileName = getImportFolder() + File.separator + value + ".xml";
                            io.setProcessTitle(value);
                            io.setMetsFilename(fileName);
                        }
                    }

                    if (StringUtils.isNotBlank(mmo.getPropertyName())) {
                        Processproperty p = new Processproperty();
                        p.setTitel(mmo.getPropertyName());
                        p.setWert(value);
                        io.getProcessProperties().add(p);
                    }
                }

                for (PersonMappingObject mmo : getConfig().getPersonList()) {
                    String firstname = map.get(mmo.getFirstnameColumn());
                    String lastname = map.get(mmo.getLastnameColumn());
                    String identifier = null;
                    if (mmo.getIdentifierColumn() != null) {
                        identifier = map.get(mmo.getIdentifierColumn());
                    }
                    if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                        try {
                            Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            p.setFirstname(firstname);
                            p.setLastname(lastname);

                            if (identifier != null) {
                                p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
                            }
                            logical.addPerson(p);
                        } catch (MetadataTypeNotAllowedException e) {
                            log.info(e);
                            // Metadata is not known or not allowed
                        }
                    }
                }

                for (GroupMappingObject gmo : getConfig().getGroupList()) {
                    try {
                        MetadataGroup group = new MetadataGroup(prefs.getMetadataGroupTypeByName(gmo.getRulesetName()));
                        for (MetadataMappingObject mmo : gmo.getMetadataList()) {
                            String value = map.get(mmo.getExcelColumn());
                            Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            md.setValue(value);
                            if (mmo.getIdentifierColumn() != null) {
                                md.setAutorityFile("gnd", "http://d-nb.info/gnd/", map.get(mmo.getIdentifierColumn()));
                            }
                            group.addMetadata(md);
                        }
                        for (PersonMappingObject pmo : gmo.getPersonList()) {
                            Person p = new Person(prefs.getMetadataTypeByName(pmo.getRulesetName()));
                            p.setFirstname(map.get(pmo.getFirstnameColumn()));
                            p.setLastname(map.get(pmo.getLastnameColumn()));

                            if (pmo.getIdentifierColumn() != null) {
                                p.setAutorityFile("gnd", "http://d-nb.info/gnd/", map.get(pmo.getIdentifierColumn()));
                            }
                            group.addMetadata(p);
                        }
                        logical.addMetadataGroup(group);

                    } catch (MetadataTypeNotAllowedException e) {
                        log.info(e);
                        // Metadata is not known or not allowed
                    }
                }

                // write mets file into import folder
                ff.write(fileName);
            } catch (WriteException | PreferencesException | MetadataTypeNotAllowedException | TypeNotAllowedForParentException e) {
                io.setImportReturnValue(ImportReturnValue.WriteError);
            }
            answer.add(io);
        }
        // end of all excel rows
        return answer;

    }

    @Override
    public List<Record> generateRecordsFromFile() {
        if (StringUtils.isBlank(workflowTitle)) {
            workflowTitle = form.getTemplate().getTitel();
        }
        boolean ignoreFirstLine = getConfig().isIgnoreFirstLine();
        List<Record> recordList = new ArrayList<>();
        InputStream file = null;
        try {
            file = new FileInputStream(importFile);

            BOMInputStream in = new BOMInputStream(file, false);

            Workbook wb = WorkbookFactory.create(in);

            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();

            if (ignoreFirstLine) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Map<Integer, String> map = new HashMap<>();
                Row row = rowIterator.next();
                int lastColumn = row.getLastCellNum();
                Integer i = 1;
                for (int cn = 0; cn < lastColumn; cn++) {
                    //                while (cellIterator.hasNext()) {
                    //                    Cell cell = cellIterator.next();
                    Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value = "";
                    switch (cell.getCellTypeEnum()) {
                        case BOOLEAN:
                            value = cell.getBooleanCellValue() ? "true" : "false";
                            break;
                        case FORMULA:
                            value = cell.getCellFormula();
                            break;
                        case NUMERIC:
                            value = String.valueOf((int) cell.getNumericCellValue());
                            break;
                        case STRING:
                            value = cell.getStringCellValue();
                            break;
                        default:
                            // none, error, blank
                            value = "";
                            break;
                    }
                    map.put(i, value);
                    i++;
                }
                Record r = new Record();
                r.setId(map.get(1));
                r.setObject(map);
                recordList.add(r);
            }

        } catch (InvalidFormatException | IOException e) {
            log.error(e);

        } finally {
            if (file != null) {
                try {
                    file.close();
                } catch (IOException e) {
                    log.error(e);
                }
            }
        }
        return recordList;
    }

    @Override
    public List<Record> splitRecords(String records) {
        return null;
    }

    @Override
    public List<Record> generateRecordsFromFilenames(List<String> filenames) {
        return null;
    }

    @Override
    public void setFile(File importFile) {
        this.importFile = importFile;

    }

    @Override
    public List<String> splitIds(String ids) {
        return null;
    }

    @Override
    public List<ImportProperty> getProperties() {
        return null;
    }

    @Override
    public List<String> getAllFilenames() {
        return null;
    }

    @Override
    public void deleteFiles(List<String> selectedFilenames) {
    }

    @Override
    public List<? extends DocstructElement> getCurrentDocStructs() {
        return null;
    }

    @Override
    public String deleteDocstruct() {
        return null;
    }

    @Override
    public String addDocstruct() {
        return null;
    }

    @Override
    public List<String> getPossibleDocstructs() {
        return null;
    }

    @Override
    public DocstructElement getDocstruct() {
        return null;
    }

    @Override
    public void setDocstruct(DocstructElement dse) {
    }

    @Override
    public PluginType getType() {
        return PluginType.Import;
    }

    public boolean isRunnableAsGoobiScript() {
        return true;
    }

    /**
     * Loads the configuration for the selected template or the default configuration, if the template was not specified.
     * 
     * The configuration is stored in a {@link Config} object
     * 
     * @param workflowTitle
     * @return
     */

    private Config loadConfig(String workflowTitle) {
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(this);
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());
        xmlConfig.setReloadingStrategy(new FileChangedReloadingStrategy());

        SubnodeConfiguration myconfig = null;
        try {

            myconfig = xmlConfig.configurationAt("//config[./template = '" + workflowTitle + "']");
        } catch (IllegalArgumentException e) {
            myconfig = xmlConfig.configurationAt("//config[./template = '*']");
        }
        Config config = new Config(myconfig);

        return config;
    }

}
