package de.intranda.goobi.plugins;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Processproperty;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.enums.ImportType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.importer.DocstructElement;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.PluginLoader;
import org.goobi.production.plugin.interfaces.IImportPluginVersion2;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.plugin.interfaces.IPlugin;
import org.goobi.production.properties.ImportProperty;

import de.intranda.goobi.plugins.util.Config;
import de.intranda.goobi.plugins.util.GroupMappingObject;
import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.goobi.plugins.util.PersonMappingObject;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
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
public class HeaderExcelImport implements IImportPluginVersion2, IPlugin {

    private Prefs prefs;
    private MassImportForm form;
    private String importFolder;
    private File file;
    private String data;
    private String currentIdentifier;
    private List<String> currentCollections = new ArrayList<>();
    private String ats;
    private String volumeNumber;
    private String processTitle;

    private String title = "intranda_import_excel_read_headerdata";

    //    private Map<String, Integer> headerOrder;

    private List<ImportType> importTypes;
    private String workflowTitle;
    private Config config;

    public HeaderExcelImport() {
        importTypes = new ArrayList<>();
        importTypes.add(ImportType.FILE);
    }

    @Override
    public Fileformat convertData() throws ImportPluginException {
        return null;
    }

    @Override
    public void setData(Record r) {
    }

    private Fileformat getRecordFromCatalogue(String identifier) throws ImportPluginException {
        ConfigOpacCatalogue coc = ConfigOpac.getInstance().getCatalogueByName(config.getOpacName());
        if (coc == null) {
            throw new ImportPluginException("Catalogue with name " + config.getOpacName() + " not found. Please check goobi_opac.xml");
        }
        IOpacPlugin myImportOpac = (IOpacPlugin) PluginLoader.getPluginByTitle(PluginType.Opac, coc.getOpacType());
        if (myImportOpac == null) {
            throw new ImportPluginException("Opac plugin " + coc.getOpacType() + " not found. Abort.");
        }
        Fileformat myRdf = null;
        try {
            myRdf = myImportOpac.search(config.getSearchField(), identifier, coc, prefs);
            if (myRdf == null) {
                throw new ImportPluginException("Could not import record " + identifier
                        + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
            }
        } catch (Exception e1) {
            throw new ImportPluginException("Could not import record " + identifier
                    + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
        }
        DocStruct ds = null;
        DocStruct anchor = null;
        try {
            ds = myRdf.getDigitalDocument().getLogicalDocStruct();
            if (ds.getType().isAnchor()) {
                anchor = ds;
                if (ds.getAllChildren() == null || ds.getAllChildren().isEmpty()) {
                    throw new ImportPluginException(
                            "Could not import record " + identifier + ". Found anchor file, but no children. Try to import the child record.");
                }
                ds = ds.getAllChildren().get(0);
            }
        } catch (PreferencesException e1) {
            throw new ImportPluginException("Could not import record " + identifier
                    + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
        }
        try {
            ats = myImportOpac.getAtstsl();

            List<? extends Metadata> sort = ds.getAllMetadataByType(prefs.getMetadataTypeByName("CurrentNoSorting"));
            if (sort != null && !sort.isEmpty()) {
                volumeNumber = sort.get(0).getValue();
            }

        } catch (Exception e) {
            ats = "";
        }

        return myRdf;
    }


    @SuppressWarnings("unchecked")
    @Override
    public List<ImportObject> generateFiles(List<Record> records) {
        List<ImportObject> answer = new ArrayList<>();

        for (Record record : records) {
            ImportObject io = new ImportObject();

            try {

                Object tempObject = record.getObject();

                List<Map<?, ?>> list = (List<Map<?, ?>>) tempObject;
                Map<String, Integer> headerOrder = (Map<String, Integer>) list.get(0);
                Map<Integer, String> rowMap = (Map<Integer, String>) list.get(1);

                // generate a mets file
                DigitalDocument digitalDocument = null;
                Fileformat ff = null;
                DocStruct logical = null;
                DocStruct anchor = null;
                if (!config.isUseOpac()) {
                    ff = new MetsMods(prefs);
                    digitalDocument = new DigitalDocument();
                    ff.setDigitalDocument(digitalDocument);
                    String publicationType = getConfig().getPublicationType();
                    DocStructType logicalType = prefs.getDocStrctTypeByName(publicationType);
                    logical = digitalDocument.createDocStruct(logicalType);
                    digitalDocument.setLogicalDocStruct(logical);
                    answer.add(io);
                } else {
                    try {
                        if (StringUtils.isBlank(config.getIdentifierHeaderName())) {
                            Helper.setFehlerMeldung("Cannot request catalogue, no identifier column defined");
                            return Collections.emptyList();
                        }

                        String catalogueIdentifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));
                        if (StringUtils.isBlank(catalogueIdentifier)) {
                            continue;
                        }
                        ff = getRecordFromCatalogue(catalogueIdentifier);
                        digitalDocument = ff.getDigitalDocument();
                        logical = digitalDocument.getLogicalDocStruct();
                        if (logical.getType().isAnchor()) {
                            anchor = logical;
                            logical = anchor.getAllChildren().get(0);
                        }
                        answer.add(io);
                    } catch (ImportPluginException e) {
                        log.error(e);
                        io.setErrorMessage(e.getMessage());
                        io.setImportReturnValue(ImportReturnValue.NoData);
                        continue;
                    }
                }

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

                io.setImportReturnValue(ImportReturnValue.ExportFinished);

                for (MetadataMappingObject mmo : getConfig().getMetadataList()) {

                    String value = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                    String identifier = null;
                    if (mmo.getNormdataHeaderName() != null) {
                        identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
                    }
                    if (StringUtils.isNotBlank(mmo.getRulesetName()) && StringUtils.isNotBlank(value)) {
                        try {
                            Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            md.setValue(value);
                            if (identifier != null) {
                                md.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);

                            }
                            if (anchor != null && "anchor".equals(mmo.getDocType())) {
                                anchor.addMetadata(md);
                            } else {
                                logical.addMetadata(md);
                            }
                        } catch (MetadataTypeNotAllowedException e) {
                            log.info(e);
                            // Metadata is not known or not allowed
                        }
                        // create a default title
                        if (mmo.getRulesetName().equalsIgnoreCase("CatalogIDDigital") && !"anchor".equals(mmo.getDocType())) {
                            fileName = getImportFolder() + File.separator + value + ".xml";
                            io.setProcessTitle(value);
                            io.setMetsFilename(fileName);
                        }
                    }
                    if (StringUtils.isNotBlank(config.getProcesstitleRule())) {
                        StringBuilder title = new StringBuilder();
                        StringTokenizer tokenizer = new StringTokenizer(config.getProcesstitleRule(), "+");
                        while (tokenizer.hasMoreTokens()) {
                            String myString = tokenizer.nextToken();
                            /*
                             * wenn der String mit ' anfängt und mit ' endet, dann den Inhalt so übernehmen
                             */
                            if (myString.startsWith("'") && myString.endsWith("'")) {
                                title.append(myString.substring(1, myString.length() - 1));
                            } else {
                                if (myString.equalsIgnoreCase("Signatur") || myString.equalsIgnoreCase("Shelfmark")) {
                                    if (StringUtils.isNotBlank(rowMap.get(headerOrder.get(myString)))) {
                                        // replace white spaces with dash, remove other special characters
                                        title.append(
                                                rowMap.get(headerOrder.get(myString)).replace(" ", "-").replace("/", "-").replaceAll("[^\\w-]", ""));
                                    }
                                } else {
                                    title.append(rowMap.get(headerOrder.get(myString)));
                                }

                            }
                        }
                        String newTitle = title.toString();
                        if (newTitle.endsWith("_")) {
                            newTitle = newTitle.substring(0, newTitle.length() - 1);
                        }
                        // remove non-ascii characters for the sake of TIFF header limits
                        String regex = ConfigurationHelper.getInstance().getProcessTitleReplacementRegex();

                        String filteredTitle = newTitle.replaceAll(regex, "");

                        // set new process title
                        fileName = getImportFolder() + File.separator + filteredTitle + ".xml";
                        io.setProcessTitle(filteredTitle);
                        io.setMetsFilename(fileName);

                    }

                    if (StringUtils.isNotBlank(mmo.getPropertyName()) && StringUtils.isNotBlank(value)) {
                        Processproperty p = new Processproperty();
                        p.setTitel(mmo.getPropertyName());
                        p.setWert(value);
                        io.getProcessProperties().add(p);
                    }
                }

                for (PersonMappingObject mmo : getConfig().getPersonList()) {
                    String firstname = "";
                    String lastname = "";
                    if (mmo.isSplitName()) {
                        String name = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                        if (StringUtils.isNotBlank(name)) {
                            if (name.contains(mmo.getSplitChar())) {
                                if (mmo.isFirstNameIsFirst()) {
                                    firstname = name.substring(0, name.lastIndexOf(mmo.getSplitChar()));
                                    lastname = name.substring(name.lastIndexOf(mmo.getSplitChar()));
                                } else {
                                    lastname = name.substring(0, name.lastIndexOf(mmo.getSplitChar())).trim();
                                    firstname = name.substring(name.lastIndexOf(mmo.getSplitChar()) + 1).trim();
                                }
                            } else {
                                lastname = name;
                            }
                        }
                    } else {
                        firstname = rowMap.get(headerOrder.get(mmo.getFirstnameHeaderName()));
                        lastname = rowMap.get(headerOrder.get(mmo.getLastnameHeaderName()));
                    }

                    String identifier = null;
                    if (mmo.getNormdataHeaderName() != null) {
                        identifier = rowMap.get(headerOrder.get(mmo.getNormdataHeaderName()));
                    }
                    if (StringUtils.isNotBlank(mmo.getRulesetName())) {
                        try {
                            Person p = new Person(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            p.setFirstname(firstname);
                            p.setLastname(lastname);

                            if (identifier != null) {
                                p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
                            }
                            if (anchor != null && "anchor".equals(mmo.getDocType())) {
                                anchor.addPerson(p);
                            } else {
                                logical.addPerson(p);
                            }

                            //                            logical.addPerson(p);
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
                            String value = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                            Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));
                            md.setValue(value);
                            if (mmo.getNormdataHeaderName() != null) {
                                md.setAutorityFile("gnd", "http://d-nb.info/gnd/", rowMap.get(headerOrder.get(mmo.getNormdataHeaderName())));
                            }
                            group.addMetadata(md);
                        }
                        for (PersonMappingObject pmo : gmo.getPersonList()) {
                            Person p = new Person(prefs.getMetadataTypeByName(pmo.getRulesetName()));
                            String firstname = "";
                            String lastname = "";
                            if (pmo.isSplitName()) {
                                String name = rowMap.get(headerOrder.get(pmo.getHeaderName()));
                                if (StringUtils.isNotBlank(name)) {
                                    if (name.contains(pmo.getSplitChar())) {
                                        if (pmo.isFirstNameIsFirst()) {
                                            firstname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
                                            lastname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
                                        } else {
                                            lastname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
                                            firstname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
                                        }
                                    } else {
                                        lastname = name;
                                    }
                                }
                            } else {
                                firstname = rowMap.get(headerOrder.get(pmo.getFirstnameHeaderName()));
                                lastname = rowMap.get(headerOrder.get(pmo.getLastnameHeaderName()));
                            }

                            p.setFirstname(firstname);
                            p.setLastname(lastname);

                            if (pmo.getNormdataHeaderName() != null) {
                                p.setAutorityFile("gnd", "http://d-nb.info/gnd/", rowMap.get(headerOrder.get(pmo.getNormdataHeaderName())));
                            }
                            group.addMetadata(p);
                        }
                        if (anchor != null && "anchor".equals(gmo.getDocType())) {
                            anchor.addMetadataGroup(group);
                        } else {
                            logical.addMetadataGroup(group);
                        }

                        //                        logical.addMetadataGroup(group);

                    } catch (MetadataTypeNotAllowedException e) {
                        log.info(e);
                        // Metadata is not known or not allowed
                    }
                }

                // write mets file into import folder
                ff.write(fileName);

                if (StringUtils.isNotBlank(config.getImageFolderHeaderName())
                        && StringUtils.isNotBlank(rowMap.get(headerOrder.get(config.getImageFolderHeaderName())))) {

                    Path imageSourceFolder = null;
                    if (config.getImageFolderPath() != null) {
                        imageSourceFolder = Paths.get(config.getImageFolderPath(), rowMap.get(headerOrder.get(config.getImageFolderHeaderName())));
                    } else {
                        imageSourceFolder = Paths.get(rowMap.get(headerOrder.get(config.getImageFolderHeaderName())));
                    }
                    if (Files.exists(imageSourceFolder) && Files.isDirectory(imageSourceFolder)) {

                        // folder name
                        String foldername = fileName.replace(".xml", "");
                        Path path = Paths.get(foldername, "images", "master_" + io.getProcessTitle() + "_media");
                        try {
                            Files.createDirectories(path.getParent());
                            if (config.isMoveImage()) {
                                StorageProvider.getInstance().move(imageSourceFolder, path);
                            }else {
                                StorageProvider.getInstance().copyDirectory(imageSourceFolder, path);
                            }
                        } catch (IOException e) {
                            log.error(e);
                        }

                    }
                }
            } catch (WriteException | PreferencesException | MetadataTypeNotAllowedException | TypeNotAllowedForParentException e) {
                io.setImportReturnValue(ImportReturnValue.WriteError);
                io.setErrorMessage(e.getMessage());
            }

        }
        // end of all excel rows
        return answer;
    }

    @Override
    public List<Record> splitRecords(String records) {
        return null;
    }

    @Override
    public List<Record> generateRecordsFromFile() {

        if (StringUtils.isBlank(workflowTitle)) {
            workflowTitle = form.getTemplate().getTitel();
        }
        config = null;
        List<Record> recordList = new ArrayList<>();
        String idColumn = getConfig().getIdentifierHeaderName();
        Map<String, Integer> headerOrder = new HashMap<>();

        InputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(file);
            BOMInputStream in = new BOMInputStream(fileInputStream, false);
            Workbook wb = WorkbookFactory.create(in);
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();

            // get header and data row number from config first
            int rowHeader = getConfig().getRowHeader();
            int rowDataStart = getConfig().getRowDataStart();
            int rowDataEnd = getConfig().getRowDataEnd();
            int rowCounter = 0;

            //  find the header row
            Row headerRow = null;
            while (rowCounter < rowHeader) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            //  read and validate the header row
            int numberOfCells = headerRow.getLastCellNum();
            for (int i = 0; i < numberOfCells; i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    String value = cell.getStringCellValue();
                    headerOrder.put(value, i);
                }
            }

            // find out the first data row
            while (rowCounter < rowDataStart - 1) {
                headerRow = rowIterator.next();
                rowCounter++;
            }

            // run through all the data rows
            while (rowIterator.hasNext() && rowCounter < rowDataEnd) {
                Map<Integer, String> map = new HashMap<>();
                Row row = rowIterator.next();
                rowCounter++;
                int lastColumn = row.getLastCellNum();
                if (lastColumn == -1) {
                    continue;
                }
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
                            //                            value = cell.getCellFormula();
                            value = cell.getRichStringCellValue().getString();
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
                    map.put(cn, value);

                }

                // just add the record if any column contains a value
                for (String v : map.values()) {
                    if (v != null && !v.isEmpty()) {
                        Record r = new Record();
                        r.setId(map.get(headerOrder.get(idColumn)));
                        List<Map<?, ?>> list = new ArrayList<>();
                        list.add(headerOrder);
                        list.add(map);

                        r.setObject(list);
                        recordList.add(r);
                        break;
                    }
                }
            }

        } catch (Exception e) {
            log.error(e);
        } finally {
            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    log.error(e);
                }
            }
        }

        return recordList;
    }

    @Override
    public List<Record> generateRecordsFromFilenames(List<String> filenames) {
        return null;
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

    @Override
    public boolean isRunnableAsGoobiScript() {
        return config.isRunAsGoobiScript();
    }

    public Config getConfig() {
        if (config == null) {
            config = loadConfig(workflowTitle);
        }
        return config;
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
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(title);
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
