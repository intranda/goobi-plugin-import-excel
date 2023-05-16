package de.intranda.goobi.plugins;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.configuration.SubnodeConfiguration;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.reloading.FileChangedReloadingStrategy;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.beans.Process;
import org.goobi.beans.Processproperty;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.enums.ImportType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.importer.DocstructElement;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.interfaces.IImportPluginVersion2;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.plugin.interfaces.IPlugin;
import org.goobi.production.properties.ImportProperty;

import de.intranda.goobi.plugins.util.ExcelConfig;
import de.intranda.goobi.plugins.util.GroupMappingObject;
import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.goobi.plugins.util.PersonMappingObject;
import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.Helper;
import de.sub.goobi.helper.StorageProvider;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import de.sub.goobi.helper.exceptions.SwapException;
import de.sub.goobi.persistence.managers.ProcessManager;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import lombok.Data;
import lombok.EqualsAndHashCode;
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
import ugh.exceptions.UGHException;
import ugh.exceptions.WriteException;
import ugh.fileformats.mets.MetsMods;

@Log4j
@Data
@PluginImplementation
public class GenericExcelImport implements IImportPluginVersion2, IPlugin {

    private static final long serialVersionUID = 3965077868027995218L;

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

    @EqualsAndHashCode.Exclude
    private boolean replaceExisting = false;
    @EqualsAndHashCode.Exclude
    private boolean moveFiles = false;

    private String title = "intranda_import_excel";

    private List<ImportType> importTypes;
    private String workflowTitle;

    @EqualsAndHashCode.Exclude
    private transient ExcelConfig config;

    public GenericExcelImport() {
        importTypes = new ArrayList<>();
        importTypes.add(ImportType.FILE);

        getConfig();
    }

    @Override
    public Fileformat convertData() throws ImportPluginException {
        return null;
    }

    @Override
    public void setData(Record r) {
        // do nothing
    }

    private Fileformat getRecordFromCatalogue(Map<Integer, String> rowMap, Map<String, Integer> headerOrder, String catalogue)
            throws ImportPluginException {
        IOpacPlugin myImportOpac = null;
        ConfigOpacCatalogue coc = null;
        for (ConfigOpacCatalogue configOpacCatalogue : ConfigOpac.getInstance().getAllCatalogues(workflowTitle)) {
            if (configOpacCatalogue.getTitle().equals(catalogue)) {
                myImportOpac = configOpacCatalogue.getOpacPlugin();
                coc = configOpacCatalogue;
            }
        }
        if (myImportOpac == null) {
            throw new ImportPluginException("Opac plugin for catalogue " + catalogue + " not found. Abort.");
        }
        Fileformat myRdf = null;
        DocStruct ds = null;
        if ("intranda_opac_json".equals(myImportOpac.getTitle())) {

            /**
             * 
             * JsonOpacPlugin jsonOpacPlugin = (JsonOpacPlugin) myImportOpac; de.intranda.goobi.plugins.util.Config jsonOpacConfig =
             * jsonOpacPlugin.getConfigForOpac(); for (MetadataMappingObject mmo : config.getMetadataList()) { if
             * (StringUtils.isNotBlank(mmo.getSearchField())) { for (SearchField sf : jsonOpacConfig.getFieldList()) { if
             * ((sf.getId()).equals(mmo.getSearchField())) { String value = rowMap.get(headerOrder.get(mmo.getHeaderName())); if
             * (StringUtils.isNotBlank(value)) { sf.setText(value); sf.setSelectedField(mmo.getHeaderName()); } } } } }
             * 
             * Direct access to the classes is not possible because of different class loaders. Replace code above with reflections:
             */

            try {
                Class<? extends Object> opacClass = myImportOpac.getClass();
                Method getConfigForOpac = opacClass.getMethod("getConfigForOpac");
                Object jsonOpacConfig = getConfigForOpac.invoke(myImportOpac);

                Class<? extends Object> jsonOpacConfigClass = jsonOpacConfig.getClass();

                Method getFieldList = jsonOpacConfigClass.getMethod("getFieldList");

                Object fieldList = getFieldList.invoke(jsonOpacConfig);
                @SuppressWarnings("unchecked")
                List<Object> searchfields = (List<Object>) fieldList;
                for (MetadataMappingObject mmo : config.getMetadataList()) {
                    if (StringUtils.isNotBlank(mmo.getSearchField())) {
                        for (Object searchField : searchfields) {
                            Class<? extends Object> searchFieldClass = searchField.getClass();

                            Method getId = searchFieldClass.getMethod("getId");

                            Method setText = searchFieldClass.getMethod("setText", String.class);
                            Method setSelectedField = searchFieldClass.getMethod("setSelectedField", String.class);

                            Object id = getId.invoke(searchField);
                            if (((String) id).equals(mmo.getSearchField())) {
                                String value = rowMap.get(headerOrder.get(mmo.getHeaderName()));
                                if (StringUtils.isNotBlank(value)) {
                                    setText.invoke(searchField, value);
                                    setSelectedField.invoke(searchField, mmo.getHeaderName());
                                }
                            }
                        }
                    }
                }
                Method search = opacClass.getMethod("search", String.class, String.class, ConfigOpacCatalogue.class, Prefs.class);

                myRdf = (Fileformat) search.invoke(myImportOpac, "", "", coc, prefs);
                try { //NOSONAR
                    ds = myRdf.getDigitalDocument().getLogicalDocStruct();
                } catch (Exception e) {
                    log.error(e);
                }
            } catch (NoSuchMethodException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
                return null;
            }

        } else {
            String identifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));
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

            try {
                ds = myRdf.getDigitalDocument().getLogicalDocStruct();
                if (ds.getType().isAnchor()) {
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
        }
        try {
            ats = myImportOpac.getAtstsl();
            if (ds != null) {
                List<? extends Metadata> sort = ds.getAllMetadataByType(prefs.getMetadataTypeByName("CurrentNoSorting"));
                if (sort != null && !sort.isEmpty()) {
                    volumeNumber = sort.get(0).getValue();
                }
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

        if (StringUtils.isBlank(workflowTitle)) {
            workflowTitle = form.getTemplate().getTitel();
            config = null;
            config = getConfig();
        }

        for (Record rec : records) { //NOSONAR
            String timestamp = Long.toString(System.currentTimeMillis());
            ImportObject io = new ImportObject();
            answer.add(io);
            try {

                Object tempObject = rec.getObject();

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
                    if (StringUtils.isNotBlank(getConfig().getAnchorPublicationType())) {
                        anchor = digitalDocument.createDocStruct(prefs.getDocStrctTypeByName(getConfig().getAnchorPublicationType()));
                        logical = digitalDocument.createDocStruct(prefs.getDocStrctTypeByName(getConfig().getPublicationType()));
                        anchor.addChild(logical);
                        digitalDocument.setLogicalDocStruct(anchor);
                    } else {
                        String publicationType = getConfig().getPublicationType();
                        DocStructType logicalType = prefs.getDocStrctTypeByName(publicationType);
                        logical = digitalDocument.createDocStruct(logicalType);
                        digitalDocument.setLogicalDocStruct(logical);
                    }
                } else {
                    try { //NOSONAR
                        boolean validRequest = false;
                        for (MetadataMappingObject mmo : config.getMetadataList()) {
                            if (StringUtils.isNotBlank(mmo.getSearchField()) && headerOrder.get(mmo.getHeaderName()) != null) {
                                validRequest = true;
                                break;
                            }
                        }

                        if (!validRequest) {
                            if (StringUtils.isBlank(config.getIdentifierHeaderName())) {
                                Helper.setFehlerMeldung("Cannot request catalogue, no identifier column defined");
                                log.error("Cannot request catalogue, no identifier column defined");
                                return Collections.emptyList();
                            }

                            Integer columnNumber = headerOrder.get(config.getIdentifierHeaderName());
                            if (columnNumber == null) {
                                Helper.setFehlerMeldung("Cannot request catalogue, identifier column '" + config.getIdentifierHeaderName()
                                + "' not found in excel file.");
                                log.error("Cannot request catalogue, identifier column '" + config.getIdentifierHeaderName()
                                + "' not found in excel file.");
                                return Collections.emptyList();
                            }
                            String catalogueIdentifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));
                            if (StringUtils.isBlank(catalogueIdentifier)) {
                                continue;
                            }
                        }

                        String catalogue = rowMap.get(headerOrder.get(config.getOpacHeader()));
                        if (StringUtils.isBlank(catalogue)) {
                            catalogue = config.getOpacName();
                        }
                        ff = getRecordFromCatalogue(rowMap, headerOrder, catalogue);
                        digitalDocument = ff.getDigitalDocument();
                        logical = digitalDocument.getLogicalDocStruct();
                        if (logical.getType().isAnchor()) {
                            anchor = logical;
                            logical = anchor.getAllChildren().get(0);
                        }

                    } catch (Exception e) {
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
                        try { //NOSONAR
                            // splitting is configured for this field
                            if (config.isSplittingAllowed() && mmo.isSplittingAllowed()) {
                                String delimiter = config.getSplittingDelimiter();
                                String[] values = value.split(delimiter);
                                for (String val : values) {
                                    if (StringUtils.isNotBlank(val)) {
                                        addMetadata(val.trim(), identifier, mmo, logical, anchor);
                                    }
                                }
                            } else if (StringUtils.isNotBlank(config.getListSplitChar()) && value.contains(config.getListSplitChar())) {
                                //multiples ?
                                String[] lstValues = value.split(config.getListSplitChar());
                                for (String strVal : lstValues) {
                                    if (strVal != null && !strVal.isEmpty()) {
                                        addMetadata(strVal, identifier, mmo, logical, anchor);
                                    }
                                }
                            } else {
                                value = addMetadata(value, identifier, mmo, logical, anchor);
                            }
                        } catch (MetadataTypeNotAllowedException e) {
                            log.info(e);
                            // Metadata is not known or not allowed
                        }
                        // create a default title
                        if ("CatalogIDDigital".equalsIgnoreCase(mmo.getRulesetName()) && !"anchor".equals(mmo.getDocType())) {
                            fileName = getImportFolder() + File.separator + value + ".xml";
                            io.setProcessTitle(value);
                            io.setMetsFilename(fileName);
                        }
                    }
                    if (StringUtils.isNotBlank(config.getProcesstitleRule())) {
                        StringBuilder titleValue = new StringBuilder();
                        StringTokenizer tokenizer = new StringTokenizer(config.getProcesstitleRule(), "+");
                        while (tokenizer.hasMoreTokens()) {
                            String myString = tokenizer.nextToken();
                            /*
                             * wenn der String mit ' anfängt und mit ' endet, dann den Inhalt so übernehmen
                             */
                            if (myString.startsWith("'") && myString.endsWith("'")) {
                                titleValue.append(myString.substring(1, myString.length() - 1));
                            } else if ("Signatur".equalsIgnoreCase(myString) || "Shelfmark".equalsIgnoreCase(myString)) {
                                if (StringUtils.isNotBlank(rowMap.get(headerOrder.get(myString)))) {
                                    // replace white spaces with dash, remove other special characters
                                    titleValue.append(
                                            rowMap.get(headerOrder.get(myString)).replace(" ", "-").replace("/", "-").replaceAll("[^\\w-]", ""));
                                }
                            } else if ("timestamp".equalsIgnoreCase(myString)) {
                                titleValue.append(timestamp);
                            } else {
                                String s = rowMap.get(headerOrder.get(myString));
                                titleValue.append(s != null ? s : "");
                            }
                        }
                        String newTitle = titleValue.toString();
                        if (newTitle.endsWith("_")) {
                            newTitle = newTitle.substring(0, newTitle.length() - 1);
                        }
                        // remove non-ascii characters for the sake of TIFF header limits
                        String regex = ConfigurationHelper.getInstance().getProcessTitleReplacementRegex();

                        String filteredTitle = newTitle.replaceAll(regex, config.getReplacement());

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

                if (!getPersonsWithRoles(rec, headerOrder, rowMap, logical, anchor)) {
                    getPersons(rec, headerOrder, rowMap, logical, anchor);
                }

                for (GroupMappingObject gmo : getConfig().getGroupList()) {
                    try { //NOSONAR
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
                                        firstname = "";
                                        lastname = name;
                                    }
                                }
                            } else {
                                firstname = rowMap.get(headerOrder.get(pmo.getFirstnameHeaderName()));
                                lastname = rowMap.get(headerOrder.get(pmo.getLastnameHeaderName()));
                            }
                            Person p = makePerson(pmo.getRulesetName(), firstname, lastname);

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

                        String folderNameRule = ConfigurationHelper.getInstance().getProcessImagesMasterDirectoryName();
                        folderNameRule = folderNameRule.replace("{processtitle}", io.getProcessTitle());

                        Path path = Paths.get(foldername, "images", folderNameRule);
                        try {
                            Files.createDirectories(path.getParent());
                            if (config.isMoveImage()) {
                                StorageProvider.getInstance().move(imageSourceFolder, path);
                            } else {
                                StorageProvider.getInstance().copyDirectory(imageSourceFolder, path);
                            }
                        } catch (IOException e) {
                            log.error(e);
                            if (config.isFailOnMissingImageFiles()) {
                                io.setImportReturnValue(ImportReturnValue.WriteError);
                                io.setErrorMessage(e.getMessage());
                            }
                        }

                    } else if (config.isFailOnMissingImageFiles()) {
                        io.setImportReturnValue(ImportReturnValue.InvalidData);
                        io.setErrorMessage("Missing images in " + imageSourceFolder);
                    } else {
                        log.info("Missing images in " + imageSourceFolder);
                    }
                }

                // check if the process exists
                if (replaceExisting) {
                    boolean dataReplaced = false;
                    Process existingProcess = ProcessManager.getProcessByExactTitle(io.getProcessTitle());
                    if (existingProcess != null) {
                        try {
                            existingProcess.writeMetadataFile(ff);
                            dataReplaced = true;
                        } catch (WriteException | PreferencesException | IOException | SwapException e) {
                            log.error(e);
                        }

                        Path sourceRootFolder = Paths.get(rec.getData());
                        moveImageIntoProcessFolder(existingProcess, sourceRootFolder);
                    }
                    if (dataReplaced) {
                        answer.remove(io);
                    }
                }

            } catch (UGHException e) {
                io.setImportReturnValue(ImportReturnValue.WriteError);
                io.setErrorMessage(e.getMessage());
            }
        }
        // end of all excel rows
        return answer;
    }

    private boolean getPersonsWithRoles(Record rec, Map<String, Integer> headerOrder, Map<Integer, String> rowMap, DocStruct logical,
            DocStruct anchor) {

        log.info("Get persons with roles");

        Boolean boWithRoles = false;
        for (PersonMappingObject pmo : getConfig().getPersonWithRoleList()) {
            String name = "";
            String firstname = "";
            String lastname = "";
            if (pmo.isSplitName()) {
                name = rowMap.get(headerOrder.get(pmo.getHeaderName()));
                if (StringUtils.isNotBlank(name)) {
                    if (name.contains(pmo.getSplitChar())) {
                        if (pmo.isFirstNameIsFirst()) {
                            firstname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
                            lastname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
                        } else {
                            lastname = name.substring(0, name.lastIndexOf(pmo.getSplitChar())).trim();
                            firstname = name.substring(name.lastIndexOf(pmo.getSplitChar()) + 1).trim();
                        }
                    } else {
                        lastname = name;
                    }
                }
            } else {
                firstname = rowMap.get(headerOrder.get(pmo.getFirstnameHeaderName()));
                lastname = rowMap.get(headerOrder.get(pmo.getLastnameHeaderName()));
            }

            log.info(firstname + " " + lastname);
            log.info(pmo.getRulesetName());

            if (StringUtils.isNotBlank(pmo.getRulesetName())) {
                try {
                    //Check if there are multiples:
                    String separator = pmo.getSplitList();
                    String roleSeparator = pmo.getSplitRole();

                    if ((StringUtils.isNotBlank(separator) && name.contains(separator))
                            || (StringUtils.isNotBlank(roleSeparator) && name.contains(roleSeparator))) {

                        boWithRoles = true;
                        String[] lstEntries = name.split(separator);
                        String[] lstNames = new String[lstEntries.length];
                        String[] lstRoles = new String[lstEntries.length];

                        for (int i = 0; i < lstEntries.length; i++) {
                            log.debug(rec.getId() + " Name/role: " + lstEntries[i]);
                            lstNames[i] = lstEntries[i].split(roleSeparator)[0];
                            lstRoles[i] = lstEntries[i].split(roleSeparator)[1];
                        }

                        for (int i = 0; i < lstNames.length; i++) {

                            //if there is a missing splitter, just add until it is missing
                            try {
                                String nameNew = lstNames[i];

                                if (StringUtils.isNotBlank(nameNew)) {
                                    if (nameNew.contains(pmo.getSplitChar())) {
                                        if (pmo.isFirstNameIsFirst()) {
                                            firstname = nameNew.substring(0, nameNew.lastIndexOf(pmo.getSplitChar()));
                                            lastname = nameNew.substring(nameNew.lastIndexOf(pmo.getSplitChar()));
                                        } else {
                                            lastname = nameNew.substring(0, nameNew.lastIndexOf(pmo.getSplitChar())).trim();
                                            firstname = nameNew.substring(nameNew.lastIndexOf(pmo.getSplitChar()) + 1).trim();
                                        }
                                    } else {
                                        firstname = "";
                                        lastname = nameNew;
                                    }
                                }

                                Person pNew = makePerson(pmo.getRulesetName(), firstname, lastname);

                                if (lstRoles.length > i) {
                                    String strRole = getRoleName(lstRoles[i]);
                                    if (strRole != null) {
                                        pNew.setRole(strRole);
                                    }
                                }

                                if (anchor != null && "anchor".equals(pmo.getDocType())) {
                                    anchor.addPerson(pNew);
                                } else {
                                    logical.addPerson(pNew);
                                }
                            } catch (ArrayIndexOutOfBoundsException e) {
                                log.error(rec.getId() + " - Person list is not consistent! " + e.getMessage());
                            }
                        }
                    }
                } catch (Exception e) {
                    log.info(e);
                    // Metadata is not known or not allowed
                }
            }
        }

        return boWithRoles;
    }

    //Translate the roles from ecxel entry to Metadata
    private String getRoleName(String role) {

        for (MetadataMappingObject mmo : getConfig().getRolesList()) {

            if (mmo.getPropertyName().contentEquals(role)) {
                return mmo.getRulesetName();
            }
        }

        return null;
    }

    private void getPersons(Record rec, Map<String, Integer> headerOrder, Map<Integer, String> rowMap, DocStruct logical, DocStruct anchor) {
        List<PersonMappingObject> personList = getConfig().getPersonList();

        //are there roles?
        String[] lstRoles = getRoles(headerOrder, rowMap, personList.size());

        for (int j = 0; j < personList.size(); j++) {
            PersonMappingObject pmo = personList.get(j);

            String strRole = null;
            //are there roles?
            if (lstRoles != null) {
                strRole = lstRoles[j];
            }

            String name = "";
            String firstname = "";
            String lastname = "";
            if (pmo.isSplitName()) {
                name = rowMap.get(headerOrder.get(pmo.getHeaderName()));
                if (StringUtils.isNotBlank(name)) {
                    if (name.contains(pmo.getSplitChar())) {
                        if (pmo.isFirstNameIsFirst()) {
                            firstname = name.substring(0, name.lastIndexOf(pmo.getSplitChar()));
                            lastname = name.substring(name.lastIndexOf(pmo.getSplitChar()));
                        } else {
                            lastname = name.substring(0, name.lastIndexOf(pmo.getSplitChar())).trim();
                            firstname = name.substring(name.lastIndexOf(pmo.getSplitChar()) + 1).trim();
                        }
                    } else {
                        lastname = name;
                    }
                }
            } else {
                firstname = rowMap.get(headerOrder.get(pmo.getFirstnameHeaderName()));
                lastname = rowMap.get(headerOrder.get(pmo.getLastnameHeaderName()));
            }

            String identifier = null;
            if (pmo.getNormdataHeaderName() != null) {
                identifier = rowMap.get(headerOrder.get(pmo.getNormdataHeaderName()));
            }
            if (StringUtils.isNotBlank(pmo.getRulesetName())) {
                try {
                    //Check if there are multiples:
                    String separator = pmo.getSplitList();

                    if (StringUtils.isNotBlank(separator) && name.contains(separator)) {

                        String[] lstNames = name.split(separator);
                        String[] lstIds = new String[lstNames.length];
                        if (pmo.getGndIds() != null && !pmo.getGndIds().isEmpty()) {
                            lstIds = rowMap.get(headerOrder.get(pmo.getGndIds())).split(separator);
                        }

                        //roles for this list:
                        lstRoles = getRoles(headerOrder, rowMap, lstNames.length);

                        for (int i = 0; i < lstNames.length; i++) {

                            //if there is a missing splitter, just add until it is missing
                            try {
                                String nameNew = lstNames[i];

                                if (StringUtils.isNotBlank(nameNew)) {
                                    if (nameNew.contains(pmo.getSplitChar())) {
                                        if (pmo.isFirstNameIsFirst()) {
                                            firstname = nameNew.substring(0, nameNew.lastIndexOf(pmo.getSplitChar()));
                                            lastname = nameNew.substring(nameNew.lastIndexOf(pmo.getSplitChar()));
                                        } else {
                                            lastname = nameNew.substring(0, nameNew.lastIndexOf(pmo.getSplitChar())).trim();
                                            firstname = nameNew.substring(nameNew.lastIndexOf(pmo.getSplitChar()) + 1).trim();
                                        }
                                    } else {
                                        firstname = "";
                                        lastname = nameNew;
                                    }
                                }

                                Person pNew = makePerson(pmo.getRulesetName(), firstname, lastname);

                                if (lstRoles != null) {
                                    String roleName = getRoleName(lstRoles[i]);
                                    if (roleName != null) {
                                        pNew.setRole(roleName);
                                    }
                                }

                                if (StringUtils.isNotEmpty(lstIds[i])) {
                                    pNew.setAutorityFile("gnd", "http://d-nb.info/gnd/", lstIds[i]);
                                }

                                if (anchor != null && "anchor".equals(pmo.getDocType())) {
                                    anchor.addPerson(pNew);
                                } else {
                                    logical.addPerson(pNew);
                                }
                            } catch (ArrayIndexOutOfBoundsException e) {
                                log.error(rec.getId() + " - Person list is not consistent! " + e.getMessage());
                            }
                        }

                    } else {
                        Person p = makePerson(pmo.getRulesetName(), firstname, lastname);

                        if (strRole != null) {
                            String roleName = getRoleName(strRole);
                            if (roleName != null) {
                                p.setRole(roleName);
                            }
                        }
                        if (identifier != null) {
                            p.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);
                        }

                        if (anchor != null && "anchor".equals(pmo.getDocType())) {
                            anchor.addPerson(p);
                        } else {
                            logical.addPerson(p);
                        }
                    }
                } catch (MetadataTypeNotAllowedException e) {
                    log.info(e);
                    // Metadata is not known or not allowed
                }
            }
        }
    }

    private String[] getRoles(Map<String, Integer> headerOrder, Map<Integer, String> rowMap, int length) {
        String[] lstRoles = null;
        if (getConfig().getRoleField() != null) {
            String value = rowMap.get(headerOrder.get(getConfig().getRoleField()));
            //multiples ?
            String strSplitListChar = config.getListSplitChar();
            if (strSplitListChar != null && value.contains(strSplitListChar)) {
                lstRoles = value.split(strSplitListChar);
            } else {
                lstRoles = new String[1];
                lstRoles[0] = value;
            }
            //Same number?
            if (lstRoles.length != length) {
                lstRoles = null;
            }
        }
        return lstRoles;
    }

    private String addMetadata(String value, String identifier, MetadataMappingObject mmo, DocStruct logical, DocStruct anchor)
            throws MetadataTypeNotAllowedException {
        Metadata md = new Metadata(prefs.getMetadataTypeByName(mmo.getRulesetName()));

        value = parseDateIfNecessary(value, md);
        md.setValue(value);
        if (identifier != null) {
            md.setAutorityFile("gnd", "http://d-nb.info/gnd/", identifier);

        }
        if (anchor != null && "anchor".equals(mmo.getDocType())) {
            anchor.addMetadata(md);
        } else {
            logical.addMetadata(md);
        }
        return value;
    }

    //If the metadatum is a date, pasre the value string to look ok
    private String parseDateIfNecessary(String value, Metadata md) {

        String strVal = null;
        if (value.contains("")) {//NOSONAR
            strVal = value.replace("", " ");//NOSONAR
        } else {
            strVal = value;
        }
        String strType = md.getType().getName();
        if ((strType.contentEquals("PublicationYear") || strType.contentEquals("PublicationStart") || strType.contentEquals("PublicationEnd")
                || strType.contentEquals("datedigit") || strType.contentEquals("dateupdate")) && strVal.length() == 8) {
            //check first 4 chars are the year:
            int year = Integer.parseInt(strVal.substring(0, 4));
            if (1600 < year && 2200 > year) {
                if (strVal.substring(4, 6).contentEquals("00")) {
                    strVal = strVal.substring(0, 4);
                } else {
                    strVal = strVal.substring(0, 4) + "/" + strVal.substring(4, 6) + "/" + strVal.substring(6, 8);
                }
            }
        }
        return strVal;
    }

    private Person makePerson(String rulesetName, String firstname, String lastname) throws MetadataTypeNotAllowedException {

        Person p = new Person(prefs.getMetadataTypeByName(rulesetName));

        Matcher m = Pattern.compile("\\((.*?)\\)").matcher(firstname);

        if (m.find()) {
            String strDates = m.group();
            String[] lstDates = strDates.split("-");
            if (lstDates.length == 2) {
                firstname = firstname.replace(strDates, "");
                firstname = firstname.replace("()", "");
            }
        }

        p.setFirstname(firstname);
        p.setLastname(lastname);

        return p;
    }

    private void moveImageIntoProcessFolder(Process existingProcess, Path sourceRootFolder) {
        if (StorageProvider.getInstance().isFileExists(sourceRootFolder)) {
            Path sourceImageFolder = Paths.get(sourceRootFolder.toString(), "images");
            Path sourceOcrFolder = Paths.get(sourceRootFolder.toString(), "ocr");
            if (StorageProvider.getInstance().isDirectory(sourceImageFolder)) {
                List<Path> dataInSourceImageFolder = StorageProvider.getInstance().listFiles(sourceImageFolder.toString());

                for (Path currentData : dataInSourceImageFolder) {
                    if (Files.isDirectory(currentData)) {
                        try {
                            FileUtils.copyDirectory(currentData.toFile(), Paths.get(existingProcess.getImagesDirectory()).toFile());
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    } else {
                        try {
                            FileUtils.copyFile(currentData.toFile(),
                                    Paths.get(existingProcess.getImagesDirectory(), currentData.getFileName().toString()).toFile());
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    }
                }
            }

            // ocr
            if (Files.exists(sourceOcrFolder)) {
                List<Path> dataInSourceImageFolder = StorageProvider.getInstance().listFiles(sourceOcrFolder.toString());
                for (Path currentData : dataInSourceImageFolder) {
                    if (Files.isRegularFile(currentData)) {
                        try {
                            copyFile(currentData, Paths.get(existingProcess.getOcrDirectory(), currentData.getFileName().toString()));
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    } else {
                        try {
                            FileUtils.copyDirectory(currentData.toFile(), Paths.get(existingProcess.getOcrDirectory()).toFile());
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    }
                }
            }
        }
    }

    private void copyFile(Path file, Path destination) throws IOException {
        if (moveFiles) {
            Files.move(file, destination, StandardCopyOption.REPLACE_EXISTING);
        } else {
            Files.copy(file, destination, StandardCopyOption.REPLACE_EXISTING);
        }
    }

    @Override
    public List<Record> splitRecords(String records) {
        return null; //NOSONAR
    }

    @SuppressWarnings("deprecation")
    @Override
    public List<Record> generateRecordsFromFile() {
        if (StringUtils.isBlank(workflowTitle)) {
            workflowTitle = form.getTemplate().getTitel();
        }
        config = null;
        List<Record> recordList = new ArrayList<>();
        String idColumn = getConfig().getIdentifierHeaderName();
        Map<String, Integer> headerOrder = new HashMap<>();

        try (InputStream fileInputStream = new FileInputStream(file); BOMInputStream in = new BOMInputStream(fileInputStream, false);
                Workbook wb = WorkbookFactory.create(in)) {
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
                    Cell cell = row.getCell(cn, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String value = "";
                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            value = cell.getBooleanCellValue() ? "true" : "false";
                            break;
                        case FORMULA:
                            value = cell.getRichStringCellValue().getString();
                            break;
                        case NUMERIC:
                            value = String.valueOf((long) cell.getNumericCellValue());
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
        }
        return recordList;
    }

    @Override
    public List<Record> generateRecordsFromFilenames(List<String> filenames) {
        return null; //NOSONAR
    }

    @Override
    public List<String> splitIds(String ids) {
        return null; //NOSONAR
    }

    @Override
    public List<ImportProperty> getProperties() {
        return null; //NOSONAR
    }

    @Override
    public List<String> getAllFilenames() {
        return null; //NOSONAR
    }

    @Override
    public void deleteFiles(List<String> selectedFilenames) {
        // do nothing
    }

    @Override
    public List<? extends DocstructElement> getCurrentDocStructs() {
        return null; //NOSONAR
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
        return null; //NOSONAR
    }

    @Override
    public DocstructElement getDocstruct() {
        return null;
    }

    @Override
    public void setDocstruct(DocstructElement dse) {
        // do nothing
    }

    @Override
    public PluginType getType() {
        return PluginType.Import;
    }

    @Override
    public boolean isRunnableAsGoobiScript() {
        return config.isRunAsGoobiScript();
    }

    public ExcelConfig getConfig() {
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

    private ExcelConfig loadConfig(String workflowTitle) {
        XMLConfiguration xmlConfig = ConfigPlugins.getPluginConfig(title);
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());
        xmlConfig.setReloadingStrategy(new FileChangedReloadingStrategy());

        SubnodeConfiguration myconfig = null;
        try {
            myconfig = xmlConfig.configurationAt("//config[./template = '" + workflowTitle + "']");
        } catch (IllegalArgumentException e) {
            try {
                myconfig = xmlConfig.configurationAt("//config[./template = '*']");
            } catch (IllegalArgumentException e1) {
                log.error("Excel import plugin: Could not read configuration. At least one <config> block with <template>*</template> is needed.");
            }
        }

        if (myconfig != null) {
            replaceExisting = myconfig.getBoolean("replaceExistingProcesses", false);
            moveFiles = myconfig.getBoolean("moveFiles", false);
        }
        return new ExcelConfig(myconfig);
    }

}
