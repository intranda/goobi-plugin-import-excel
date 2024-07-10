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
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.StringTokenizer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
import de.intranda.goobi.plugins.util.ImportObjectException;
import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.goobi.plugins.util.PersonMappingObject;
import de.intranda.goobi.plugins.util.VariableReplacer;
import de.intranda.goobi.plugins.util.VolumeGenerator;
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
import lombok.extern.log4j.Log4j2;
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
import ugh.exceptions.TypeNotAllowedAsChildException;
import ugh.exceptions.TypeNotAllowedForParentException;
import ugh.exceptions.UGHException;
import ugh.exceptions.WriteException;
import ugh.fileformats.mets.MetsMods;

@Log4j2
@Data
@PluginImplementation
public class GenericExcelImport implements IImportPluginVersion2, IPlugin {

    private static final String CATALOGIDDIGITAL = "CatalogIDDigital";
    private static final long serialVersionUID = 3965077868027995218L;
    private static final String TITLE = "intranda_import_excel";

    @EqualsAndHashCode.Exclude
    private final transient ConfigOpac configOpac;
    @EqualsAndHashCode.Exclude
    private final transient ConfigurationHelper configHelper;
    @EqualsAndHashCode.Exclude
    private final transient XMLConfiguration xmlConfig;
    @EqualsAndHashCode.Exclude
    private transient ExcelConfig config;

    private Prefs prefs;
    private MassImportForm form;
    private String importFolder;
    private File file;
    private String data;
    private String currentIdentifier;
    private List<String> currentCollections = new ArrayList<>();
    private String ats;
    private String processTitle;

    @EqualsAndHashCode.Exclude
    private boolean replaceExisting = false;
    @EqualsAndHashCode.Exclude
    private boolean moveFiles = false;

    private List<ImportType> importTypes;
    private String workflowTitle;

    public GenericExcelImport() {
        this(ConfigOpac.getInstance(), ConfigPlugins.getPluginConfig(TITLE), ConfigurationHelper.getInstance());
    }

    public GenericExcelImport(ConfigOpac configOpac, XMLConfiguration xmlConfig, ConfigurationHelper configHelper) {
        importTypes = new ArrayList<>();
        importTypes.add(ImportType.FILE);
        this.configOpac = configOpac;
        this.xmlConfig = xmlConfig;
        this.configHelper = configHelper;
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

    /**
     * 
     * @param rowMap
     * @param headerOrder
     * @param catalogue
     * @return
     * @throws ImportPluginException If a general error prevents the import from working. If this is throws, all imports should be aborted
     * @throws ImportObjectException An error with importing the current data, most likely a metadata mapping error or unreachable opac record url
     */
    private Fileformat getRecordFromCatalogue(Map<Integer, String> rowMap, Map<String, Integer> headerOrder, String catalogue)
            throws ImportPluginException, ImportObjectException {
        IOpacPlugin myImportOpac = null;
        ConfigOpacCatalogue coc = null;
        for (ConfigOpacCatalogue configOpacCatalogue : configOpac.getAllCatalogues(workflowTitle)) {
            if (configOpacCatalogue.getTitle().equals(catalogue)) {
                myImportOpac = configOpacCatalogue.getOpacPlugin();
                coc = configOpacCatalogue;
            }
        }
        if (myImportOpac == null) {
            throw new ImportPluginException("Opac plugin for catalogue " + catalogue + " not found. Abort.");
        }
        Fileformat myRdf = null;
        if ("intranda_opac_json".equals(myImportOpac.getTitle())) {

            try {
                Method search = loadSearchMethod(rowMap, headerOrder, myImportOpac);
                myRdf = (Fileformat) search.invoke(myImportOpac, "", "", coc, prefs);

            } catch (NoSuchMethodException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
                throw new ImportPluginException("Cannot perform import: Search method of configured OPAC plugin not found", e);
            }

        } else {
            String identifier = rowMap.get(headerOrder.get(config.getIdentifierHeaderName()));
            try {

                myRdf = myImportOpac.search(config.getSearchField(), identifier, coc, prefs);
                if (myRdf == null) {
                    throw new ImportObjectException("Could not import record " + identifier
                            + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
                }
            } catch (Exception e1) {
                throw new ImportObjectException("Could not import record " + identifier
                        + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
            }

            try {
                DocStruct ds = myRdf.getDigitalDocument().getLogicalDocStruct();
                if (ds.getType().isAnchor()) {

                    Optional<VolumeGenerator> volumeGenerator = this.config.getVolumeGenerator(ds.getType().getName());
                    if (volumeGenerator.isPresent()) {
                        new ArrayList<>(Optional.ofNullable(ds.getAllChildren()).orElse(Collections.emptyList()))
                                .forEach(child -> child.getParent().removeChild(child));
                        generateVolumes(myRdf, ds, volumeGenerator.get());
                    }

                    if (ds.getAllChildren() == null || ds.getAllChildren().isEmpty()) {
                        throw new ImportObjectException(
                                "Could not import record " + identifier + ". Found anchor file, but no children. Try to import the child record.");
                    }
                }
            } catch (PreferencesException | TypeNotAllowedAsChildException e1) {
                throw new ImportObjectException("Could not import record " + identifier
                        + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
            }
        }
        try {
            ats = myImportOpac.getAtstsl();
        } catch (Exception e) {
            ats = "";
        }

        return myRdf;
    }

    public void generateVolumes(Fileformat myRdf, DocStruct ds, VolumeGenerator volumeGenerator)
            throws PreferencesException, TypeNotAllowedAsChildException {
        List<DocStruct> children = volumeGenerator.createVolumes(myRdf.getDigitalDocument(), getPrefs());
        for (DocStruct child : children) {
            ds.addChild(child);
        }
        for (MetadataGroup group : ds.getAllMetadataGroupsByType(prefs.getMetadataGroupTypeByName(volumeGenerator.getMetadataGroupType()))) {
            ds.removeMetadataGroup(group);
        }
    }

    public Method loadSearchMethod(Map<Integer, String> rowMap, Map<String, Integer> headerOrder, IOpacPlugin myImportOpac)
            throws NoSuchMethodException, IllegalAccessException, InvocationTargetException {
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
        return opacClass.getMethod("search", String.class, String.class, ConfigOpacCatalogue.class, Prefs.class);
    }

    @Override
    public List<ImportObject> generateFiles(List<Record> records) {
        List<ImportObject> answer = new ArrayList<>();

        if (StringUtils.isBlank(workflowTitle)) {
            workflowTitle = form.getTemplate().getTitel();
        }
        config = null;
        config = getConfig();

        for (Record rec : records) { //NOSONAR
            // generate a mets file
            try {
                Map<String, Integer> headerOrder = getHeaderOrder(rec);
                Map<Integer, String> rowMap = getRowMap(rec);
                Fileformat ff = generateFileformat(rec);
                createPhysicalStructure(ff);

                DocStruct anchor = getAnchor(ff);
                List<DocStruct> allVolumes = new ArrayList<>(getLogicalDocStructs(ff));
                if (anchor != null) {
                    addCollection(anchor, getConfig().getCollection());
                    allVolumes.forEach(anchor::removeChild);
                }
                for (DocStruct logical : allVolumes) {
                    String timestamp = Long.toString(System.currentTimeMillis());

                    ImportObject io = new ImportObject();
                    io.setImportReturnValue(ImportReturnValue.ExportFinished);

                    if (anchor != null) {
                        anchor.addChild(logical);
                    }
                    VariableReplacer vr = new VariableReplacer(ff.getDigitalDocument(), prefs, null, null);

                    addCollection(logical, getConfig().getCollection());
                    for (MetadataMappingObject mmo : getConfig().getMetadataList()) {
                        addMappedMetadata(io, timestamp, headerOrder, rowMap, logical, anchor, mmo, vr);
                    }
                    if (StringUtils.isBlank(io.getProcessTitle())) {
                        String title = logical.getAllMetadataByType(prefs.getMetadataTypeByName(CATALOGIDDIGITAL))
                                .stream()
                                .findFirst()
                                .map(Metadata::getValue)
                                .orElse("" + timestamp);
                        String fileName = getImportFolder() + File.separator + title + ".xml";
                        io.setProcessTitle(title);
                        io.setMetsFilename(fileName);
                    }

                    if (!getPersonsWithRoles(rec, headerOrder, rowMap, logical, getAnchor(ff))) {
                        getPersons(rec, headerOrder, rowMap, logical, getAnchor(ff));
                    }
                    for (GroupMappingObject gmo : getConfig().getGroupList()) {
                        addMetadataGroup(headerOrder, rowMap, logical, getAnchor(ff), gmo);
                    }

                    if (!ff.write(io.getMetsFilename())) {
                        throw new ImportObjectException("Failed to write mets file " + io.getMetsFilename());
                    }

                    writeFiles(headerOrder, rowMap, io.getMetsFilename(), io.getProcessTitle());

                    // check if the process exists
                    if (replaceExisting) {
                        boolean dataReplaced = replaceExistingProcess(rec, ff, io);
                        if (dataReplaced) {
                            answer.remove(io);
                        }
                    }

                    if (anchor != null) {
                        anchor.removeChild(logical);
                    }
                    answer.add(io);
                }

            } catch (ImportObjectException e) {
                log.error(e);
                ImportObject io = new ImportObject();
                answer.add(io);
                io.setErrorMessage(e.getMessage());
                io.setImportReturnValue(ImportReturnValue.NoData);
            } catch (UGHException e) {
                log.error(e);
                ImportObject io = new ImportObject();
                answer.add(io);
                io.setErrorMessage(e.getMessage());
                io.setImportReturnValue(ImportReturnValue.WriteError);
            } catch (ImportPluginException e) {
                Helper.setFehlerMeldung(e.getMessage());
                log.error(e.getMessage());
                return Collections.emptyList();
            }
        }
        // end of all excel rows
        return answer;
    }

    public void createPhysicalStructure(Fileformat ff)
            throws TypeNotAllowedForParentException, PreferencesException, MetadataTypeNotAllowedException {
        DocStructType physicalType = prefs.getDocStrctTypeByName("BoundBook");
        DocStruct physical = ff.getDigitalDocument().createDocStruct(physicalType);
        ff.getDigitalDocument().setPhysicalDocStruct(physical);
        Metadata imagePath = new Metadata(prefs.getMetadataTypeByName("pathimagefiles"));
        imagePath.setValue("./images/");
        physical.addMetadata(imagePath);
    }

    public boolean replaceExistingProcess(Record rec, Fileformat ff, ImportObject io) {
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
        return dataReplaced;
    }

    public void writeFiles(Map<String, Integer> headerOrder, Map<Integer, String> rowMap, String fileName, String processTitle)
            throws ImportObjectException {
        if (StringUtils.isNotBlank(config.getImageFolderHeaderName())
                && StringUtils.isNotBlank(rowMap.get(headerOrder.get(config.getImageFolderHeaderName())))) {

            Path imageSourceFolder = null;
            if (config.getImageFolderPath() != null) {
                imageSourceFolder = Paths.get(config.getImageFolderPath(), rowMap.get(headerOrder.get(config.getImageFolderHeaderName())));
            } else {
                imageSourceFolder = Paths.get(rowMap.get(headerOrder.get(config.getImageFolderHeaderName())));
            }
            if (StorageProvider.getInstance().isDirectory(imageSourceFolder)) {

                // folder name
                String foldername = fileName.replace(".xml", "");

                String folderNameRule = configHelper.getProcessImagesMasterDirectoryName();
                folderNameRule = folderNameRule.replace("{processtitle}", processTitle);

                Path path = Paths.get(foldername, "images", folderNameRule);
                try {
                    StorageProvider.getInstance().createDirectories(path.getParent());
                    if (config.isMoveImage()) {
                        StorageProvider.getInstance().move(imageSourceFolder, path);
                    } else {
                        StorageProvider.getInstance().copyDirectory(imageSourceFolder, path, false);
                    }
                } catch (IOException e) {
                    log.error(e);
                    if (config.isFailOnMissingImageFiles()) {
                        throw new ImportObjectException(e.getMessage());
                    }
                }

            } else if (config.isFailOnMissingImageFiles()) {
                throw new ImportObjectException("Missing images in " + imageSourceFolder);
            } else {
                log.info("Missing images in " + imageSourceFolder);
            }
        }
    }

    public void addMetadataGroup(Map<String, Integer> headerOrder, Map<Integer, String> rowMap, DocStruct logical, DocStruct anchor,
            GroupMappingObject gmo) throws PreferencesException {
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

    public Fileformat generateFileformat(Record rec) throws ImportPluginException, ImportObjectException {
        Fileformat ff = null;

        if (!config.isUseOpac()) {
            try {
                DigitalDocument digitalDocument = null;
                DocStruct logical = null;
                DocStruct anchor = null;
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
            } catch (UGHException e) {
                throw new ImportObjectException("Cannot import record. Error creating metadata", e);
            }
        } else {
            boolean validRequest = false;
            for (MetadataMappingObject mmo : config.getMetadataList()) {
                if (StringUtils.isNotBlank(mmo.getSearchField()) && getHeaderOrder(rec).get(mmo.getHeaderName()) != null) {
                    validRequest = true;
                    break;
                }
            }

            if (!validRequest) {
                if (StringUtils.isBlank(config.getIdentifierHeaderName())) {
                    throw new ImportPluginException("Cannot request catalogue, no identifier column defined");
                }

                Integer columnNumber = getHeaderOrder(rec).get(config.getIdentifierHeaderName());
                if (columnNumber == null) {
                    throw new ImportPluginException(
                            "Cannot request catalogue, identifier column '" + config.getIdentifierHeaderName() + "' not found in excel file.");
                }
                String catalogueIdentifier = getRowMap(rec).get(getHeaderOrder(rec).get(config.getIdentifierHeaderName()));
                if (StringUtils.isBlank(catalogueIdentifier)) {
                    throw new ImportObjectException("Cannot request catalogue, identifier row empty for record '" + rec.getId() + "'");
                }
            }

            String catalogue = getRowMap(rec).get(getHeaderOrder(rec).get(config.getOpacHeader()));
            if (StringUtils.isBlank(catalogue)) {
                catalogue = config.getOpacName();
            }
            ff = getRecordFromCatalogue(getRowMap(rec), getHeaderOrder(rec), catalogue);
        }
        return ff;
    }

    @SuppressWarnings("unchecked")
    public Map<Integer, String> getRowMap(Record rec) {
        Object tempObject = rec.getObject();
        List<Map<?, ?>> list = (List<Map<?, ?>>) tempObject;
        return (Map<Integer, String>) list.get(1);
    }

    @SuppressWarnings("unchecked")
    public Map<String, Integer> getHeaderOrder(Record rec) {
        Object tempObject = rec.getObject();
        List<Map<?, ?>> list = (List<Map<?, ?>>) tempObject;
        return (Map<String, Integer>) list.get(0);
    }

    public DocStruct getAnchor(Fileformat ff) {
        try {
            if (ff != null && ff.getDigitalDocument() != null && ff.getDigitalDocument().getLogicalDocStruct() != null) {
                DocStruct ds = ff.getDigitalDocument().getLogicalDocStruct();
                if (ds.getType().isAnchor()) {
                    return ds;
                }
            }
        } catch (PreferencesException e) {
            log.error("Error getting digital document from fileformat", e);
        }
        return null;
    }

    public List<DocStruct> getLogicalDocStructs(Fileformat ff) throws PreferencesException {
        if (ff != null && ff.getDigitalDocument() != null && ff.getDigitalDocument().getLogicalDocStruct() != null) {
            DocStruct ds = ff.getDigitalDocument().getLogicalDocStruct();
            if (ds.getType().isAnchor()) {
                return Optional.ofNullable(ds.getAllChildren()).orElse(Collections.emptyList());
            } else {
                return List.of(ds);
            }
        }
        return Collections.emptyList();
    }

    public ImportObject addMappedMetadata(ImportObject io, String timestamp, Map<String, Integer> headerOrder, Map<Integer, String> rowMap,
            DocStruct logical, DocStruct anchor, MetadataMappingObject mmo, VariableReplacer vr) {
        String value = rowMap.get(headerOrder.get(mmo.getHeaderName()));
        String identifier = null;
        String fileName = null;

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
            if (CATALOGIDDIGITAL.equalsIgnoreCase(mmo.getRulesetName()) && !"anchor".equals(mmo.getDocType())) {
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
                        titleValue.append(rowMap.get(headerOrder.get(myString)).replace(" ", "-").replace("/", "-").replaceAll("[^\\w-]", ""));
                    }
                } else if ("timestamp".equalsIgnoreCase(myString)) {
                    titleValue.append(timestamp);
                } else if (myString.startsWith("(") || myString.startsWith("{")) {
                    titleValue.append(vr.replace(myString));
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
            String regex = this.configHelper.getProcessTitleReplacementRegex();
            String filteredTitle = newTitle.replaceAll(regex, config.getReplacement());
            log.info("Generated process title {}", filteredTitle);
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
        return io;
    }

    public void addCollection(DocStruct logical, String col) {
        // add collections if configured
        try {
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
        } catch (MetadataTypeNotAllowedException e) {
            log.error("Cannot add collection to DocStruct:" + e.toString());
        }
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
                    if (StorageProvider.getInstance().isDirectory(currentData)) {
                        try {
                            StorageProvider.getInstance().copyDirectory(currentData, Paths.get(existingProcess.getImagesDirectory()), false);
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    } else {
                        try {
                            StorageProvider.getInstance()
                                    .copyFile(currentData, Paths.get(existingProcess.getImagesDirectory(), currentData.getFileName().toString()));
                        } catch (IOException | SwapException e) {
                            log.error(e);
                        }
                    }
                }
            }

            // ocr
            if (StorageProvider.getInstance().isFileExists(sourceOcrFolder)) {
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
                            StorageProvider.getInstance().copyDirectory(currentData, Paths.get(existingProcess.getOcrDirectory()), false);
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
            StorageProvider.getInstance().move(file, destination);
        } else {
            StorageProvider.getInstance().copyFile(file, destination);
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
                    map.put(cn, getCellValue(row, cn));
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

    public String getCellValue(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
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
        return value;
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

    @Override
    public String getTitle() {
        return TITLE;
    }

}
