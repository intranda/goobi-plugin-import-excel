package de.intranda.goobi.plugins;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.io.input.BOMInputStream;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.enums.ImportType;
import org.goobi.production.enums.PluginType;
import org.goobi.production.importer.DocstructElement;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.PluginLoader;
import org.goobi.production.plugin.interfaces.IImportPluginVersion2;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.goobi.production.properties.ImportProperty;

import de.sub.goobi.config.ConfigPlugins;
import de.sub.goobi.forms.MassImportForm;
import de.sub.goobi.helper.exceptions.ImportPluginException;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import lombok.Data;
import lombok.extern.log4j.Log4j;
import net.xeoh.plugins.base.annotations.PluginImplementation;
import ugh.dl.DocStruct;
import ugh.dl.Fileformat;
import ugh.dl.Metadata;
import ugh.dl.MetadataType;
import ugh.dl.Prefs;
import ugh.exceptions.MetadataTypeNotAllowedException;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.WriteException;
import ugh.fileformats.mets.MetsMods;

@Log4j
@Data
@PluginImplementation
public class OpacExcelImport implements IImportPluginVersion2 {

    private Prefs prefs;
    private MassImportForm form;
    private String importFolder;
    private File file;
    private String data;
    private String currentIdentifier;
    private List<String> currentCollections = new ArrayList<>();
    private String ats;
    private String volumeNumber;

    private String title = "intranda_import_opac_excel";
    // AC\\d+ for AC numbers, \\d+X? for PPN
    private String identifierPattern = "";

    // which opac entry from goobi_opac should be used?
    private String opacName = "ALMA WUW";

    // 12 is the ppn field, 8000 is epn
    private String searchField = "12";

    public OpacExcelImport() {
        XMLConfiguration xmlConfig = null;
        try {
            xmlConfig = ConfigPlugins.getPluginConfig(title);
        } catch (NoSuchMethodError e) {
            xmlConfig = ConfigPlugins.getPluginConfig(this);

        }
        identifierPattern = xmlConfig.getString("identifierPattern", "AC\\d+");
        opacName = xmlConfig.getString("opacName", "ALMA WUW");
        searchField = xmlConfig.getString("searchField", "12");
    }

    @Override
    public void setData(Record r) {
        data = r.getData();
    }

    @Override
    public String getProcessTitle() {
        String answer = "";

        if (StringUtils.isNotBlank(this.ats)) {
            answer = ats.toLowerCase() + "_" + this.currentIdentifier;
        } else {
            answer = this.currentIdentifier;
        }
        if (StringUtils.isNotBlank(volumeNumber)) {
            answer = answer + "_" + volumeNumber;
        }
        return answer;
    }

    @Override
    public List<ImportObject> generateFiles(List<Record> records) {

        List<ImportObject> answer = new ArrayList<>();

        for (Record r : records) {
            form.addProcessToProgressBar();
            currentCollections = r.getCollections();
            this.data = r.getData();
            ImportObject io = new ImportObject();
            Fileformat ff = null;
            try {
                ff = convertData();
            } catch (ImportPluginException e1) {
                if (StringUtils.isNotBlank(e1.getMessage())) {
                    io.setErrorMessage(e1.getMessage());
                }
            }
            io.setProcessTitle(getProcessTitle());
            if (ff != null) {
                r.setId(this.currentIdentifier);
                try {
                    MetsMods mm = new MetsMods(this.prefs);
                    mm.setDigitalDocument(ff.getDigitalDocument());
                    String fileName = getImportFolder() + getProcessTitle() + ".xml";
                    log.debug("Writing '" + fileName + "' into given folder...");
                    mm.write(fileName);
                    io.setMetsFilename(fileName);
                    io.setImportReturnValue(ImportReturnValue.ExportFinished);

                } catch (PreferencesException e) {
                    log.error(currentIdentifier + ": " + e.getMessage(), e);
                    io.setImportReturnValue(ImportReturnValue.InvalidData);
                } catch (WriteException e) {
                    log.error(currentIdentifier + ": " + e.getMessage(), e);
                    io.setImportReturnValue(ImportReturnValue.WriteError);
                }
            } else {
                io.setImportReturnValue(ImportReturnValue.InvalidData);
                if (StringUtils.isBlank(io.getErrorMessage())) {
                    io.setErrorMessage("Could not create metadata record. See log file for additional information.");
                }
            }
            answer.add(io);
        }

        return answer;
    }

    @Override
    public List<Record> splitRecords(String records) {
        return null;
    }

    @Override
    public List<Record> generateRecordsFromFile() {
        List<Record> records = new ArrayList<>();
        InputStream fis = null;
        try {
            fis = new FileInputStream(file);

            BOMInputStream in = new BOMInputStream(fis, false);

            Workbook wb = WorkbookFactory.create(in);

            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell != null) {
                        String value = cell.getStringCellValue();
                        if (StringUtils.isNotBlank(value) && value.trim().matches(identifierPattern)) {
                            Record r = new Record();
                            r.setId(value.trim());
                            r.setData(value.trim());
                            records.add(r);
                        }
                    }
                }

            }
        } catch (InvalidFormatException | IOException e) {
            log.error(e);

        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    log.error(e);
                }
            }
        }
        return records;
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
    public List<ImportType> getImportTypes() {
        List<ImportType> list = new ArrayList<>();
        list.add(ImportType.FILE);
        return list;
    }

    @Override
    public Fileformat convertData() throws ImportPluginException {
        currentIdentifier = data;

        ConfigOpacCatalogue coc = ConfigOpac.getInstance().getCatalogueByName(opacName);
        if (coc == null) {
            throw new ImportPluginException("Catalogue with name " + opacName + " not found. Please check goobi_opac.xml");
        }
        IOpacPlugin myImportOpac = (IOpacPlugin) PluginLoader.getPluginByTitle(PluginType.Opac, coc.getOpacType());
        if (myImportOpac == null) {
            throw new ImportPluginException("Opac plugin " + coc.getOpacType() + " not found. Abort.");
        }
        Fileformat myRdf = null;
        try {
            myRdf = myImportOpac.search(searchField, currentIdentifier, coc, prefs);
            if (myRdf == null) {
                throw new ImportPluginException("Could not import record " + currentIdentifier
                        + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
            }
        } catch (Exception e1) {
            throw new ImportPluginException("Could not import record " + currentIdentifier
                    + ". Usually this means a ruleset mapping is not correct or the record can not be found in the catalogue.");
        }
        DocStruct ds = null;
        DocStruct anchor = null;
        try {
            ds = myRdf.getDigitalDocument().getLogicalDocStruct();
            if (ds.getType().isAnchor()) {
                anchor = ds;
                if (ds.getAllChildren() == null || ds.getAllChildren().isEmpty()) {
                    throw new ImportPluginException("Could not import record " + currentIdentifier
                            + ". Found anchor file, but no children. Try to import the child record.");
                }
                ds = ds.getAllChildren().get(0);
            }
        } catch (PreferencesException e1) {
            throw new ImportPluginException("Could not import record " + currentIdentifier
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

        // add collection
        if (currentCollections != null && !currentCollections.isEmpty()) {
            MetadataType mdt = prefs.getMetadataTypeByName("singleDigCollection");
            for (String col : currentCollections) {
                try {
                    Metadata md = new Metadata(mdt);
                    md.setValue(col);
                    ds.addMetadata(md);
                    if (anchor != null) {
                        Metadata md2 = new Metadata(mdt);
                        md2.setValue(col);
                        anchor.addMetadata(md2);
                    }
                } catch (MetadataTypeNotAllowedException e) {
                    log.error(e);
                }
            }
        }

        return myRdf;

    }

    @Override
    public PluginType getType() {
        return PluginType.Import;
    }

    @Override
    public boolean isRunnableAsGoobiScript() {
        return false;
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
}
