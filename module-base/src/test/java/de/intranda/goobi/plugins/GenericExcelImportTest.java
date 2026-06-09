package de.intranda.goobi.plugins;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.io.FileUtils;
import org.goobi.beans.Process;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.junit.Assert;
import org.junit.Test;
import org.mockito.Mockito;

import de.intranda.goobi.plugins.util.MetadataMappingObject;
import de.intranda.ugh.extension.MarcFileformat;
import de.sub.goobi.config.ConfigurationHelper;
import de.sub.goobi.forms.MassImportForm;
import de.unigoettingen.sub.search.opac.ConfigOpac;
import de.unigoettingen.sub.search.opac.ConfigOpacCatalogue;
import ugh.dl.Fileformat;
import ugh.dl.Prefs;
import ugh.exceptions.PreferencesException;
import ugh.exceptions.ReadException;
import ugh.fileformats.mets.MetsMods;

public class GenericExcelImportTest {

    @Test
    public void test() throws Exception {
        File importFile = new File("src/test/resources/9923254553502466.xlsx");
        File almaRecordFile = new File("src/test/resources/9923254553502466.alma.xml");
        File importFolder = new File("src/test/resources/output");
        if (importFolder.exists()) {
            FileUtils.deleteDirectory(importFolder);
        }
        importFolder.mkdir();
        XMLConfiguration xmlConfig = new XMLConfiguration(new File("src/test/resources/plugin_intranda_import_excel.xml"));
        Prefs prefs = new Prefs();
        prefs.loadPrefs("src/test/resources/edinburgh.xml");

        ConfigOpac configOpac = Mockito.mock(ConfigOpac.class);
        MassImportForm form = Mockito.mock(MassImportForm.class);
        Process template = Mockito.mock(Process.class);
        ConfigOpacCatalogue cat = Mockito.mock(ConfigOpacCatalogue.class);
        IOpacPlugin opacPlugin = Mockito.mock(IOpacPlugin.class);
        ConfigurationHelper configHelper = Mockito.mock(ConfigurationHelper.class);

        Mockito.when(configHelper.getProcessImagesMasterDirectoryName()).thenReturn("{processtitle}_master");
        Mockito.when(configHelper.getProcessTitleReplacementRegex()).thenReturn("\\W+");
        Mockito.when(opacPlugin.search("12", "9923254553502466", cat, prefs)).thenReturn(loadFileformatFromMarc(almaRecordFile, prefs));
        Mockito.when(cat.getOpacPlugin()).thenReturn(opacPlugin);
        Mockito.when(cat.getTitle()).thenReturn("MockOpac");
        Mockito.when(configOpac.getAllCatalogues(Mockito.anyString())).thenReturn(List.of(cat));
        Mockito.when(form.getTemplate()).thenReturn(template);
        Mockito.when(template.getTitel()).thenReturn("");

        GenericExcelImport excelImport = new GenericExcelImport(configOpac, xmlConfig, configHelper);
        excelImport.setForm(form);
        excelImport.setFile(importFile);
        excelImport.setPrefs(prefs);
        excelImport.setImportFolder(importFolder.getAbsolutePath());
        List<Record> records = excelImport.generateRecordsFromFile();

        Assert.assertEquals(1, records.size());
        Assert.assertEquals("9923254553502466_2", records.get(0).getId());

        List<ImportObject> importFiles = excelImport.generateFiles(records);
        Assert.assertEquals(45, importFiles.size());
        Assert.assertEquals(ImportReturnValue.ExportFinished, importFiles.get(0).getImportReturnValue());
    }

    private Fileformat loadFileformatFromMarc(File almaRecordFile, Prefs prefs) throws ReadException, PreferencesException {
        Fileformat marc = new MarcFileformat(prefs);
        marc.read(almaRecordFile.toString());
        Fileformat mets = new MetsMods(prefs);
        mets.setDigitalDocument(marc.getDigitalDocument());
        return mets;
    }

    private GenericExcelImport buildPlugin() throws Exception {
        XMLConfiguration xmlConfig = new XMLConfiguration(new File("src/test/resources/plugin_intranda_import_excel.xml"));
        return new GenericExcelImport(
                Mockito.mock(ConfigOpac.class),
                xmlConfig,
                Mockito.mock(ConfigurationHelper.class));
    }

    private Record buildRecord(Map<String, Integer> headerOrder, Map<Integer, String> rowMap, int rowNumber) {
        Record r = new Record();
        r.setData(String.valueOf(rowNumber));
        List<Map<?, ?>> list = new ArrayList<>();
        list.add(headerOrder);
        list.add(rowMap);
        r.setObject(list);
        return r;
    }

    @Test
    public void testValidateExcelData_missingColumn() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setHeaderName("Titel");
        mmo.setRequired(true);

        Map<String, Integer> headerOrder = new HashMap<>();

        List<String> errors = plugin.validateExcelData(List.of(mmo), headerOrder, new ArrayList<>());
        assertEquals(1, errors.size());
        assertTrue(errors.get(0).contains("Titel"));
        assertTrue(errors.get(0).contains("does not exist"));
    }

    @Test
    public void testValidateExcelData_requiredFieldEmpty() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setHeaderName("Titel");
        mmo.setRequired(true);

        Map<String, Integer> headerOrder = new HashMap<>();
        headerOrder.put("Titel", 0);

        Map<Integer, String> rowMap = new HashMap<>();
        rowMap.put(0, "");

        Record record = buildRecord(headerOrder, rowMap, 2);

        List<String> errors = plugin.validateExcelData(List.of(mmo), headerOrder, List.of(record));
        assertEquals(1, errors.size());
        assertTrue(errors.get(0).contains("Row 2"));
        assertTrue(errors.get(0).contains("Titel"));
        assertTrue(errors.get(0).contains("empty"));
    }

    @Test
    public void testValidateExcelData_patternMismatch() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setHeaderName("Signatur");
        mmo.setPattern("\\d+");

        Map<String, Integer> headerOrder = new HashMap<>();
        headerOrder.put("Signatur", 0);

        Map<Integer, String> rowMap = new HashMap<>();
        rowMap.put(0, "abc");

        Record record = buildRecord(headerOrder, rowMap, 3);

        List<String> errors = plugin.validateExcelData(List.of(mmo), headerOrder, List.of(record));

        assertEquals(1, errors.size());
        assertTrue(errors.get(0).contains("Row 3"));
        assertTrue(errors.get(0).contains("Signatur"));
        assertTrue(errors.get(0).contains("The value 'abc' does not match the expected format"));
    }

    @Test
    public void testValidateExcelData_invalidContent() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setHeaderName("Sprache");
        mmo.setValidContent(List.of("de", "en", "fr"));

        Map<String, Integer> headerOrder = new HashMap<>();
        headerOrder.put("Sprache", 0);

        Map<Integer, String> rowMap = new HashMap<>();
        rowMap.put(0, "it");

        Record record = buildRecord(headerOrder, rowMap, 4);

        List<String> errors = plugin.validateExcelData(List.of(mmo), headerOrder, List.of(record));
        assertEquals(1, errors.size());
        assertTrue(errors.get(0).contains("Row 4"));
        assertTrue(errors.get(0).contains("Sprache"));
        assertTrue(errors.get(0).contains("it"));
    }

    @Test
    public void testValidateExcelData_noErrorsForValidData() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo1 = new MetadataMappingObject();
        mmo1.setHeaderName("Titel");
        mmo1.setRequired(true);

        MetadataMappingObject mmo2 = new MetadataMappingObject();
        mmo2.setHeaderName("Sprache");
        mmo2.setValidContent(List.of("de", "en", "fr"));

        Map<String, Integer> headerOrder = new HashMap<>();
        headerOrder.put("Titel", 0);
        headerOrder.put("Sprache", 1);

        Map<Integer, String> rowMap = new HashMap<>();
        rowMap.put(0, "Ein Titel");
        rowMap.put(1, "de");

        Record record = buildRecord(headerOrder, rowMap, 2);

        List<String> errors = plugin.validateExcelData(List.of(mmo1, mmo2), headerOrder, List.of(record));

        assertTrue(errors.isEmpty());
    }

    @Test
    public void testValidateExcelData_emptyValueSkipsPatternAndContent() throws Exception {
        GenericExcelImport plugin = buildPlugin();

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setHeaderName("Sprache");
        mmo.setPattern("\\d+");
        mmo.setValidContent(List.of("de", "en"));
        // required=false (default)

        Map<String, Integer> headerOrder = new HashMap<>();
        headerOrder.put("Sprache", 0);

        Map<Integer, String> rowMap = new HashMap<>();
        rowMap.put(0, ""); // empty, but not required

        Record record = buildRecord(headerOrder, rowMap, 2);

        List<String> errors = plugin.validateExcelData(List.of(mmo), headerOrder, List.of(record));

        assertTrue(errors.isEmpty());
    }

}
