package de.intranda.goobi.plugins;

import java.io.File;
import java.util.List;

import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.io.FileUtils;
import org.goobi.beans.Process;
import org.goobi.production.enums.ImportReturnValue;
import org.goobi.production.importer.ImportObject;
import org.goobi.production.importer.Record;
import org.goobi.production.plugin.interfaces.IOpacPlugin;
import org.junit.Assert;
import org.mockito.Mockito;

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

    //    @Test
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
        Assert.assertEquals("9923254553502466", records.get(0).getId());

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

}
