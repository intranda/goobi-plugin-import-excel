package de.intranda.goobi.plugins.util;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;

import org.apache.commons.configuration.ConfigurationException;
import org.apache.commons.configuration.XMLConfiguration;
import org.apache.commons.configuration.tree.xpath.XPathExpressionEngine;
import org.junit.Test;

public class ExcelConfigTest {

    @Test
    public void testReadVolumeGeneratorConfigs() throws ConfigurationException {
        File configFile = new File("src/test/resources/plugin_intranda_import_excel.xml");
        assertTrue(configFile.isFile());
        XMLConfiguration xmlConfig = new XMLConfiguration(configFile);
        xmlConfig.setExpressionEngine(new XPathExpressionEngine());
        ExcelConfig config = new ExcelConfig(xmlConfig.configurationAt("/config[./template='*']"));
        assertEquals("PeriodicalVolume", config.getVolumeGenerator("Periodical").map(VolumeGenerator::getVolumeType).orElse(null));
        assertEquals("VolumeData", config.getVolumeGenerator("Periodical").map(VolumeGenerator::getMetadataGroupType).orElse(null));
        assertEquals("Volume", config.getVolumeGenerator("MultiVolumeWork").map(VolumeGenerator::getVolumeType).orElse(null));
    }

}
