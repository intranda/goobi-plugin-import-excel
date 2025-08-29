package de.intranda.goobi.plugins.util;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import lombok.Data;
import lombok.extern.log4j.Log4j;
import ugh.dl.Corporate;
import ugh.dl.DigitalDocument;
import ugh.dl.DocStruct;
import ugh.dl.Metadata;
import ugh.dl.MetadataGroup;
import ugh.dl.MetadataGroupType;
import ugh.dl.MetadataType;
import ugh.dl.Person;
import ugh.dl.Prefs;
import ugh.exceptions.TypeNotAllowedForParentException;
import ugh.exceptions.UGHException;

@Data
@Log4j
public class VolumeGenerator {

    private final String volumeType;
    private final String metadataGroupType;

    public List<DocStruct> createVolumes(DigitalDocument digDoc, Prefs prefs) {

        DocStruct anchor = digDoc.getLogicalDocStruct();
        return Optional.ofNullable(anchor.getAllMetadataGroupsByType(prefs.getMetadataGroupTypeByName(this.metadataGroupType)))
                .orElse(Collections.emptyList())
                .stream()
                .flatMap(group -> {
                    List<Metadata> ids = Optional.ofNullable(group.getMetadataByType("CatalogIDDigital")).orElse(Collections.emptyList());
                    List<DocStruct> volumes = new ArrayList<>();
                    for (int i = 0; i < ids.size(); i++) {
                        try {
                            DocStruct volume = digDoc.createDocStruct(prefs.getDocStrctTypeByName(volumeType));
                            transferMetadata(group, volume, i);
                            volumes.add(volume);
                        } catch (TypeNotAllowedForParentException e) {
                            log.error(String.format("Error adding DocStruct %s to %s: %s", volumeType, anchor.getType().getName(), e.toString()));
                        }
                    }
                    return volumes.stream();
                })
                .collect(Collectors.toList());

    }

    public void transferMetadata(MetadataGroup group, DocStruct volume, int valueIndex) {

        List<String> mdTypes =
                group.getMetadataList().stream().map(Metadata::getType).map(MetadataType::getName).distinct().collect(Collectors.toList());
        for (String metadataType : mdTypes) {
            try {
                Metadata md = group.getMetadataByType(metadataType).stream().skip(valueIndex).findFirst().orElse(null);
                if (md != null) {
                    volume.addMetadata(md);
                }
            } catch (UGHException e) {
                log.error(String.format("Error adding metadata %s to %s: %s", metadataType, volume.getType().getName(), e.toString()));
            }
        }

        List<String> groupTypes = group.getAllMetadataGroups()
                .stream()
                .map(MetadataGroup::getType)
                .map(MetadataGroupType::getName)
                .distinct()
                .collect(Collectors.toList());
        for (String groupType : groupTypes) {
            try {
                MetadataGroup subGroup = group.getAllMetadataGroupsByName(groupType).stream().skip(valueIndex).findFirst().orElse(null);
                if (subGroup != null) {
                    volume.addMetadataGroup(subGroup);
                }
            } catch (UGHException e) {
                log.error(String.format("Error adding group %s to %s: %s", groupType, volume.getType().getName(), e.toString()));
            }
        }

        List<String> corporateTypes =
                group.getCorporateList().stream().map(Metadata::getType).map(MetadataType::getName).distinct().collect(Collectors.toList());
        for (String corporateType : corporateTypes) {
            try {
                Corporate corporate = group.getCorporateByType(corporateType).stream().skip(valueIndex).findFirst().orElse(null);
                if (corporate != null) {
                    volume.addCorporate(corporate);
                }
            } catch (UGHException e) {
                log.error(String.format("Error adding corporate %s to %s: %s", corporateType, volume.getType().getName(), e.toString()));
            }
        }

        List<String> personTypes =
                group.getPersonList().stream().map(Metadata::getType).map(MetadataType::getName).distinct().collect(Collectors.toList());
        for (String personType : personTypes) {
            try {
                Person person = group.getPersonByType(personType).stream().skip(valueIndex).findFirst().orElse(null);
                if (person != null) {
                    volume.addPerson(person);
                }
            } catch (UGHException e) {
                log.error(String.format("Error adding person %s to %s: %s", personType, volume.getType().getName(), e.toString()));
            }
        }
    }

}
