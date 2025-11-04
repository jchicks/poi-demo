package com.afs.docwizard.word.poidemo.service;

import com.afs.docwizard.word.poidemo.dto.RequirementsInfo;
import com.afs.docwizard.word.poidemo.managers.RequirementDocManager;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.IOException;


@Slf4j
@Service
@RequiredArgsConstructor
public class DocxContentControlService {

  private final ObjectProvider<RequirementDocManager> requirementDocManagerProvider;

  public Resource retrieveTemplate() throws IOException {
    //  Demonstration.  This can be provided from a service
    Resource resource = null;
    RequirementsInfo info = new RequirementsInfo();

    RequirementDocManager requirementDocManager =
      requirementDocManagerProvider.getObject();

    try {
      requirementDocManager.updateMission(info);
//      requirementDocManager.updatePurpose(info);
//      requirementDocManager.updateHistoricalContext(info);

      resource = requirementDocManager.save();
    }
    catch (Exception e) {
      log.error("what the heck?", e);
      throw e;
    }
    finally {
      requirementDocManager.destroy();
    }

    return resource;
  }


}
