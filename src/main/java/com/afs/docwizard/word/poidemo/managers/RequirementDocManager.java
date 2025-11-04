package com.afs.docwizard.word.poidemo.managers;

import com.afs.docwizard.word.poidemo.dto.RequirementsInfo;
import com.afs.docwizard.word.poidemo.util.ParagraphParser;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.PreDestroy;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.beans.factory.config.ConfigurableBeanFactory;
import org.springframework.context.annotation.Scope;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import static com.afs.docwizard.word.poidemo.util.ControlContentUtil.getXmlObjectsByAlias;
import static com.afs.docwizard.word.poidemo.util.ControlContentUtil.setBlockSdtContentText;


@Slf4j
@Component
@RequiredArgsConstructor
@Scope(ConfigurableBeanFactory.SCOPE_PROTOTYPE)
public class RequirementDocManager {

  @Getter
  @Value("classpath:templates/PWS_Editable_Form_Blank_No_Overall_Doc_Cntrl3.docx")
  private Resource templateResource;

  private XWPFDocument document = null;

  private InputStream inputStream = null;

  /**
   * Initializes the WordprocessingMLPackage instance by loading the template resource.
   */
  @PostConstruct
  public void initialize() {
    try {
      inputStream = templateResource.getInputStream();
      document = new XWPFDocument(inputStream);

      log.info("Initialized XWPFDocument from template: {}", templateResource.getFilename());
    }
    catch (IOException e) {
      log.error("Failed to initialize WordprocessingMLPackage from template: {}", templateResource, e);
      throw new IllegalStateException("Could not load DOCX template", e);
    }
  }

  /**
   * Cleans up the WordprocessingMLPackage and prevents potential memory leaks.
   */
  @PreDestroy
  public void destroy() {
    try {
      if (document != null) {
        document.close(); // Close the XWPFDocument
        log.info("Closed XWPFDocument for template: {}", templateResource.getFilename());
      }
    }
    catch (IOException e) {
      log.error("Failed to close XWPFDocument", e);
    }
    finally {
      document = null; // Dereference the object for GC
    }

    try {
      if (inputStream != null) {
        inputStream.close(); // Close the input stream
        log.info("Closed InputStream for template: {}", templateResource.getFilename());
      }
    }
    catch (IOException e) {
      log.error("Failed to close InputStream", e);
    }
    finally {
      inputStream = null; // Dereference the object for GC
    }

    log.info("Cleaned up resources for template: {}", templateResource.getFilename());
  }

  public void updateMission(RequirementsInfo requirementsInfo) {
    updateSimpleControl("FBI Mission", requirementsInfo.getMission());
  }

  public void updatePurpose(RequirementsInfo requirementsInfo) {
    updateSimpleControl("Purpose", requirementsInfo.getPurpose());
  }

  public void updateHistoricalContext(RequirementsInfo requirementsInfo) {
    updateSimpleControl("Historical Context", requirementsInfo.getHistoricalContext());
  }

  public Resource save() throws IOException {
    var tempFile = File.createTempFile("updated-template", ".docx");

    tempFile.deleteOnExit(); // Schedule this file to be deleted when the JVM exits

    var fos = new FileOutputStream(tempFile);

    document.write(fos);

    return new FileSystemResource(tempFile);
  }

  private void updateSimpleControl(String key, String unbrokenLines) {
    var lines = new ParagraphParser(unbrokenLines)
      .getParagraphs()
      .lines()
      .toList();

    var hits = getXmlObjectsByAlias(document, key);

    hits
      .stream()
      .filter(CTSdtBlock.class::isInstance)
      .map(CTSdtBlock.class::cast)
      .forEach(sdt -> setBlockSdtContentText(sdt, lines));
  }
}
