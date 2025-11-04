package com.afs.docwizard.word.poidemo.managers;

import com.afs.docwizard.word.poidemo.dto.RequirementsInfo;
import com.afs.docwizard.word.poidemo.util.ParagraphParser;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.PreDestroy;
import lombok.Getter;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.beans.factory.config.ConfigurableBeanFactory;
import org.springframework.context.annotation.Scope;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;


@Slf4j
@Component
@RequiredArgsConstructor
@Scope(ConfigurableBeanFactory.SCOPE_PROTOTYPE)
public class RequirementDocManager {

  // Namespace decl used by XMLBeans selectPath
  private static final String W_NS_DECL =
    "declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ";

  @Getter
  @Value("classpath:templates/PWS_Editable_Form_Blank_No_Overall_Doc_Cntrl3.docx")
  private Resource templateResource;


  private Map<String, SdtElement> controlMap = null;

  private XWPFDocument doc = null;

  private InputStream inputStream = null;

  /**
   * Initializes the WordprocessingMLPackage instance by loading the template resource.
   */
  @PostConstruct
  public void initialize() {
    try {
      inputStream = templateResource.getInputStream();
      doc = new XWPFDocument(inputStream);

      log.info("Initialized XWPFDocument from template: {}", templateResource.getFilename());

      initializeControls();
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
      if (doc != null) {
        doc.close(); // Close the XWPFDocument
        log.info("Closed XWPFDocument for template: {}", templateResource.getFilename());
      }
    }
    catch (IOException e) {
      log.error("Failed to close XWPFDocument", e);
    }
    finally {
      doc = null; // Dereference the object for GC
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
    var lines = new ParagraphParser(requirementsInfo.getMission())
      .getParagraphs()
      .lines()
      .toList();

    updateSimpleControl("FBI Mission", lines);
  }

  public void updatePurpose(RequirementsInfo requirementsInfo) {
    var lines = new ParagraphParser(requirementsInfo.getPurpose())
      .getParagraphs()
      .lines()
      .toList();

    updateSimpleControl("Purpose", lines);
  }

  public void updateHistoricalContext(RequirementsInfo requirementsInfo) {
    var lines = new ParagraphParser(requirementsInfo.getHistoricalContext())
      .getParagraphs()
      .lines()
      .toList();

    updateSimpleControl("Historical Context", lines);
  }

  public Resource save() throws IOException {
    File tempFile = File.createTempFile("updated-template", ".docx");
    tempFile.deleteOnExit(); // Schedule this file to be deleted when the JVM exits
    wordprocessingMLPackage.save(tempFile);

    log.info("processed controls");

    return new FileSystemResource(tempFile);
  }

  private void updateSimpleControl(String key, List<String> lines) {


  }

  private void initializeControls() {

    // Find ALL SDTs in the document (block/run/cell) with XMLBeans
    XmlObject[] sdts = doc.getDocument().selectPath(
      W_NS_DECL + "$this//w:sdt"
    );

    for (XmlObject xo : sdts) {
      // We only handle block-level SDTs here (your sample is block-level)
      if (xo instanceof CTSdtBlock) {
        CTSdtBlock sdt = (CTSdtBlock) xo;
        if (matchesAliasOrTag(sdt.getSdtPr(), "FBI Mission")) {
          setBlockSdtContentText(sdt, "This is the new me!!!");
        }
      }
      // If you later run into inline/cell SDTs, handle:
      // else if (xo instanceof CTSdtRun) { ... }
      // else if (xo instanceof CTSdtCell) { ... }
    }

  }

  private static boolean matchesAliasOrTag(CTSdtPr pr, String target) {
    if (pr == null) return false;
    String alias = pr.isSetAlias() ? pr.getAlias().getVal() : null;
    String tag   = pr.isSetTag()   ? pr.getTag().getVal()   : null;
    return target.equals(alias) || target.equals(tag);
  }

  private static void setBlockSdtContentText(CTSdtBlock sdt, String text) {
    // Ensure <w:sdtContent> exists
    CTSdtContentBlock content = sdt.isSetSdtContent() ? sdt.getSdtContent()
      : sdt.addNewSdtContent();

    // Reuse first paragraph if present, else create one
    CTP p = content.sizeOfPArray() > 0 ? content.getPArray(0) : CTP.Factory.newInstance();

    // Clear existing runs in that paragraph
    for (int i = p.sizeOfRArray() - 1; i >= 0; i--) {
      p.removeR(i);
    }

    // Add a new run with text
    CTR r = p.addNewR();

    // Optionally apply default run properties from SDT properties (if present)
    if (sdt.isSetSdtPr() && sdt.getSdtPr().isSetRPr()) {
      r.setRPr(sdt.getSdtPr().getRPr());
    }

    CTText t = r.addNewT();
    t.setStringValue(text);
    // If you need to preserve leading/trailing spaces: t.setSpace(STXmlSpace.PRESERVE);

    // Replace content with exactly one paragraph
    content.setPArray(new CTP[]{ p });
  }
}
