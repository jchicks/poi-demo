package com.afs.docwizard.word.poidemo.util;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.ArrayList;
import java.util.List;

/**
 * Utility class for manipulating and interacting with content controls in WordprocessingML (Word) documents.
 * Provides methods to perform operations involving structured document tags (SDTs) within Word documents.
 */
public class ControlContentUtil {

  // WordprocessingML namespace URI
  private static final String WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  /**
   * Fetch XML objects by alias from an XWPFDocument.
   *
   * @param document the XWPFDocument to search in.
   * @param alias the value of the alias tag to look for.
   * @return a list of matching XmlObjects.
   * @throws RuntimeException if XPath query execution or any other error occurs.
   */
  public static List<XmlObject> getXmlObjectsByAlias(XWPFDocument document, String alias) {
    try {
      // Build XPath query with explicit namespace declaration
      String xpathQuery = """
        declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $this//w:sdt[w:sdtPr/w:alias[@w:val='%s'] or w:sdtPr/w:tag[@w:val='%s']]
        """.formatted(escape(alias), escape(alias)).trim();

      // Execute the query
      XmlObject[] results = document.getDocument().selectPath(xpathQuery);

      // Convert the results to a list
      List<XmlObject> resultList = new ArrayList<>();
      for (XmlObject result : results) {
        resultList.add(result);
      }
      return resultList;
    } catch (Exception e) {
      throw new RuntimeException("Error executing XPath query with alias: " + alias, e);
    }
  }


  /**
   * Builds an XPath query string to locate control content elements in an XML structure by alias or tag value.
   * The query searches for elements with either an alias or tag attribute that matches the specified target value.
   *
   * @param target the alias or tag value to use in the query.
   * @return the constructed XPath query string.
   */
  public static String buildControlContentAliasQuery(String target) {
    return """
      $this//w:sdt[ w:sdtPr/w:alias[@w:val="%s"] or w:sdtPr/w:tag[@w:val="%s"] ]
      """
      .formatted(escape(target), escape(target))
      .trim();
  }

  /**
   * Escapes double quotes in a string for safe usage in XPath queries.
   *
   * @param s the input string to escape.
   * @return the string with double quotes escaped.
   */
  private static String escapeDoubleQuotes(String s) {
    // Minimal escape for double quotes in XPath string literal
    return s.replace("\"", "\\\"");
  }
  /**
   * Escapes a string to safely use it in an XPath query.
   *
   * @param s the input string to escape.
   * @return the escaped string.
   */
  private static String escape(String s) {
    // Minimal escaping for single quotes in XPath string literals
    return s.replace("'", "&apos;");
  }

  public static void setBlockSdtContentText(CTSdtBlock sdt, String text) {
    CTSdtContentBlock content = sdt.isSetSdtContent() ? sdt.getSdtContent() : sdt.addNewSdtContent();

    // Ensure exactly one paragraph
    CTP p = (content.sizeOfPArray() > 0) ? content.getPArray(0) : content.addNewP();

    // Clear existing runs
    for (int i = p.sizeOfRArray() - 1; i >= 0; i--)
      p.removeR(i);

    // Add new run + text
    CTR r = p.addNewR();
    if (sdt.isSetSdtPr() && sdt.getSdtPr().isSetRPr()) {
      r.setRPr(sdt.getSdtPr().getRPr()); // keep any run defaults defined on the SDT
    }
    CTText t = r.addNewT();
    t.setStringValue(text);
    // If you need to preserve leading/trailing spaces: t.setSpace(STXmlSpace.PRESERVE);

    // Keep only this paragraph in the content
    content.setPArray(new CTP[]{ p });
  }

}
