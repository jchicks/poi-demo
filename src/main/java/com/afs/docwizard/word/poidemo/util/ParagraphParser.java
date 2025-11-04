package com.afs.docwizard.word.poidemo.util;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;


@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ParagraphParser {

  private String text;

  public String getParagraphs() {
    return getTrimmedText()
      .lines()
      .collect(Collectors.joining("\n"));
  }

  public String getFirstParagraph() {
    return getTrimmedText()
      .lines()
      .findFirst()
      .orElse("");
  }

  public List<String> getRemainingParagraphs() {
    return getTrimmedText()
      .lines()
      .skip(1)
      .toList();
  }

  public String getRemainingParagraphsAsString() {
    return getTrimmedText()
      .lines()
      .skip(1)
      .collect(Collectors.joining("\n"));
  }

  public String getTrimmedText() {
    return Optional
      .ofNullable(text)
      .map(String::trim)
      .orElse("");
  }
}
