package com.afs.docwizard.word.poidemo.controller;

import com.afs.docwizard.word.poidemo.service.DocxContentControlService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/files")
@RequiredArgsConstructor
public class DemoController {

  private final DocxContentControlService docxContentControlService;

  @Operation(summary = "Get a sample file", description = "Returns a demo file as a resource.")
  @ApiResponses(value = {
    @ApiResponse(responseCode = "200", description = "Successfully retrieved the file"),
    @ApiResponse(responseCode = "404", description = "File not found")
  })
  @GetMapping("/sample-file")
  public ResponseEntity<Resource> getFile() {
    try {
      // Retrieve the updated Word document as a Resource
      var file = docxContentControlService.retrieveTemplate();

      // Set the correct headers for a Microsoft Word file
      return ResponseEntity.ok()
        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=updated-template.docx")
        .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        .body(file);

    } catch (Exception e) {
      // Return 500 in case of exceptions
      return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
    }
  }
}
