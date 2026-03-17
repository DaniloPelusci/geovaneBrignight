package br.com.portfoliopelusci.inspecao.controller;

import br.com.portfoliopelusci.inspecao.dto.UploadZipResponse;
import br.com.portfoliopelusci.inspecao.service.UploadInspecaoZipService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/inspecoes")
public class InspecaoUploadController {

    private final UploadInspecaoZipService service;

    public InspecaoUploadController(UploadInspecaoZipService service) {
        this.service = service;
    }

    @PostMapping("/upload-zip")
    public ResponseEntity<UploadZipResponse> uploadZip(@RequestParam("file") MultipartFile file) throws IOException {
        UploadZipResponse response = service.processar(file);
        return ResponseEntity.status(HttpStatus.CREATED).body(response);
    }
}
