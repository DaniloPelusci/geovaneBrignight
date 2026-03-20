package br.com.portfoliopelusci.inspecao.controller;

import br.com.portfoliopelusci.inspecao.dto.UploadZipResponse;
import br.com.portfoliopelusci.inspecao.service.UploadInspecaoZipService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping({"/inspecoes", "/foto-inspections"})
public class InspecaoUploadController {

    private final UploadInspecaoZipService service;

    public InspecaoUploadController(UploadInspecaoZipService service) {
        this.service = service;
    }

    @Operation(
            summary = "Importa ZIP de fotos de inspeção",
            description = "Recebe um arquivo ZIP, processa as fotos por inspeção e retorna um resumo da importação."
    )
    @ApiResponses(value = {
            @ApiResponse(
                    responseCode = "201",
                    description = "ZIP processado com sucesso.",
                    content = @Content(
                            mediaType = "application/json",
                            schema = @Schema(implementation = UploadZipResponse.class)
                    )
            ),
            @ApiResponse(responseCode = "400", description = "Arquivo inválido para processamento."),
            @ApiResponse(responseCode = "500", description = "Erro inesperado ao processar ZIP.")
    })
    @PostMapping({"/upload-zip", "/import-zip"})
    public ResponseEntity<UploadZipResponse> uploadZip(@RequestParam("file") MultipartFile file) throws IOException {
        UploadZipResponse response = service.processar(file);
        return ResponseEntity.status(HttpStatus.CREATED).body(response);
    }
}
