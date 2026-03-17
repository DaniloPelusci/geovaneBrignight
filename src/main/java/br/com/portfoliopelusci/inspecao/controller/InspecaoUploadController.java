package br.com.portfoliopelusci.inspecao.controller;

import br.com.portfoliopelusci.inspecao.dto.UploadZipResponse;
import br.com.portfoliopelusci.inspecao.service.UploadInspecaoZipService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/inspecoes")
@Tag(name = "Inspeções", description = "Operações para ingestão de arquivos ZIP de inspeções")
public class InspecaoUploadController {

    private final UploadInspecaoZipService service;

    public InspecaoUploadController(UploadInspecaoZipService service) {
        this.service = service;
    }

    @Operation(
            summary = "Faz upload de um arquivo ZIP de inspeções",
            description = "Recebe um arquivo .zip cujo nome representa o ID do inspetor e processa fotos organizadas por pastas de ordem de serviço."
    )
    @ApiResponses(value = {
            @ApiResponse(responseCode = "201", description = "ZIP processado com sucesso",
                    content = @Content(mediaType = "application/json", schema = @Schema(implementation = UploadZipResponse.class))),
            @ApiResponse(responseCode = "400", description = "Arquivo inválido (não é .zip)", content = @Content),
            @ApiResponse(responseCode = "500", description = "Erro ao ler ou processar o arquivo ZIP", content = @Content)
    })
    @PostMapping(value = "/upload-zip", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<UploadZipResponse> uploadZip(
            @Parameter(description = "Arquivo ZIP para upload", required = true,
                    content = @Content(mediaType = MediaType.APPLICATION_OCTET_STREAM_VALUE,
                            schema = @Schema(type = "string", format = "binary")))
            @RequestParam("file") MultipartFile file) throws IOException {
        UploadZipResponse response = service.processar(file);
        return ResponseEntity.status(HttpStatus.CREATED).body(response);
    }
}
