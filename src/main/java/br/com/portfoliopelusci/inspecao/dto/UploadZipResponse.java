package br.com.portfoliopelusci.inspecao.dto;

import io.swagger.v3.oas.annotations.media.Schema;

@Schema(description = "Resumo do processamento do arquivo ZIP enviado")
public record UploadZipResponse(
        @Schema(description = "Identificador do inspetor derivado do nome do arquivo ZIP", example = "12345")
        String inspetorId,
        @Schema(description = "Quantidade de inspeções afetadas pelo processamento", example = "2")
        int inspecoesAfetadas,
        @Schema(description = "Quantidade total de fotos salvas", example = "3")
        int fotosSalvas) {
}
