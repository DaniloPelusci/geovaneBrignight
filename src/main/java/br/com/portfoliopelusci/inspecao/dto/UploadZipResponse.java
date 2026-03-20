package br.com.portfoliopelusci.inspecao.dto;

import io.swagger.v3.oas.annotations.media.Schema;

@Schema(
        name = "UploadZipResponse",
        description = "Resumo do processamento do ZIP com fotos de inspeção."
)
public record UploadZipResponse(String inspetorId, int inspecoesAfetadas, int fotosSalvas) {

    public UploadZipResponse(
            @Schema(
                    description = "Identificador do inspetor extraído do nome do arquivo ZIP.",
                    example = "INSP-001"
            ) String inspetorId,
            @Schema(
                    description = "Quantidade de inspeções (worders) criadas ou atualizadas.",
                    example = "3"
            ) int inspecoesAfetadas,
            @Schema(
                    description = "Quantidade total de fotos salvas no processamento.",
                    example = "18"
            ) int fotosSalvas
    ) {
        this.inspetorId = inspetorId;
        this.inspecoesAfetadas = inspecoesAfetadas;
        this.fotosSalvas = fotosSalvas;
    }
}
