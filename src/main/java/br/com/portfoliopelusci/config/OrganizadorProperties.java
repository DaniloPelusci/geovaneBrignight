package br.com.portfoliopelusci.config;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties(prefix = "organizador")
public class OrganizadorProperties {

    @NotBlank
    private String excelPath;

    @NotBlank
    private String otherExcelPath;

    @NotBlank
    private String outputExcelPath;

    @NotBlank
    private String zipFolderPath;

    @NotBlank
    private String parentZipPath;

    @NotBlank
    private String sourceBasePath;

    @NotBlank
    private String destBasePath;

    @NotBlank
    private String allOrdersBasePath;

    @NotNull
    private Integer sheetIndex = 0;

    // Timezone opcional (default em serviço = America/Sao_Paulo)
    private String timezone = "America/Sao_Paulo";

    private Columns columns = new Columns();

    private boolean dryRun = true;

    private boolean overwriteExisting = true;

    public static class Columns {
        @NotBlank
        private String numero = "Numero";
        @NotBlank
        private String tipo = "Tipo";
        // NOVO: coluna de data (ex.: DUEDATE)
        @NotBlank
        private String data = "DUEDATE";

        public String getNumero() {
            return numero;
        }
        public void setNumero(String numero) {
            this.numero = numero;
        }

        public String getTipo() {
            return tipo;
        }
        public void setTipo(String tipo) {
            this.tipo = tipo;
        }

        public String getData() {
            return data;
        }
        public void setData(String data) {
            this.data = data;
        }
    }

    public String getExcelPath() {
        return excelPath;
    }
    public void setExcelPath(String excelPath) {
        this.excelPath = excelPath;
    }

    public String getOtherExcelPath() {
        return otherExcelPath;
    }
    public void setOtherExcelPath(String otherExcelPath) {
        this.otherExcelPath = otherExcelPath;
    }

    public String getOutputExcelPath() {
        return outputExcelPath;
    }
    public void setOutputExcelPath(String outputExcelPath) {
        this.outputExcelPath = outputExcelPath;
    }

    public String getZipFolderPath() {
        return zipFolderPath;
    }
    public void setZipFolderPath(String zipFolderPath) {
        this.zipFolderPath = zipFolderPath;
    }

    public String getParentZipPath() {
        return parentZipPath;
    }
    public void setParentZipPath(String parentZipPath) {
        this.parentZipPath = parentZipPath;
    }

    public String getSourceBasePath() {
        return sourceBasePath;
    }
    public void setSourceBasePath(String sourceBasePath) {
        this.sourceBasePath = sourceBasePath;
    }

    public String getDestBasePath() {
        return destBasePath;
    }
    public void setDestBasePath(String destBasePath) {
        this.destBasePath = destBasePath;
    }

    public String getAllOrdersBasePath() {
        return allOrdersBasePath;
    }
    public void setAllOrdersBasePath(String allOrdersBasePath) {
        this.allOrdersBasePath = allOrdersBasePath;
    }

    public Integer getSheetIndex() {
        return sheetIndex;
    }
    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getTimezone() {
        return timezone;
    }
    public void setTimezone(String timezone) {
        this.timezone = timezone;
    }

    public Columns getColumns() {
        return columns;
    }
    public void setColumns(Columns columns) {
        this.columns = columns;
    }

    public boolean isDryRun() {
        return dryRun;
    }
    public void setDryRun(boolean dryRun) {
        this.dryRun = dryRun;
    }

    public boolean isOverwriteExisting() {
        return overwriteExisting;
    }
    public void setOverwriteExisting(boolean overwriteExisting) {
        this.overwriteExisting = overwriteExisting;
    }
}
