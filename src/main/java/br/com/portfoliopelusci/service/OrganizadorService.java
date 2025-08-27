package br.com.portfoliopelusci.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import br.com.portfoliopelusci.config.OrganizadorProperties;

import java.io.IOException;
import java.nio.file.*;
import java.text.Normalizer;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;

@Service
public class OrganizadorService {

    private final OrganizadorProperties props;

    public OrganizadorService(OrganizadorProperties props) {
        this.props = props;
    }

    public void processar() throws IOException {
        Path excel      = Path.of(props.getExcelPath());
        Path sourceBase = Path.of(props.getSourceBasePath());
        Path destBase   = Path.of(props.getDestBasePath());
        int sheetIndex  = props.getSheetIndex();
        boolean dryRun  = props.isDryRun();

        if (!Files.exists(excel)) {
            throw new IllegalArgumentException("Planilha não encontrada: " + excel);
        }
        if (!Files.isDirectory(sourceBase)) {
            throw new IllegalArgumentException("Pasta de origem inválida: " + sourceBase);
        }
        Files.createDirectories(destBase);

        // Timezone (default SP)
        String tz = "America/Sao_Paulo";
        try {
            String configuredTz = (props.getTimezone() != null && !props.getTimezone().isBlank())
                    ? props.getTimezone() : "America/Sao_Paulo";
            ZoneId.of(configuredTz); // valida
            tz = configuredTz;
        } catch (Exception ignored) {}
        ZoneId zone = ZoneId.of(tz);
        LocalDate hoje = LocalDate.now(zone);

        try (Workbook wb = new XSSFWorkbook(Files.newInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(sheetIndex);
            if (sheet == null) throw new IllegalArgumentException("Aba " + sheetIndex + " não encontrada no Excel.");

            DataFormatter fmt = new DataFormatter();
            Row header = sheet.getRow(sheet.getFirstRowNum());
            if (header == null) throw new IllegalArgumentException("Cabeçalho não encontrado na planilha.");
            Map<String, Integer> map = mapHeader(header, fmt);

            String hNumero = props.getColumns() != null ? props.getColumns().getNumero() : "WORDER";
            String hTipo   = props.getColumns() != null ? props.getColumns().getTipo()   : "OTYPE";
            String hData   = (props.getColumns() != null && props.getColumns().getData() != null && !props.getColumns().getData().isBlank())
                    ? props.getColumns().getData() : "DUEDATE";

            int idxNumero = idx(map, hNumero);
            int idxTipo   = idx(map, hTipo);
            int idxData   = idx(map, hData);

            int first = sheet.getFirstRowNum() + 1;
            int last  = sheet.getLastRowNum();

            for (int r = first; r <= last; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String numero = fmt.formatCellValue(row.getCell(idxNumero)).trim();
                String tipo   = fmt.formatCellValue(row.getCell(idxTipo)).trim();
                LocalDate due = readLocalDate(row.getCell(idxData), fmt, zone);

                if (numero.isBlank()) {
                    log("AVISO (linha " + (r+1) + "): Numero vazio. Ignorando.");
                    continue;
                }
                if (tipo.isBlank()) {
                    log("AVISO (linha " + (r+1) + "): Tipo vazio (Numero=" + numero + "). Ignorando.");
                    continue;
                }
                if (due == null) {
                    log("AVISO (linha " + (r+1) + "): Data inválida (Numero=" + numero + "). Usando URGENCIA=SEM_DATA.");
                }

                String urg = computeUrgencia(hoje, due); // ATRASADA | URGENTE | NORMAL | SEM_DATA

                // Pasta origem (original) pelo número
                Path src = sourceBase.resolve(numero);
                if (!Files.exists(src) || !Files.isDirectory(src)) {
                    log("AVISO (linha " + (r+1) + "): Pasta da ordem não encontrada: " + src);
                    continue;
                }

                // Pasta de destino por tipo
                Path tipoDir = destBase.resolve(safeName(tipo));

                // Nome final: NUMERO TIPO URGENCIA
                String finalName = safeName((numero + " " + tipo + " " + urg).trim());

                Path dest = uniquePath(tipoDir.resolve(finalName));

                if (dryRun) {
                    log("[DRY-RUN] Copiar: " + src + " -> " + dest + " (DUEDATE=" + (due != null ? due : "-") + ", urg=" + urg + ")");
                } else {
                    Files.createDirectories(tipoDir);
                    copyDirectory(src, dest);
                    log("COPIADO: " + src.getFileName() + " -> " + tipoDir.getFileName() + "/" + dest.getFileName() + " (urg=" + urg + ")");
                }
            }
        }
    }

    /* ===== Helpers ===== */

    private static String computeUrgencia(LocalDate hoje, LocalDate data) {
        if (data == null) return "SEM_DATA";
        if (data.isBefore(hoje)) return "ATRASADA";
        if (data.isEqual(hoje))  return "URGENTE";
        return "NORMAL";
    }

    // Lê datas do Excel suportando célula de data nativa e texto (MM/dd/yyyy, ISO)
    private static LocalDate readLocalDate(Cell cell, DataFormatter fmt, ZoneId zone) {
        if (cell == null) return null;

        // Caso seja uma célula de data do Excel
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue().toInstant().atZone(zone).toLocalDate();
        }

        // Texto (ex.: "06/08/2025", "6/8/25", "06/08/2025 00:00", "2025-06-08")
        String raw = fmt.formatCellValue(cell);
        if (raw == null || raw.isBlank()) return null;
        raw = raw.trim();

        // Se vier com hora, corta a parte da hora
        int space = raw.indexOf(' ');
        if (space > 0) raw = raw.substring(0, space).trim();

        // Suporte a: MM/dd/uuuu, M/d/uuuu, MM/dd/uu, M/d/uu e ISO yyyy-MM-dd
        List<DateTimeFormatter> patterns = List.of(
            DateTimeFormatter.ofPattern("MM/dd/uuuu"),
            DateTimeFormatter.ofPattern("M/d/uuuu"),
            DateTimeFormatter.ofPattern("MM/dd/uu"),
            DateTimeFormatter.ofPattern("M/d/uu"),
            DateTimeFormatter.ISO_LOCAL_DATE
        );

        for (DateTimeFormatter f : patterns) {
            try { return LocalDate.parse(raw, f); }
            catch (DateTimeParseException ignored) {}
        }
        return null;
    }

    private static void copyDirectory(Path source, Path target) throws IOException {
        try (var stream = Files.walk(source)) {
            stream.forEach(path -> {
                try {
                    Path relative = source.relativize(path);
                    Path destino = target.resolve(relative);
                    if (Files.isDirectory(path)) {
                        Files.createDirectories(destino);
                    } else {
                        Files.createDirectories(destino.getParent());
                        Files.copy(path, destino, StandardCopyOption.REPLACE_EXISTING);
                    }
                } catch (IOException e) {
                    throw new RuntimeException("Erro ao copiar: " + path + " -> " + e.getMessage(), e);
                }
            });
        }
    }

    private static Map<String, Integer> mapHeader(Row header, DataFormatter fmt) {
        Map<String, Integer> map = new HashMap<>();
        for (int c = header.getFirstCellNum(); c < header.getLastCellNum(); c++) {
            String raw = fmt.formatCellValue(header.getCell(c));
            if (raw != null && !raw.isBlank()) {
                map.put(normalize(raw), c);
            }
        }
        return map;
    }

    private static int idx(Map<String, Integer> colIndex, String headerName) {
        Integer idx = colIndex.get(normalize(headerName));
        if (idx == null) {
            throw new IllegalArgumentException("Coluna '" + headerName + "' não encontrada no cabeçalho (normalizado).");
        }
        return idx;
    }

    private static String normalize(String s) {
        if (s == null) return "";
        String n = Normalizer.normalize(s, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
        return n.toLowerCase(Locale.ROOT).trim();
    }

    private static String safeName(String s) {
        String n = s;
        n = n.replaceAll("\\p{Cntrl}", " ");
        n = n.replaceAll("[\\\\/:*?\"<>|]", "_");
        n = n.trim().replaceAll("\\s+", " ");
        Set<String> reserved = Set.of("CON","PRN","AUX","NUL","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9");
        if (reserved.contains(n.toUpperCase(Locale.ROOT))) n = "_" + n + "_";
        if (n.length() > 120) n = n.substring(0, 120);
        return n;
    }

    private static Path uniquePath(Path dest) {
        if (!Files.exists(dest)) return dest;
        String base = dest.getFileName().toString();
        Path parent = dest.getParent();
        int i = 1;
        while (true) {
            Path candidate = parent.resolve(base + "-" + i);
            if (!Files.exists(candidate)) return candidate;
            i++;
        }
    }

    private static void log(String s) {
        System.out.println(s);
    }
}
