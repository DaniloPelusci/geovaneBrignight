package br.com.portfoliopelusci.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import br.com.portfoliopelusci.config.OrganizadorProperties;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.*;
import java.text.Normalizer;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;
import java.util.stream.Stream;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Serviço responsável por organizar os arquivos de ordem de serviço
 * com base nas informações fornecidas em uma planilha Excel. Também
 * disponibiliza operações auxiliares para criação de pastas e
 * processamento de arquivos ZIP.
 */
@Service
public class OrganizadorService {

    /** Configurações injetadas a partir do arquivo de propriedades. */
    private final OrganizadorProperties props;

    /**
     * Construtor padrão que recebe as propriedades de configuração.
     *
     * @param props configurações utilizadas pelo serviço
     */
    public OrganizadorService(OrganizadorProperties props) {
        this.props = props;
    }

    /**
     * Processa a planilha e copia cada ordem de serviço para
     * a pasta de destino adequada, de acordo com o tipo e a
     * urgência calculada.
     */
    public void processar() throws IOException {
        // Caminho do Excel e diretórios de origem/destino
        Path excel      = Path.of(props.getExcelPath());
        Path sourceBase = Path.of(props.getSourceBasePath());
        Path destBase   = Path.of(props.getDestBasePath());
        int sheetIndex  = props.getSheetIndex();
        boolean dryRun  = props.isDryRun();

        // Valida existência dos caminhos configurados
        if (!Files.exists(excel)) {
            throw new IllegalArgumentException("Planilha não encontrada: " + excel);
        }
        if (!Files.isDirectory(sourceBase)) {
            throw new IllegalArgumentException("Pasta de origem inválida: " + sourceBase);
        }
        // Garante que a pasta de destino exista
        Files.createDirectories(destBase);

        // Lista todas as pastas existentes na origem para detectar
        // quais não foram mencionadas na planilha
        Set<Path> restantes;
        try (Stream<Path> stream = Files.list(sourceBase)) {
            restantes = stream.filter(Files::isDirectory).collect(Collectors.toSet());
        }

        // Define timezone para operações com datas (padrão: São Paulo)
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
            // Abre a aba desejada da planilha
            Sheet sheet = wb.getSheetAt(sheetIndex);
            if (sheet == null) throw new IllegalArgumentException("Aba " + sheetIndex + " não encontrada no Excel.");

            DataFormatter fmt = new DataFormatter();
            Row header = sheet.getRow(sheet.getFirstRowNum());
            if (header == null) throw new IllegalArgumentException("Cabeçalho não encontrado na planilha.");
            // Mapeia os índices das colunas pelo nome do cabeçalho
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

            // Percorre todas as linhas de dados da planilha
            for (int r = first; r <= last; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Lê número, tipo e data de vencimento da linha
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

                // Calcula a urgência (R=atrasado, Y=hoje, B=futuro, N=sem data)
                String urg = computeUrgencia(hoje, due); // R | Y | B | N

                // Pasta origem (original) pelo número
                Path src = sourceBase.resolve(numero);
                if (!Files.exists(src) || !Files.isDirectory(src)) {
                    log("AVISO (linha " + (r+1) + "): Pasta da ordem não encontrada: " + src);
                    continue;
                }
                restantes.remove(src);

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

        // Move as pastas que não estavam na planilha para uma pasta especial
        Path semDocDir = sourceBase.resolve(safeName("não tem no documento"));
        Files.createDirectories(semDocDir);
        for (Path dir : restantes) {
            Path destino = uniquePath(semDocDir.resolve(dir.getFileName()));
            if (dryRun) {
                log("[DRY-RUN] Mover: " + dir.getFileName() + " -> " + semDocDir.getFileName() + "/" + destino.getFileName());
            } else {
                Files.move(dir, destino);
            }
            log("SEM PLANILHA: " + dir.getFileName());
        }

        // Atualiza outra planilha adicionando entradas ausentes
        String otherPath = props.getOtherExcelPath();
        if (otherPath != null && !otherPath.isBlank()) {
            Path other = Path.of(otherPath);
            mergeMissingRows(excel, other, sheetIndex, dryRun);
        }

    }

    /**
     * Cria pastas vazias para cada ordem listada na planilha.
     * Usa o valor da coluna "WORDER" (ou o configurado em application.yml)
     * e cria uma pasta com esse nome dentro de {@code sourceBasePath}.
     */
    public void criarPastas() throws IOException {
        // Caminhos e configuração básica
        Path excel      = Path.of(props.getExcelPath());
        Path baseDir    = Path.of(props.getSourceBasePath());
        int sheetIndex  = props.getSheetIndex();
        boolean dryRun  = props.isDryRun();

        // Garante existência dos arquivos/pastas necessários
        if (!Files.exists(excel)) {
            throw new IllegalArgumentException("Planilha não encontrada: " + excel);
        }
        Files.createDirectories(baseDir);

        try (Workbook wb = new XSSFWorkbook(Files.newInputStream(excel))) {
            Sheet sheet = wb.getSheetAt(sheetIndex);
            if (sheet == null) {
                throw new IllegalArgumentException("Aba " + sheetIndex + " não encontrada no Excel.");
            }

            DataFormatter fmt = new DataFormatter();
            Row header = sheet.getRow(sheet.getFirstRowNum());
            if (header == null) {
                throw new IllegalArgumentException("Cabeçalho não encontrado na planilha.");
            }
            // Obtém índice da coluna com o número da ordem
            Map<String, Integer> map = mapHeader(header, fmt);

            String hNumero = props.getColumns() != null ? props.getColumns().getNumero() : "WORDER";
            int idxNumero = idx(map, hNumero);

            int first = sheet.getFirstRowNum() + 1;
            int last  = sheet.getLastRowNum();
            // Para cada linha cria uma pasta com o número informado
            for (int r = first; r <= last; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String numero = fmt.formatCellValue(row.getCell(idxNumero)).trim();
                if (numero.isBlank()) continue;

                Path dir = baseDir.resolve(safeName(numero));
                if (Files.exists(dir)) {
                    log("EXISTE: " + dir.getFileName());
                } else if (dryRun) {
                    log("[DRY-RUN] Criar pasta: " + dir);
                } else {
                    Files.createDirectories(dir);
                    log("PASTA CRIADA: " + dir.getFileName());
                }
            }
        }
    }

    /**
     * Processa um único arquivo ZIP enviado via upload e executa o
     * método {@link #processar()} sobre o conteúdo extraído.
     */
    public void processarZip(MultipartFile zip) throws IOException {
        // Diretório base onde o conteúdo será extraído
        Path sourceRoot = Path.of(props.getSourceBasePath());
        Files.createDirectories(sourceRoot);

        String filename = zip.getOriginalFilename() != null ? zip.getOriginalFilename() : "upload.zip";
        String baseName = filename.endsWith(".zip") ? filename.substring(0, filename.length() - 4) : filename;
        Path unzipDir = uniquePath(sourceRoot.resolve(baseName));

        // Garante que a pasta exista, mesmo que o ZIP esteja vazio
        Files.createDirectories(unzipDir);
        try (InputStream in = zip.getInputStream()) {
            unzip(in, unzipDir);
        }

        // Se algo deu errado na extração aborta o processamento
        if (!Files.exists(unzipDir)) {
            log("AVISO: Diretório de extração não criado: " + unzipDir);
            return;
        }

        // Loga as pastas encontradas após a extração
        try (Stream<Path> stream = Files.list(unzipDir)) {
            List<Path> entries = stream.collect(Collectors.toList());
            if (entries.isEmpty()) {
                log("AVISO: Arquivo ZIP vazio: " + filename);
                return;
            }
            log("Arquivo ZIP recebido: " + filename);
            entries.stream()
                   .filter(Files::isDirectory)
                   .forEach(p -> log("Ordem encontrada: " + p.getFileName()));
        }

        // Processa o conteúdo extraído reutilizando o método principal
        String originalSource = props.getSourceBasePath();
        try {
            props.setSourceBasePath(unzipDir.toString());
            processar();
        } finally {
            props.setSourceBasePath(originalSource);
        }
    }

    /**
     * Processa todos os arquivos ZIP encontrados na pasta configurada,
     * executando {@link #processZipFile(Path)} para cada um deles.
     */
    public void processarZips() throws IOException {
        String zipFolderPath = props.getZipFolderPath();
        if (zipFolderPath == null || zipFolderPath.isBlank()) {
            throw new IllegalArgumentException("Caminho da pasta de ZIPs não definido.");
        }
        Path zipsDir = Path.of(zipFolderPath);
        if (!Files.exists(zipsDir) || !Files.isDirectory(zipsDir)) {
            throw new IllegalArgumentException("Pasta de ZIPs não encontrada: " + zipsDir);
        }

        // Lista todos os arquivos ZIP presentes na pasta configurada
        List<Path> zipFiles;
        try (Stream<Path> stream = Files.list(zipsDir)) {
            zipFiles = stream.filter(p -> Files.isRegularFile(p) && p.toString().toLowerCase().endsWith(".zip"))
                             .collect(Collectors.toList());
        }

        if (zipFiles.isEmpty()) {
            log("AVISO: Nenhum arquivo ZIP encontrado em: " + zipsDir);
            return;
        }

        // Processa cada arquivo ZIP individualmente
        for (Path zipPath : zipFiles) {
            processZipFile(zipPath);
        }
    }

    /**
     * Processa um arquivo ZIP localizado no sistema de arquivos.
     * O conteúdo é extraído e tratado pelo método {@link #processar()}.
     */
    private void processZipFile(Path zipPath) throws IOException {
        Path sourceRoot = Path.of(props.getSourceBasePath());
        Files.createDirectories(sourceRoot);

        String filename = zipPath.getFileName().toString();
        String baseName = filename.endsWith(".zip") ? filename.substring(0, filename.length() - 4) : filename;
        Path unzipDir = uniquePath(sourceRoot.resolve(baseName));

        // Extrai o ZIP para uma pasta temporária
        Files.createDirectories(unzipDir);
        try (InputStream in = Files.newInputStream(zipPath)) {
            unzip(in, unzipDir);
        }

        // Verifica se a extração ocorreu corretamente
        if (!Files.exists(unzipDir)) {
            log("AVISO: Diretório de extração não criado: " + unzipDir);
            return;
        }

        // Lista as ordens encontradas no ZIP
        try (Stream<Path> stream = Files.list(unzipDir)) {
            List<Path> entries = stream.collect(Collectors.toList());
            if (entries.isEmpty()) {
                log("AVISO: Arquivo ZIP vazio: " + filename);
                return;
            }
            log("Arquivo ZIP processado: " + filename);
            entries.stream()
                   .filter(Files::isDirectory)
                   .forEach(p -> log("Ordem encontrada: " + p.getFileName()));
        }

        // Processa as ordens extraídas reutilizando método principal
        String originalSource = props.getSourceBasePath();
        try {
            props.setSourceBasePath(unzipDir.toString());
            processar();
        } finally {
            props.setSourceBasePath(originalSource);
        }
    }

    /**
     * Extrai todos os arquivos ZIP presentes na pasta informada.
     * Cada ZIP é descompactado para uma subpasta com o mesmo nome
     * (sem a extensão ".zip").
     *
     * @param folderPath caminho da pasta contendo os arquivos ZIP
     */
    public void extrairTodos(String folderPath) throws IOException {
        if (folderPath == null || folderPath.isBlank()) {
            throw new IllegalArgumentException("Caminho da pasta não definido.");
        }

        Path dir = Path.of(folderPath);
        if (!Files.exists(dir) || !Files.isDirectory(dir)) {
            throw new IllegalArgumentException("Pasta de ZIPs não encontrada: " + dir);
        }

        try (Stream<Path> stream = Files.list(dir)) {
            List<Path> zips = stream
                    .filter(p -> Files.isRegularFile(p) && p.toString().toLowerCase().endsWith(".zip"))
                    .collect(Collectors.toList());

            for (Path zipPath : zips) {
                String fileName = zipPath.getFileName().toString();
                String baseName = fileName.endsWith(".zip") ? fileName.substring(0, fileName.length() - 4) : fileName;
                Path targetDir = dir.resolve(baseName);
                Files.createDirectories(targetDir);
                try (InputStream in = Files.newInputStream(zipPath)) {
                    unzip(in, targetDir);
                }
                // Remove eventual pasta duplicada criada pelo ZIP
                fixNestedFolder(targetDir, baseName);
                log("Arquivo ZIP extraído: " + fileName + " para " + targetDir);
            }
        }
    }

    /**
     * Processa o arquivo ZIP pai configurado nas propriedades.
     */
    public void processarZipPai() throws IOException {
        processarZipPai(props.getParentZipPath());
    }

    /**
     * Processa um arquivo ZIP pai que contém diversos arquivos ZIP de
     * inspetores. Cada ZIP interno é extraído para uma pasta separada
     * e suas ordens são copiadas para um diretório consolidado.
     */
    public void processarZipPai(String parentZipPath) throws IOException {
        if (parentZipPath == null || parentZipPath.isBlank()) {
            throw new IllegalArgumentException("Caminho do ZIP pai não definido.");
        }
        Path zipPai = Path.of(parentZipPath);
        if (!Files.exists(zipPai) || !Files.isRegularFile(zipPai)) {
            throw new IllegalArgumentException("ZIP pai não encontrado: " + zipPai);
        }

        // Pastas de destino
        Path destBase = Path.of(props.getDestBasePath());
        Path allOrdersBase = Path.of(props.getAllOrdersBasePath());
        Files.createDirectories(destBase);
        Files.createDirectories(allOrdersBase);

        // Descompacta o ZIP pai em uma pasta temporária
        Path tempDir = Files.createTempDirectory("zip-pai");
        try (InputStream in = Files.newInputStream(zipPai)) {
            unzip(in, tempDir);
        }

        // Para cada ZIP interno encontrado, extrai e copia as ordens
        try (Stream<Path> innerStream = Files.walk(tempDir)) {
            for (Path innerZip : innerStream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.toString().toLowerCase().endsWith(".zip"))
                    .collect(Collectors.toList())) {
                String zipName = innerZip.getFileName().toString();
                String rawBaseName = zipName.endsWith(".zip")
                        ? zipName.substring(0, zipName.length() - 4)
                        : zipName;
                String inspectorPart = rawBaseName;
                int dash = inspectorPart.indexOf('-');
                if (dash >= 0 && dash + 1 < inspectorPart.length()) {
                    inspectorPart = inspectorPart.substring(dash + 1);
                }
                String inspectorFolder = safeName(capitalize(inspectorPart));
                String baseName = safeName(rawBaseName);
                Path inspectorDir = destBase.resolve(inspectorFolder).resolve(baseName);
                if (Files.exists(inspectorDir)) {
                    if (props.isOverwriteExisting()) {
                        if (!props.isDryRun()) deleteRecursively(inspectorDir);
                    } else {
                        log("IGNORADO: pasta do inspetor já existe: " + inspectorDir);
                        continue;
                    }
                }
                if (props.isDryRun()) {
                    log("[DRY-RUN] Descompactar: " + innerZip + " -> " + inspectorDir);
                } else {
                    Files.createDirectories(inspectorDir);
                    try (InputStream in = Files.newInputStream(innerZip)) {
                        unzip(in, inspectorDir);
                    }
                    fixNestedFolder(inspectorDir, baseName);
                    // Extrai eventuais ZIPs de ordens já organizando as pastas
                    extrairTodos(inspectorDir.toString());
                }

                // Organiza as ordens extraídas utilizando a lógica principal
                // e copia o resultado para a pasta consolidada de todas as ordens
                String originalSource = props.getSourceBasePath();
                String originalDest   = props.getDestBasePath();
                try {
                    props.setSourceBasePath(inspectorDir.toString());
                    props.setDestBasePath(allOrdersBase.toString());
                    processar();
                } finally {
                    props.setSourceBasePath(originalSource);
                    props.setDestBasePath(originalDest);
                }
            }
        }

        // Após organizar todas as ordens, renomeia as pastas de origem para
        // refletir o nome final presente em allOrdersBasePath.
        renameLeafFolders(destBase, allOrdersBase);
    }

    /* ===== Helpers ===== */

    /**
     * Renomeia os diretórios de último nível em {@code sourceRoot} para que
     * correspondam aos nomes gerados em {@code allOrdersRoot}. O mapeamento é
     * feito com base no número da ordem, extraído do nome final (antes do
     * primeiro espaço).
     */
    private void renameLeafFolders(Path sourceRoot, Path allOrdersRoot) throws IOException {
        if (!Files.isDirectory(sourceRoot) || !Files.isDirectory(allOrdersRoot)) return;

        Map<String, String> nameMap = new HashMap<>();
        try (Stream<Path> stream = Files.walk(allOrdersRoot)) {
            stream.filter(Files::isDirectory)
                  .filter(this::isLeafSafe)
                  .forEach(p -> {
                      String finalName = p.getFileName().toString();
                      String[] parts = finalName.split(" ", 2);
                      if (parts.length > 0) {
                          nameMap.put(parts[0], finalName);
                      }
                  });
        }

        if (nameMap.isEmpty()) return;

        try (Stream<Path> stream = Files.walk(sourceRoot)) {
            for (Path dir : stream.filter(Files::isDirectory).collect(Collectors.toList())) {
                if (!isLeafSafe(dir)) continue;
                String nome = dir.getFileName().toString();
                String novoNome = nameMap.get(nome);
                if (novoNome != null && !novoNome.equals(nome)) {
                    Path target = uniquePath(dir.getParent().resolve(safeName(novoNome)));
                    if (props.isDryRun()) {
                        log("[DRY-RUN] Renomear: " + dir + " -> " + target);
                    } else {
                        Files.move(dir, target);
                    }
                }
            }
        }
    }

    /** Verifica se o diretório não possui subpastas, ignorando erros. */
    private boolean isLeafSafe(Path dir) {
        try (Stream<Path> stream = Files.list(dir)) {
            return stream.noneMatch(Files::isDirectory);
        } catch (IOException e) {
            return false;
        }
    }

    /**
     * Descompacta um arquivo ZIP a partir de um {@link InputStream}
     * para o diretório informado, criando as pastas conforme necessário.
     */
    private static void unzip(InputStream in, Path target) throws IOException {
        try (ZipInputStream zin = new ZipInputStream(in)) {
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                Path newPath = target.resolve(entry.getName()).normalize();
                if (!newPath.startsWith(target)) {
                    throw new IOException("Entrada inválida: " + entry.getName());
                }
                if (entry.isDirectory()) {
                    Files.createDirectories(newPath);
                } else {
                    Files.createDirectories(newPath.getParent());
                    Files.copy(zin, newPath, StandardCopyOption.REPLACE_EXISTING);
                }
            }
        }
    }

    /**
     * Calcula a urgência de uma ordem com base na data de vencimento.
     *
     * @param hoje data atual
     * @param data data de vencimento
     * @return código de urgência
     */
    private static String computeUrgencia(LocalDate hoje, LocalDate data) {
        if (data == null) return "N";
        if (data.isBefore(hoje)) return "R";
        if (data.isEqual(hoje))  return "Y";
        return "B";
    }

    /**
     * Lê datas do Excel, suportando tanto células de data nativas
     * quanto texto em diferentes formatos (MM/dd/yyyy, ISO etc.).
     */
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

    /**
     * Atualiza a planilha de destino adicionando linhas ausentes
     * e organizando as colunas conforme o mapeamento desejado.
     */
    private void mergeMissingRows(Path source, Path dest, int sheetIndex, boolean dryRun) throws IOException {
        try (Workbook srcWb = new XSSFWorkbook(Files.newInputStream(source));
             Workbook dstWb = Files.exists(dest)
                     ? new XSSFWorkbook(Files.newInputStream(dest))
                     : new XSSFWorkbook()) {

            DataFormatter fmt = new DataFormatter();

            Sheet srcSheet = srcWb.getSheetAt(sheetIndex);
            if (srcSheet == null) {
                throw new IllegalArgumentException("Aba " + sheetIndex + " não encontrada no Excel.");
            }
            Row srcHeader = srcSheet.getRow(srcSheet.getFirstRowNum());
            if (srcHeader == null) {
                throw new IllegalArgumentException("Cabeçalho não encontrado na planilha.");
            }
            Map<String, Integer> srcMap = mapHeader(srcHeader, fmt);

            int idxDate      = idx(srcMap, "DUEDATE");
            int idxInspector = idx(srcMap, "INSPECTOR");
            int idxAddress   = idx(srcMap, "ADDRESS1");
            int idxCity      = idx(srcMap, "CITY");
            int idxZip       = idx(srcMap, "ZIP");
            int idxOtype     = idx(srcMap, "OTYPE");
            int idxWorder    = idx(srcMap, "WORDER");

            Sheet dstSheet;
            if (dstWb.getNumberOfSheets() == 0) {
                dstSheet = dstWb.createSheet(srcSheet.getSheetName());
                Row header = dstSheet.createRow(0);
                String[] headers = {"Data", "Inspector", "Address", "City", "zipcode", "OTYPE", "Worder"};
                for (int i = 0; i < headers.length; i++) {
                    header.createCell(i).setCellValue(headers[i]);
                }
            } else {
                dstSheet = dstWb.getSheetAt(0);
            }

            Row dstHeader = dstSheet.getRow(dstSheet.getFirstRowNum());
            Map<String, Integer> dstMap = mapHeader(dstHeader, fmt);
            int dstIdxWorder = idx(dstMap, "Worder");

            Set<String> existentes = new HashSet<>();
            int tFirst = dstSheet.getFirstRowNum() + 1;
            int tLast = dstSheet.getLastRowNum();
            for (int r = tFirst; r <= tLast; r++) {
                Row row = dstSheet.getRow(r);
                if (row == null) continue;
                String w = fmt.formatCellValue(row.getCell(dstIdxWorder)).trim();
                if (!w.isBlank()) existentes.add(w);
            }

            Map<String, String> inspectorMap = mapInspectorsFromFolders();

            int destRowNum = tLast + 1;
            int sFirst = srcSheet.getFirstRowNum() + 1;
            int sLast  = srcSheet.getLastRowNum();
            for (int r = sFirst; r <= sLast; r++) {
                Row sRow = srcSheet.getRow(r);
                if (sRow == null) continue;
                String worder = fmt.formatCellValue(sRow.getCell(idxWorder)).trim();
                if (worder.isBlank() || existentes.contains(worder)) {
                    continue;
                }
                if (dryRun) {
                    log("[DRY-RUN] Adicionar Worder=" + worder + " ao Excel: " + dest);
                } else {
                    Row dRow = dstSheet.createRow(destRowNum++);
                    dRow.createCell(0).setCellValue(fmt.formatCellValue(sRow.getCell(idxDate)));
                    String inspector = inspectorMap.get(worder);
                    if (inspector == null || inspector.isBlank()) {
                        inspector = fmt.formatCellValue(sRow.getCell(idxInspector));
                    }
                    dRow.createCell(1).setCellValue(inspector);
                    dRow.createCell(2).setCellValue(fmt.formatCellValue(sRow.getCell(idxAddress)));
                    dRow.createCell(3).setCellValue(fmt.formatCellValue(sRow.getCell(idxCity)));
                    dRow.createCell(4).setCellValue(fmt.formatCellValue(sRow.getCell(idxZip)));
                    dRow.createCell(5).setCellValue(fmt.formatCellValue(sRow.getCell(idxOtype)));
                    dRow.createCell(6).setCellValue(worder);
                }
                existentes.add(worder);
            }

            if (dryRun) {
                log("[DRY-RUN] Atualizar planilha: " + dest);
            } else {
                Path parent = dest.getParent();
                if (parent != null) Files.createDirectories(parent);
                try (OutputStream out = Files.newOutputStream(dest)) {
                    dstWb.write(out);
                }
                log("PLANILHA ATUALIZADA: " + dest);
            }
        }
    }

    private Map<String, String> mapInspectorsFromFolders() {
        Map<String, String> map = new HashMap<>();

        Path destBase = Path.of(props.getDestBasePath());
        loadInspectors(destBase, map);

        Path source = Path.of(props.getSourceBasePath());
        if (source.getNameCount() > 1) {
            Path root = source.getParent().getParent();
            loadInspectors(root, map);
        }

        return map;
    }

    private void loadInspectors(Path base, Map<String, String> map) {
        if (base == null || !Files.isDirectory(base)) return;
        try (Stream<Path> stream = Files.walk(base)) {
            stream.filter(Files::isDirectory).forEach(p -> {
                String name = p.getFileName().toString();
                String numero = name.split(" ")[0];
                if (numero.chars().allMatch(Character::isDigit)) {
                    Path rel;
                    try {
                        rel = base.relativize(p);
                    } catch (IllegalArgumentException e) {
                        return;
                    }
                    if (rel.getNameCount() >= 3) {
                        String inspector = rel.getName(0).toString();
                        map.putIfAbsent(numero, inspector);
                    }
                }
            });
        } catch (IOException e) {
            log("AVISO: não foi possível ler pastas de inspetores: " + e.getMessage());
        }
    }

    /**
     * Copia recursivamente uma pasta para outra localização.
     */
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

    /**
     * Remove recursivamente um diretório e todo o seu conteúdo.
     */
    private static void deleteRecursively(Path path) throws IOException {
        if (!Files.exists(path)) return;
        try (var stream = Files.walk(path)) {
            stream.sorted(Comparator.reverseOrder()).forEach(p -> {
                try {
                    Files.delete(p);
                } catch (IOException e) {
                    throw new RuntimeException("Erro ao deletar: " + p + " -> " + e.getMessage(), e);
                }
            });
        }
    }

    /**
     * Remove uma pasta duplicada quando o conteúdo do ZIP já contém
     * um diretório com o mesmo nome do destino.
     */
    private static void fixNestedFolder(Path dir, String baseName) throws IOException {
        Path nested = dir.resolve(baseName);
        if (Files.exists(nested) && Files.isDirectory(nested)) {
            try (Stream<Path> children = Files.list(nested)) {
                for (Path child : children.collect(Collectors.toList())) {
                    Files.move(child, dir.resolve(child.getFileName()), StandardCopyOption.REPLACE_EXISTING);
                }
            }
            deleteRecursively(nested);
        }
    }

    /**
     * Cria um mapa de nome de coluna para índice baseado no cabeçalho do Excel.
     */
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

    /**
     * Obtém o índice de uma coluna pelo nome normalizado.
     */
    private static int idx(Map<String, Integer> colIndex, String headerName) {
        Integer idx = colIndex.get(normalize(headerName));
        if (idx == null) {
            throw new IllegalArgumentException("Coluna '" + headerName + "' não encontrada no cabeçalho (normalizado).");
        }
        return idx;
    }

    /**
     * Normaliza um texto removendo acentos e convertendo para minúsculas.
     */
    private static String normalize(String s) {
        if (s == null) return "";
        String n = Normalizer.normalize(s, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
        return n.toLowerCase(Locale.ROOT).trim();
    }

    /**
     * Coloca a primeira letra em maiúscula e o restante em minúscula.
     */
    private static String capitalize(String s) {
        if (s == null || s.isBlank()) return s;
        return s.substring(0, 1).toUpperCase(Locale.ROOT) + s.substring(1).toLowerCase(Locale.ROOT);
    }

    /**
     * Ajusta um nome para ser seguro para uso em sistemas de arquivos.
     */
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

    /**
     * Gera um caminho único caso o destino já exista, adicionando
     * um sufixo numérico incremental.
     */
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

    /** Escreve uma mensagem simples no console. */
    private static void log(String s) {
        System.out.println(s);
    }
}
