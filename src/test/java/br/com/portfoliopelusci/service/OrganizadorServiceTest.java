package br.com.portfoliopelusci.service;

import br.com.portfoliopelusci.config.OrganizadorProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

class OrganizadorServiceTest {

    @Test
    void processarZipPaiExtraiECopiaOrdens() throws IOException {
        Path temp = Files.createTempDirectory("org");
        Path dest = temp.resolve("dest");
        Path allOrders = temp.resolve("todas");
        Files.createDirectories(dest);
        Files.createDirectories(allOrders);

        Path inspectorZip = createInspectorZip(temp, "0828-Geovane", "350394452", "dados");
        Path parentZip = temp.resolve("pai.zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(parentZip))) {
            addZipEntry(out, inspectorZip);
        }

        Path excel = createExcel(temp.resolve("plan.xlsx"), "350394452", "A");

        OrganizadorProperties props = new OrganizadorProperties();
        props.setExcelPath(excel.toString());
        props.setZipFolderPath(temp.toString());
        props.setParentZipPath(parentZip.toString());
        props.setSourceBasePath(temp.resolve("src").toString());
        props.setDestBasePath(dest.toString());
        props.setAllOrdersBasePath(allOrders.toString());
        props.setDryRun(false);
        props.setOverwriteExisting(true);

        OrganizadorService service = new OrganizadorService(props);
        service.processarZipPai();

        Path ordemDest = dest.resolve("Geovane").resolve("0828-Geovane").resolve("350394452 A N").resolve("data.txt");
        Path ordemAll = allOrders.resolve("A").resolve("350394452 A N").resolve("data.txt");
        assertTrue(Files.exists(ordemDest));
        assertTrue(Files.exists(ordemAll));
        assertEquals("dados", Files.readString(ordemAll, StandardCharsets.UTF_8));
    }

    @Test
    void processarZipPaiNaoSobrescreveQuandoOverwriteFalse() throws IOException {
        Path temp = Files.createTempDirectory("org2");
        Path dest = temp.resolve("dest");
        Path allOrders = temp.resolve("todas");
        Files.createDirectories(dest);
        Files.createDirectories(allOrders);

        Path inspectorZip1 = createInspectorZip(temp, "0828-Geovane", "350394452", "A");
        Path parentZip1 = temp.resolve("pai1.zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(parentZip1))) {
            addZipEntry(out, inspectorZip1);
        }

        Path excel = createExcel(temp.resolve("plan.xlsx"), "350394452", "A");

        OrganizadorProperties props = new OrganizadorProperties();
        props.setExcelPath(excel.toString());
        props.setZipFolderPath(temp.toString());
        props.setParentZipPath(parentZip1.toString());
        props.setSourceBasePath(temp.resolve("src").toString());
        props.setDestBasePath(dest.toString());
        props.setAllOrdersBasePath(allOrders.toString());
        props.setDryRun(false);
        props.setOverwriteExisting(true);

        OrganizadorService service = new OrganizadorService(props);
        service.processarZipPai();

        Path inspectorZip2 = createInspectorZip(temp, "0828-Geovane", "350394452", "B");
        Path parentZip2 = temp.resolve("pai2.zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(parentZip2))) {
            addZipEntry(out, inspectorZip2);
        }

        props.setParentZipPath(parentZip2.toString());
        props.setOverwriteExisting(false);
        service.processarZipPai();

        Path ordemAll = allOrders.resolve("A").resolve("350394452 A N").resolve("data.txt");
        assertEquals("A", Files.readString(ordemAll, StandardCharsets.UTF_8));
    }

    @Test
    void processarZipPaiExtraiZipsDeOrdens() throws IOException {
        Path temp = Files.createTempDirectory("org3");
        Path dest = temp.resolve("dest");
        Path allOrders = temp.resolve("todas");
        Files.createDirectories(dest);
        Files.createDirectories(allOrders);

        Path orderZip = createOrderZip(temp, "350394452", "dados");
        Path inspectorZip = createInspectorZipComFilho(temp, "0828-Geovane", orderZip);
        Path parentZip = temp.resolve("pai3.zip");
        try (ZipOutputStream out = new ZipOutputStream(Files.newOutputStream(parentZip))) {
            addZipEntry(out, inspectorZip);
        }

        Path excel = createExcel(temp.resolve("plan.xlsx"), "350394452", "A");

        OrganizadorProperties props = new OrganizadorProperties();
        props.setExcelPath(excel.toString());
        props.setParentZipPath(parentZip.toString());
        props.setSourceBasePath(temp.resolve("src").toString());
        props.setDestBasePath(dest.toString());
        props.setAllOrdersBasePath(allOrders.toString());
        props.setDryRun(false);
        props.setOverwriteExisting(true);

        OrganizadorService service = new OrganizadorService(props);
        service.processarZipPai();

        Path ordemDest = dest.resolve("Geovane").resolve("0828-Geovane").resolve("350394452 A N").resolve("data.txt");
        Path ordemAll = allOrders.resolve("A").resolve("350394452 A N").resolve("data.txt");
        assertTrue(Files.exists(ordemDest));
        assertTrue(Files.exists(ordemAll));
        assertEquals("dados", Files.readString(ordemAll, StandardCharsets.UTF_8));
    }

    private static Path createExcel(Path file, String numero, String tipo) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet();
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("WORDER");
            header.createCell(1).setCellValue("OTYPE");
            Row row = sheet.createRow(1);
            row.createCell(0).setCellValue(numero);
            row.createCell(1).setCellValue(tipo);
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
        return file;
    }

    private static Path createInspectorZip(Path dir, String inspector, String ordem, String content) throws IOException {
        Path zipPath = dir.resolve(inspector + ".zip");
        try (ZipOutputStream zout = new ZipOutputStream(Files.newOutputStream(zipPath))) {
            zout.putNextEntry(new ZipEntry(ordem + "/"));
            zout.closeEntry();
            zout.putNextEntry(new ZipEntry(ordem + "/data.txt"));
            zout.write(content.getBytes(StandardCharsets.UTF_8));
            zout.closeEntry();
        }
        return zipPath;
    }

    private static Path createOrderZip(Path dir, String ordem, String content) throws IOException {
        Path zipPath = dir.resolve(ordem + ".zip");
        try (ZipOutputStream zout = new ZipOutputStream(Files.newOutputStream(zipPath))) {
            zout.putNextEntry(new ZipEntry(ordem + "/"));
            zout.closeEntry();
            zout.putNextEntry(new ZipEntry(ordem + "/data.txt"));
            zout.write(content.getBytes(StandardCharsets.UTF_8));
            zout.closeEntry();
        }
        return zipPath;
    }

    private static Path createInspectorZipComFilho(Path dir, String inspector, Path childZip) throws IOException {
        Path zipPath = dir.resolve(inspector + ".zip");
        try (ZipOutputStream zout = new ZipOutputStream(Files.newOutputStream(zipPath))) {
            zout.putNextEntry(new ZipEntry(childZip.getFileName().toString()));
            Files.copy(childZip, zout);
            zout.closeEntry();
        }
        return zipPath;
    }

    private static void addZipEntry(ZipOutputStream parent, Path file) throws IOException {
        parent.putNextEntry(new ZipEntry(file.getFileName().toString()));
        Files.copy(file, parent);
        parent.closeEntry();
    }
}
