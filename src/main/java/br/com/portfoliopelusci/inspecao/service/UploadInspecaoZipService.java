package br.com.portfoliopelusci.inspecao.service;

import br.com.portfoliopelusci.inspecao.dto.UploadZipResponse;
import br.com.portfoliopelusci.inspecao.entity.FotoInspecao;
import br.com.portfoliopelusci.inspecao.entity.Inspecao;
import br.com.portfoliopelusci.inspecao.entity.Inspetor;
import br.com.portfoliopelusci.inspecao.repository.InspecaoRepository;
import br.com.portfoliopelusci.inspecao.repository.InspetorRepository;
import jakarta.transaction.Transactional;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

@Service
public class UploadInspecaoZipService {

    private final InspetorRepository inspetorRepository;
    private final InspecaoRepository inspecaoRepository;

    public UploadInspecaoZipService(InspetorRepository inspetorRepository, InspecaoRepository inspecaoRepository) {
        this.inspetorRepository = inspetorRepository;
        this.inspecaoRepository = inspecaoRepository;
    }

    @Transactional
    public UploadZipResponse processar(MultipartFile zipFile) throws IOException {
        validarZip(zipFile);

        String inspetorId = extrairInspetorId(zipFile.getOriginalFilename());
        Inspetor inspetor = inspetorRepository.findById(inspetorId)
                .orElseGet(() -> inspetorRepository.save(new Inspetor(inspetorId)));

        int fotosSalvas = 0;
        Set<String> inspecoesAfetadas = new HashSet<>();

        try (InputStream inputStream = zipFile.getInputStream();
             ZipInputStream zipInputStream = new ZipInputStream(inputStream)) {

            ZipEntry entry;
            while ((entry = zipInputStream.getNextEntry()) != null) {
                if (entry.isDirectory()) {
                    continue;
                }

                String caminho = entry.getName();
                String[] partes = caminho.split("/");
                if (partes.length < 2) {
                    continue;
                }

                String worder = partes[0].trim();
                String nomeArquivo = Path.of(caminho).getFileName().toString();
                if (worder.isBlank() || nomeArquivo.isBlank()) {
                    continue;
                }

                byte[] bytes = zipInputStream.readAllBytes();
                if (bytes.length == 0) {
                    continue;
                }

                Inspecao inspecao = inspecaoRepository.findByInspetorIdAndWorder(inspetorId, worder)
                        .orElseGet(() -> inspecaoRepository.save(new Inspecao(worder, inspetor)));

                String contentType = detectarContentType(nomeArquivo);
                inspecao.adicionarFoto(new FotoInspecao(nomeArquivo, nomeArquivo, contentType, bytes));
                inspecaoRepository.save(inspecao);

                fotosSalvas++;
                inspecoesAfetadas.add(worder);
            }
        }

        return new UploadZipResponse(inspetorId, inspecoesAfetadas.size(), fotosSalvas);
    }

    private static void validarZip(MultipartFile zipFile) {
        if (zipFile == null || zipFile.isEmpty()) {
            throw new IllegalArgumentException("Arquivo ZIP obrigatório.");
        }

        String originalFilename = zipFile.getOriginalFilename();
        if (originalFilename == null || !originalFilename.toLowerCase(Locale.ROOT).endsWith(".zip")) {
            throw new IllegalArgumentException("O arquivo precisa ser um .zip.");
        }
    }

    private static String extrairInspetorId(String nomeArquivoZip) {
        String baseName = nomeArquivoZip.substring(0, nomeArquivoZip.length() - 4).trim();
        if (baseName.isBlank()) {
            throw new IllegalArgumentException("Não foi possível identificar o id do inspetor no nome do ZIP.");
        }
        return baseName;
    }

    private static String detectarContentType(String nomeArquivo) {
        String nome = nomeArquivo.toLowerCase(Locale.ROOT);
        if (nome.endsWith(".jpg") || nome.endsWith(".jpeg")) {
            return MediaType.IMAGE_JPEG_VALUE;
        }
        if (nome.endsWith(".png")) {
            return MediaType.IMAGE_PNG_VALUE;
        }
        if (nome.endsWith(".webp")) {
            return "image/webp";
        }
        return MediaType.APPLICATION_OCTET_STREAM_VALUE;
    }
}
