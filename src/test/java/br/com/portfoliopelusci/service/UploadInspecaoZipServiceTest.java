package br.com.portfoliopelusci.service;

import br.com.portfoliopelusci.inspecao.dto.UploadZipResponse;
import br.com.portfoliopelusci.inspecao.entity.Inspecao;
import br.com.portfoliopelusci.inspecao.entity.Inspetor;
import br.com.portfoliopelusci.inspecao.repository.InspecaoRepository;
import br.com.portfoliopelusci.inspecao.repository.InspetorRepository;
import br.com.portfoliopelusci.inspecao.service.UploadInspecaoZipService;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class UploadInspecaoZipServiceTest {

    @Mock
    private InspetorRepository inspetorRepository;

    @Mock
    private InspecaoRepository inspecaoRepository;

    @InjectMocks
    private UploadInspecaoZipService service;

    @Test
    void deveProcessarZipESalvarRelacoes() throws IOException {
        MockMultipartFile file = new MockMultipartFile(
                "file",
                "12345.zip",
                "application/zip",
                zipComArquivos(
                        "1001/foto1.jpg", "jpeg-data".getBytes(),
                        "1001/foto2.png", "png-data".getBytes(),
                        "1002/foto3.jpg", "jpeg2-data".getBytes())
        );

        when(inspetorRepository.findById("12345")).thenReturn(Optional.empty());
        when(inspetorRepository.save(any(Inspetor.class))).thenAnswer(invocation -> invocation.getArgument(0));
        when(inspecaoRepository.findByInspetorIdAndWorder(anyString(), anyString())).thenReturn(Optional.empty());
        when(inspecaoRepository.save(any(Inspecao.class))).thenAnswer(invocation -> invocation.getArgument(0));

        UploadZipResponse response = service.processar(file);

        assertEquals("12345", response.inspetorId());
        assertEquals(2, response.inspecoesAfetadas());
        assertEquals(3, response.fotosSalvas());

        verify(inspetorRepository).save(any(Inspetor.class));
        verify(inspecaoRepository, atLeast(3)).save(any(Inspecao.class));
    }

    @Test
    void deveFalharQuandoNaoForZip() {
        MockMultipartFile file = new MockMultipartFile("file", "semzip.txt", "text/plain", "oi".getBytes());

        IllegalArgumentException exception = assertThrows(IllegalArgumentException.class, () -> service.processar(file));

        assertEquals("O arquivo precisa ser um .zip.", exception.getMessage());
        verifyNoInteractions(inspetorRepository, inspecaoRepository);
    }

    private static byte[] zipComArquivos(Object... dados) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try (ZipOutputStream zipOutputStream = new ZipOutputStream(baos)) {
            for (int i = 0; i < dados.length; i += 2) {
                String nome = (String) dados[i];
                byte[] conteudo = (byte[]) dados[i + 1];

                zipOutputStream.putNextEntry(new ZipEntry(nome));
                zipOutputStream.write(conteudo);
                zipOutputStream.closeEntry();
            }
        }
        return baos.toByteArray();
    }
}
