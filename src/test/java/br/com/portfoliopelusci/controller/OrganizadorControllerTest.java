package br.com.portfoliopelusci.controller;

import br.com.portfoliopelusci.config.OrganizadorProperties;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.web.servlet.MockMvc;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@SpringBootTest
@AutoConfigureMockMvc
class OrganizadorControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @Autowired
    private OrganizadorProperties props;

    @Test
    void atualizaCaminhos() throws Exception {
        Path temp = Files.createTempDirectory("orgPaths");
        String excel = temp.resolve("plan.xlsx").toString();
        String source = temp.resolve("src").toString();
        String dest = temp.resolve("dest").toString();

        mockMvc.perform(post("/organizar/paths")
                        .param("excel", excel)
                        .param("source", source)
                        .param("dest", dest))
                .andExpect(status().isOk());

        assertEquals(excel, props.getExcelPath());
        assertEquals(source, props.getSourceBasePath());
        assertEquals(dest, props.getDestBasePath());
    }
}
