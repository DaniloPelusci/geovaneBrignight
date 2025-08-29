package br.com.portfoliopelusci.controller;


import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import br.com.portfoliopelusci.config.OrganizadorProperties;
import br.com.portfoliopelusci.service.OrganizadorService;

@RestController
@RequestMapping("/organizar")
public class OrganizadorController {

    private final OrganizadorService service;
    private final OrganizadorProperties props;

    public OrganizadorController(OrganizadorService service, OrganizadorProperties props) {
        this.service = service;
        this.props = props;
    }

    /** Dispara usando os valores do application.yml (padrão). */
    @PostMapping
    public String organizar() {
        try {
            service.processar();
            return "Processo concluído (dryRun=" + props.isDryRun() + "). Verifique os logs.";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }

    /** Opcional: permite sobrescrever dryRun rapidamente */
    @PostMapping("/run")
    public String organizarComDry(@RequestParam(defaultValue = "true") boolean dryRun) {
        try {
            props.setDryRun(dryRun);
            service.processar();
            return "Processo concluído (dryRun=" + props.isDryRun() + ").";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }

    /** Cria pastas vazias para cada ordem presente na planilha */
    @PostMapping("/folders")
    public String criarPastas() {
        try {
            service.criarPastas();
            return "Pastas processadas (dryRun=" + props.isDryRun() + ").";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }

    /** Recebe um ZIP contendo as ordens e processa localmente */
    @PostMapping("/upload")
    public String organizarZip(@RequestParam("file") MultipartFile zip) {
        try {
            service.processarZip(zip);
            return "Processo concluído (zip).";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }

    /** Processa todos os ZIPs na pasta configurada */
    @PostMapping("/zip")
    public String organizarZipLocal() {
        try {
            service.processarZips();
            return "Processo concluído (zips locais).";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }
}
