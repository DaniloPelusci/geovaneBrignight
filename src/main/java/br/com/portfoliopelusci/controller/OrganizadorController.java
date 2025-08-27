package br.com.portfoliopelusci.controller;


import org.springframework.web.bind.annotation.*;

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
}
