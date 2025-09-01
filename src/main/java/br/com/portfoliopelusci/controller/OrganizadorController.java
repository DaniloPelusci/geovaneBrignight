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

    /**
     * Inicia o processo de organização utilizando as configurações padrão
     * definidas no {@code application.yml}. Este método lê a planilha
     * configurada, identifica as ordens existentes na pasta de origem e copia
     * cada uma delas para a pasta de destino correspondente, respeitando o
     * modo de execução definido em {@link OrganizadorProperties#isDryRun()}.
     *
     * @return mensagem indicando o resultado da operação e se foi apenas um
     *         "dry run" (simulação) ou uma execução real
     */
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

    /**
     * Variante do endpoint principal que permite sobrescrever rapidamente o
     * valor de {@code dryRun}. Ao enviar {@code dryRun=false}, o processo passa
     * a copiar/mover os arquivos de fato; caso contrário, apenas registra nos
     * logs o que seria feito.
     *
     * @param dryRun indica se a execução deve ser apenas simulada
     * @return mensagem com o status do processamento
     */
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

    /**
     * Varre a planilha configurada e cria, dentro da pasta de origem,
     * diretórios vazios para cada número de ordem encontrado. Útil para
     * preparar a estrutura de diretórios antes do recebimento dos arquivos.
     * O comportamento de criação pode ser apenas simulado quando
     * {@code dryRun} estiver ativado.
     *
     * @return mensagem indicando o resultado da criação das pastas
     */
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

    /**
     * Recebe via upload um arquivo ZIP contendo múltiplas ordens. O conteúdo
     * é extraído para a pasta de origem e, em seguida, o mesmo fluxo de
     * organização do endpoint principal é executado sobre os arquivos
     * extraídos.
     *
     * @param zip arquivo compactado com as pastas das ordens
     * @return mensagem relatando a conclusão do processamento do ZIP
     */
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

    /**
     * Percorre uma pasta previamente configurada em busca de arquivos ZIP e
     * processa cada um deles sequencialmente. Para cada ZIP encontrado é
     * realizada a extração, seguida da execução do processo de organização.
     *
     * @return mensagem informando a finalização do processamento de todos os
     *         ZIPs locais
     */
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

    /**
     * Processa um arquivo ZIP "pai" que contém, em seu interior, outros
     * arquivos ZIP (geralmente separados por inspetor). Cada ZIP interno é
     * extraído para uma pasta nomeada com o inspetor e o nome do ZIP, e uma
     * cópia das ordens é enviada para a pasta central de todas as ordens.
     *
     * @return mensagem sobre a conclusão do processamento do ZIP pai
     */
    @PostMapping("/zip-parent")
    public String organizarZipPai() {
        try {
            service.processarZipPai();
            return "Processo concluído (zip pai).";
        } catch (Exception e) {
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }
}
