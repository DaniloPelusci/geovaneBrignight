package br.com.portfoliopelusci.inspecao.entity;

import jakarta.persistence.*;

@Entity
@Table(name = "foto_inspecao")
public class FotoInspecao {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "nome_arquivo", nullable = false)
    private String nomeArquivo;

    @Column(name = "descricao", nullable = false)
    private String descricao;

    @Column(name = "tipo_conteudo")
    private String tipoConteudo;

    @Lob
    @Basic(fetch = FetchType.LAZY)
    @Column(name = "conteudo", nullable = false, columnDefinition = "LONGBLOB")
    private byte[] conteudo;

    @ManyToOne(fetch = FetchType.LAZY, optional = false)
    @JoinColumn(name = "inspecao_id", nullable = false)
    private Inspecao inspecao;

    protected FotoInspecao() {
    }

    public FotoInspecao(String nomeArquivo, String descricao, String tipoConteudo, byte[] conteudo) {
        this.nomeArquivo = nomeArquivo;
        this.descricao = descricao;
        this.tipoConteudo = tipoConteudo;
        this.conteudo = conteudo;
    }

    void setInspecao(Inspecao inspecao) {
        this.inspecao = inspecao;
    }

    public String getDescricao() {
        return descricao;
    }

    public String getNomeArquivo() {
        return nomeArquivo;
    }
}
