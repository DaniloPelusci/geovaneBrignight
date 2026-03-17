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

    public FotoInspecao(String nomeArquivo, String tipoConteudo, byte[] conteudo) {
        this.nomeArquivo = nomeArquivo;
        this.tipoConteudo = tipoConteudo;
        this.conteudo = conteudo;
    }

    void setInspecao(Inspecao inspecao) {
        this.inspecao = inspecao;
    }
}
