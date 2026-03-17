package br.com.portfoliopelusci.inspecao.entity;

import jakarta.persistence.*;

import java.util.ArrayList;
import java.util.List;

@Entity
@Table(name = "inspecao",
        uniqueConstraints = @UniqueConstraint(name = "uk_inspecao_inspetor_worder", columnNames = {"inspetor_id", "worder"}))
public class Inspecao {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "worder", nullable = false, length = 100)
    private String worder;

    @ManyToOne(fetch = FetchType.LAZY, optional = false)
    @JoinColumn(name = "inspetor_id", nullable = false)
    private Inspetor inspetor;

    @OneToMany(mappedBy = "inspecao", cascade = CascadeType.ALL, orphanRemoval = true)
    private List<FotoInspecao> fotos = new ArrayList<>();

    protected Inspecao() {
    }

    public Inspecao(String worder, Inspetor inspetor) {
        this.worder = worder;
        this.inspetor = inspetor;
    }

    public void adicionarFoto(FotoInspecao foto) {
        fotos.add(foto);
        foto.setInspecao(this);
    }

    public Long getId() {
        return id;
    }

    public String getWorder() {
        return worder;
    }

    public Inspetor getInspetor() {
        return inspetor;
    }
}
