package br.com.portfoliopelusci.inspecao.entity;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "inspetor")
public class Inspetor {

    @Id
    @Column(name = "id", nullable = false, length = 100)
    private String id;

    protected Inspetor() {
    }

    public Inspetor(String id) {
        this.id = id;
    }

    public String getId() {
        return id;
    }
}
