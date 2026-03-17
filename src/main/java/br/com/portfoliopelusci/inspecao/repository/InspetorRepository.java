package br.com.portfoliopelusci.inspecao.repository;

import br.com.portfoliopelusci.inspecao.entity.Inspetor;
import org.springframework.data.jpa.repository.JpaRepository;

public interface InspetorRepository extends JpaRepository<Inspetor, String> {
}
