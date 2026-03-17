package br.com.portfoliopelusci.inspecao.repository;

import br.com.portfoliopelusci.inspecao.entity.Inspecao;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.Optional;

public interface InspecaoRepository extends JpaRepository<Inspecao, Long> {

    Optional<Inspecao> findByInspetorIdAndWorder(String inspetorId, String worder);
}
