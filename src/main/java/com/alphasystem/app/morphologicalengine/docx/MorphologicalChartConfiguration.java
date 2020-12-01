package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationBuilder;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * @author sali
 */
@Configuration
public class MorphologicalChartConfiguration {

    @Bean
    AbbreviatedConjugationFactory abbreviatedConjugationFactory() {
        return AbbreviatedConjugationAdapter::new;
    }

    @Bean
    DetailedConjugationFactory detailedConjugationFactory() {
        return DetailedConjugationAdapter::new;
    }

    @Bean
    SupplierFactory supplierFactory(@Autowired ConjugationBuilder conjugationBuilder) {
        return conjugationData -> new MorphologicalChartSupplier(conjugationData, conjugationBuilder);
    }

    @Bean
    MorphologicalChartEngineFactory morphologicalChartEngineFactory(
            @Autowired AbbreviatedConjugationFactory abbreviatedConjugationFactory,
            @Autowired DetailedConjugationFactory detailedConjugationFactory,
            @Autowired SupplierFactory supplierFactory) {
        return conjugationTemplate -> new MorphologicalChartEngine(abbreviatedConjugationFactory,
                detailedConjugationFactory, supplierFactory, conjugationTemplate);
    }

}
