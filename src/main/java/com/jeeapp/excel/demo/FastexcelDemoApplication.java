package com.jeeapp.excel.demo;

import io.swagger.v3.oas.models.info.Info;
import org.springdoc.core.GroupedOpenApi;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class FastexcelDemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(FastexcelDemoApplication.class, args);
	}

	@Bean
	public GroupedOpenApi fastexcelGroup() {
		return GroupedOpenApi.builder()
			.group("fastexcel")
			.addOpenApiCustomiser(openApi -> openApi.info(new Info().title("Fast Excel Demo API").version("v0.0.x")))
			.packagesToScan("com.jeeapp.excel.demo")
			.build();
	}
}
