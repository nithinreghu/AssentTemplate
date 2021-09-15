package com.flex.assenttemplate;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.flex.assenttemplate.util.ValidationUtil;

@SpringBootApplication
public class AssentTemplateApplication implements ApplicationRunner {

	@Value("${bomTemplateFileName}")
	private String bomTemplateFileName;

	@Value("${bomTemplateFirstRow:2}")
	private Integer bomTemplateFirstRow;

	@Value("${mstrFileName}")
	private String mstrFileName;

	@Value("${mstrFirstRow:2}")
	private Integer mstrFirstRow;

	public static void main(String[] args) {
		SpringApplication.run(AssentTemplateApplication.class, args);
	}

	@Override
	public void run(ApplicationArguments args) throws Exception {
		ValidationUtil.validateBomTemplate(bomTemplateFileName, bomTemplateFirstRow, mstrFileName, mstrFirstRow);
	}

}
