package com.flex.assenttemplate;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.flex.assenttemplate.service.AssentService;
import com.flex.assenttemplate.service.ValidationService;

@SpringBootApplication
public class AssentTemplateApplication implements ApplicationRunner {

	@Value("${userInput}")
	private int userInput;

	@Autowired
	AssentService assentService;

	@Autowired
	ValidationService validationService;

	public static void main(String[] args) {
		SpringApplication.run(AssentTemplateApplication.class, args);
	}

	@Override
	public void run(ApplicationArguments args) throws Exception {

		long time = System.currentTimeMillis();

		if (userInput == 1) {

			System.out.println("..................................................................");
			System.out.println("........Validating BOM template...................................");
			validationService.validateBomTemplate();

		} else if (userInput == 2) {

			System.out.println("..................................................................");
			System.out.println("........Generating Assents template...............................");

			assentService.generateAssentTemplate();

		} else {
			throw new RuntimeException("....Invalid value... Enter 1 or 2 and try again..");
		}

		System.out.println("..................................................................");
		System.out.println("..................................................................");
		System.out.println(
				"....Time taken to complete the process: " + (System.currentTimeMillis() - time) / 1000 + " seconds");
		System.out.println("..................................................................");
		System.out.println("..................................................................");

	}

}
