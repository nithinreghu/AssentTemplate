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

		if (userInput == 1) {
			validationService.validateBomTemplate();

		} else if (userInput == 2) {
			assentService.generateAssentTemplate();

		} else {
			throw new RuntimeException("Invalid value... Enter 1 or 2 and try again..");
		}

	}

}
