package com.alithya.proof.poi.word.model;

import org.springframework.stereotype.Component;

@Component
public class ParagraphsObject {

	private String title;
	private String paragraphe;

	private String value01;
	private String value02;
	private String value03;


	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getParagraphe() {
		return paragraphe;
	}

	public void setParagraphe(String paragraphe) {
		this.paragraphe = paragraphe;
	}

}
