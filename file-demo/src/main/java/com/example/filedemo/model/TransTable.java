package com.example.filedemo.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name="TRTABLE")

public class TransTable {
	
	
	@Id
	@GeneratedValue(strategy = GenerationType.AUTO)

	@Column(name="ID")
	private Long id;

	@Column(name="STATUS")
	private String status;
	
	@Column(name="TEAM")
	private String team;
	
	@Column(name="FILE_NAMES")
	private String fileName;
	
	public Long getId() {
		return id;
	}
	public void setId(Long id) {
		this.id = id;
	}
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	public String getTeam() {
		return team;
	}
	public void setTeam(String team) {
		this.team = team;
	}
	public String getFileName() {
		return fileName;
	}
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}


}
