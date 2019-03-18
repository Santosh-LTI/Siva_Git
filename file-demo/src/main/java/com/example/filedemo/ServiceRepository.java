package com.example.filedemo;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.example.filedemo.model.TransTable;

@Repository
public interface ServiceRepository extends CrudRepository<TransTable, Long>{



	

}
