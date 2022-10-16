package io.github.yyfcode.fastexcel.demo.entity;

import java.util.Date;
import java.util.LinkedHashSet;
import java.util.Set;

import lombok.Data;

/**
 * Simple business object representing a pet.
 * @author Justice
 */
@Data
public class Pet {

	private String name;

	private Date birthDate;

//	private PetType type;

	private Set<Visit> visits = new LinkedHashSet<>();
}
