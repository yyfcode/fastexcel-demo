package io.github.yyfcode.fastexcel.demo.entity;

import javax.validation.constraints.NotBlank;

import lombok.Data;

/**
 * Simple JavaBean domain object representing an person.
 * @author Justice
 */
@Data
public class Person {

	@NotBlank
	private String fullName;
}
