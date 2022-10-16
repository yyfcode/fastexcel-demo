package io.github.yyfcode.fastexcel.demo.entity;

import javax.validation.constraints.Digits;
import java.util.List;

import lombok.Data;

/**
 * Simple JavaBean domain object representing an owner.
 *
 * @author Justice
 */
@Data
public class Owner extends Person{

	private String address;

	private String city;

	@Digits(fraction = 0, integer = 10)
	private String telephone;

	private List<Pet> pets;
}
