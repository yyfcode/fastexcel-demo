package io.github.yyfcode.fastexcel.demo.entity;

import javax.validation.constraints.Digits;
import java.util.ArrayList;
import java.util.List;

import io.github.yyfcode.fastexcel.annotation.ExcelProperty;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty.Header;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Simple JavaBean domain object representing an owner.
 * @author Justice
 */
@Data
public class Owner {

	@ExcelProperty(name = "fullName", column = 0, width = 20)
	private String fullName;

	@ExcelProperty(name = "address", column = 1, width = 35)
	private String address;

	@ExcelProperty(name = "city", column = 2, width = 20)
	private String city;

	@ExcelProperty(name = "telephone", column = 3, width = 20)
	@Digits(fraction = 0, integer = 10)
	private String telephone;

	@ExcelProperty(name = "pets", column = 4, header = @Header(fillBackgroundColor = IndexedColors.GREY_25_PERCENT))
	private List<Pet> pets = new ArrayList<>();

	public void addPet(Pet pet) {
		pets.add(pet);
	}
}
