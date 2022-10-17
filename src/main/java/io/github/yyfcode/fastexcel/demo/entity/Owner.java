package io.github.yyfcode.fastexcel.demo.entity;

import javax.validation.constraints.Digits;
import java.util.ArrayList;
import java.util.List;

import io.github.yyfcode.fastexcel.annotation.ExcelProperty;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty.Comment;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty.Header;
import lombok.Data;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Simple JavaBean domain object representing an owner.
 * @author Justice
 */
@Data
public class Owner {

	@ExcelProperty(name = "fullName", column = 0,
		header = @Header(fillForegroundColor = IndexedColors.TEAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String fullName;

	@ExcelProperty(name = "address", column = 1,
		header = @Header(fillForegroundColor = IndexedColors.TEAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String address;

	@ExcelProperty(name = "city", column = 2,
		header = @Header(fillForegroundColor = IndexedColors.TEAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String city;

	@ExcelProperty(name = "telephone", column = 3,
		header = @Header(fillForegroundColor = IndexedColors.TEAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	@Digits(fraction = 0, integer = 10)
	private String telephone;

	@ExcelProperty(name = "pets", column = 4,
		header = @Header(fillForegroundColor = IndexedColors.ROSE, fillPatternType = FillPatternType.SOLID_FOREGROUND,
			comment = @Comment("pet list")))
	private List<Pet> pets = new ArrayList<>();

	public void addPet(Pet pet) {
		pets.add(pet);
	}
}
