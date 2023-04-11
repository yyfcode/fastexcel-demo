package com.jeeapp.excel.demo.entity;

import javax.validation.constraints.Digits;
import java.util.ArrayList;
import java.util.List;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Comment;
import com.jeeapp.excel.annotation.ExcelProperty.Header;

/**
 * Simple JavaBean domain object representing an owner.
 * @author Justice
 */
@Data
@NoArgsConstructor
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
