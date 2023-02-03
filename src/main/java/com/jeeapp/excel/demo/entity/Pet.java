package com.jeeapp.excel.demo.entity;

import java.util.Date;
import java.util.List;
import java.util.Set;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Comment;
import com.jeeapp.excel.annotation.ExcelProperty.Header;
import com.jeeapp.excel.annotation.ExcelProperty.Validation;

/**
 * Simple business object representing a pet.
 * @author Justice
 */
@Data
@AllArgsConstructor
public class Pet {

	@ExcelProperty(name = "name", column = 0, width = 20,
		header = @Header(fillForegroundColor = IndexedColors.DARK_GREEN, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String name;

	@ExcelProperty(name = "type", column = 1, width = 20,
		header = @Header(fillForegroundColor = IndexedColors.BLUE, fillPatternType = FillPatternType.SOLID_FOREGROUND),
		validation = @Validation(validationType = ValidationType.LIST, explicitListValues = {"dog", "cat"}))
	private String type;

	@ExcelProperty(name = "birthDate", column = 2, format = "yyyy-MM-dd", width = 20,
		header = @Header(fillForegroundColor = IndexedColors.GREEN, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private Date birthDate;

	@ExcelProperty(name = "health", column = 2, format = "00", width = 20,
		header = @Header(fillForegroundColor = IndexedColors.YELLOW, fillPatternType = FillPatternType.SOLID_FOREGROUND,
		comment = @Comment("Health status of pets")))
	private Integer health;

	@ExcelProperty(name = "visits", column = 3,
		header = @Header(fillForegroundColor = IndexedColors.ORANGE, fillPatternType = FillPatternType.SOLID_FOREGROUND,
			comment = @Comment("visit list")))
	private Set<Visit> visits;

	@ExcelProperty(name = "foods", column = 4,
		header = @Header(fillForegroundColor = IndexedColors.TEAL, fillPatternType = FillPatternType.SOLID_FOREGROUND,
			comment = @Comment("food list")))
	private List<Food> foods;
}
