package com.jeeapp.excel.demo.entity;

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
 * @author Justice
 */
@Data
@AllArgsConstructor
public class Food {

	@ExcelProperty(name = "name", column = 0,
		header = @Header(fillForegroundColor = IndexedColors.TURQUOISE, fillPatternType = FillPatternType.SOLID_FOREGROUND,
			comment = @Comment("Health status of pets")),
		validation = @Validation(validationType = ValidationType.LIST, explicitListValues = {"dog", "cat"}))
	private String name;

	@ExcelProperty(name = "quantity", column = 1,
		header = @Header(fillForegroundColor = IndexedColors.DARK_RED, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private Integer quantity;
}
