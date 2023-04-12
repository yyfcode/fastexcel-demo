package com.jeeapp.excel.demo.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
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
@NoArgsConstructor
@AllArgsConstructor
public class Food {

	@ExcelProperty(name = "name2", column = 0,width = 20,
		header = @Header(fillForegroundColor = IndexedColors.TURQUOISE, fillPatternType = FillPatternType.SOLID_FOREGROUND,
			comment = @Comment("Health status of pets")),
		validation = @Validation(validationType = ValidationType.LIST, explicitListValues = {"meat", "water", "egg"}))
	private String name;

	@ExcelProperty(name = "quantity", column = 1,width = 20,
		header = @Header(fillForegroundColor = IndexedColors.DARK_RED, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private Integer quantity;
}
