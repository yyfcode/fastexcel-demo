package com.jeeapp.excel.demo.entity;

import java.util.Date;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import com.jeeapp.excel.annotation.ExcelProperty;
import com.jeeapp.excel.annotation.ExcelProperty.Header;

/**
 * Simple JavaBean domain object representing a visit.
 * @author Justice
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class Visit {

	@ExcelProperty(name = "name1", column = 0, width = 20,
		header = @Header(fillForegroundColor = IndexedColors.CORAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String name;

	@ExcelProperty(name = "visitDate", column = 1, format = "yyyy/MM/dd HH:mm:ss", width = 30,
		header = @Header(fillForegroundColor = IndexedColors.BROWN, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private Date visitDate;

	@ExcelProperty(name = "description", column = 2, width = 20,
		header = @Header(fillForegroundColor = IndexedColors.SKY_BLUE, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String description;
}
