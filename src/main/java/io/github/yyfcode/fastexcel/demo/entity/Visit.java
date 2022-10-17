package io.github.yyfcode.fastexcel.demo.entity;

import java.util.Date;

import io.github.yyfcode.fastexcel.annotation.ExcelProperty;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty.Header;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Simple JavaBean domain object representing a visit.
 * @author Justice
 */
@Data
@AllArgsConstructor
public class Visit {

	@ExcelProperty(name = "name", column = 0,
		header = @Header(fillForegroundColor = IndexedColors.CORAL, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String name;

	@ExcelProperty(name = "visitDate", column = 1, format = "yyyy/MM/dd HH:mm:ss", width = 30,
		header = @Header(fillForegroundColor = IndexedColors.BROWN, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private Date visitDate;

	@ExcelProperty(name = "description", column = 2, width = 20,
		header = @Header(fillForegroundColor = IndexedColors.SKY_BLUE, fillPatternType = FillPatternType.SOLID_FOREGROUND))
	private String description;
}
