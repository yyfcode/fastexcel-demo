package io.github.yyfcode.fastexcel.demo.entity;

import java.util.Date;
import java.util.Set;

import io.github.yyfcode.fastexcel.annotation.ExcelProperty;
import io.github.yyfcode.fastexcel.annotation.ExcelProperty.Header;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * Simple business object representing a pet.
 * @author Justice
 */
@Data
@AllArgsConstructor
public class Pet {

	@ExcelProperty(name = "name", column = 0, width = 20)
	private String name;

	@ExcelProperty(name = "type", column = 1, width = 20)
	private String type;

	@ExcelProperty(name = "birthDate", column = 2, format = "yyyy/MM/dd", width = 20)
	private Date birthDate;

	@ExcelProperty(name = "visits", column = 3, header = @Header(fillBackgroundColor = IndexedColors.ORANGE))
	private Set<Visit> visits;
}
