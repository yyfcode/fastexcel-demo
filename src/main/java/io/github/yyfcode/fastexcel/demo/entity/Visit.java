package io.github.yyfcode.fastexcel.demo.entity;

import java.util.Date;

import io.github.yyfcode.fastexcel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * Simple JavaBean domain object representing a visit.
 * @author Justice
 */
@Data
@AllArgsConstructor
public class Visit {

	@ExcelProperty(name = "visitDate", column = 0, format = "yyyy/MM/dd HH:mm:ss", width = 20)
	private Date visitDate;

	@ExcelProperty(name = "description", column = 1, width = 20)
	private String description;
}
