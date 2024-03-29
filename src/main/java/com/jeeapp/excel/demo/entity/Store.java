package com.jeeapp.excel.demo.entity;

import javax.validation.constraints.NotBlank;
import java.util.Date;

import lombok.Data;
import lombok.NoArgsConstructor;
import com.jeeapp.excel.annotation.ExcelProperty;

/**
 * @author Justice
 */
@Data
@NoArgsConstructor
public class Store {

	@NotBlank
	@ExcelProperty(name = "name", column = 0, width = 20)
	private String name;

	@ExcelProperty(name = "address", column = 1, width = 20)
	private String address;

	@ExcelProperty(name = "createDate", column = 2, width = 20, format = "dd/MM/yy")
	private Date createDate;
}
