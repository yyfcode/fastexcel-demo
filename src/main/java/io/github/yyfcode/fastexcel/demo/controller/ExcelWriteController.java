package io.github.yyfcode.fastexcel.demo.controller;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.util.Date;

import io.github.yyfcode.fastexcel.builder.WorkbookBuilder;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * @author Justice
 */
@Tag(name = "test excel write")
@Controller
@RequestMapping("excelWrite")
public class ExcelWriteController {

	@Operation(summary = "Simple write")
	@PostMapping(value = "simpleWrite", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void simpleWrite(HttpServletResponse response) throws Exception {
		Workbook workbook = new WorkbookBuilder(new SXSSFWorkbook())
			.createSheet("Sheet 1")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 2")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 3")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=simpleWrite.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Default cell style")
	@PostMapping(value = "defaultCellStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void defaultCellStyle(HttpServletResponse response) throws Exception {
		// Default cell style
		Workbook workbook = WorkbookBuilder.builder()
			.createSheet("Sheet 1")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 2")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 3")
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=defaultCellStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Custom cell style")
	@PostMapping(value = "customCellStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void customCellStyle(HttpServletResponse response) throws Exception {
		Workbook workbook = new WorkbookBuilder(new SXSSFWorkbook())
			// Set global cell style
			.matchingAll()
			.setFontHeight(20)
			.setFontName("Arial")
			.setVerticalAlignment(VerticalAlignment.CENTER)
			.setAlignment(HorizontalAlignment.CENTER)
			.addCellStyle()
			.createSheet("Sheet 1")
			// Set cell style of Sheet 1
			.matchingAll()
			.setFontColor(IndexedColors.RED)
			.addCellStyle()
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 2")
			// Set cell style of Sheet 2
			.matchingAll()
			.setFontColor(IndexedColors.BLUE)
			.addCellStyle()
			// Set cell style of Sheet 2 and row number > 0
			.matchingCell(cell -> cell.getRowIndex() > 0)
			.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT)
			.setFillForegroundColor(IndexedColors.GREY_25_PERCENT)
			.addCellStyle()
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createSheet("Sheet 3")
			// Set cell style of Sheet 3
			.matchingAll()
			.setFontColor(IndexedColors.GREEN)
			.setItalic(true)
			.addCellStyle()
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.createRow(new Object[]{"cell1", "cell2", "cell3"})
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=customCellStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Custom column style")
	@PostMapping(value = "customColumnStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void customColumnStyle(HttpServletResponse response) throws Exception {
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(40)
			.createSheet("Sheet 1")
			.matchingColumn(0)
			.setDataFormat("yyyy-MM-dd")
			.addCellStyle()
			.matchingColumn(1)
			.setDataFormat("#.##00")
			.addCellStyle()
			.matchingColumn(2)
			.setDataFormat("[=1]\"male\";[=2]\"female\"")
			.addCellStyle()
			.createRow(new Object[]{new Date(), 22.121f, 1})
			.createRow(new Object[]{new Date(), 123.1d, 2})
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=customColumnStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "CellRange merge")
	@PostMapping(value = "cellRangeMerge", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void cellRangeMerge(HttpServletResponse response) throws Exception {
		Workbook workbook = new WorkbookBuilder(new SXSSFWorkbook())
			.createSheet("Sheet 1")
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.addCellRange(1, 2, 1, 2)
			.merge()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=cellRangeMerge.xlsx");
			workbook.write(out);
		}
	}
}
