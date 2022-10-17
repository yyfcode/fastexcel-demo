package io.github.yyfcode.fastexcel.demo.controller;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;

import io.github.yyfcode.fastexcel.builder.WorkbookBuilder;
import io.github.yyfcode.fastexcel.demo.entity.Owner;
import io.github.yyfcode.fastexcel.demo.entity.Pet;
import io.github.yyfcode.fastexcel.demo.entity.Visit;
import io.github.yyfcode.fastexcel.util.CellUtils;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
		Workbook workbook = new WorkbookBuilder(new XSSFWorkbook())
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
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.WHITE)
			.setBorderColor(IndexedColors.GREY_25_PERCENT)
			.setBorder(BorderStyle.THIN)
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

	@Operation(summary = "Simple object write")
	@PostMapping(value = "simpleObjectWrite", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void javaBeanWrite(HttpServletResponse response) throws Exception {
		Owner george = new Owner();
		george.setFullName("George Franklin");
		george.setAddress("110 W. Liberty St.");
		george.setCity("Madison");
		george.setTelephone("6085551023");
		Owner joe = new Owner();
		joe.setFullName("Joe Bloggs");
		joe.setAddress("123 Caramel Street");
		joe.setCity("London");
		joe.setTelephone("01616291589");
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.rowType(Owner.class)
			// All fields
//			.createHeader()
			// Partial fields
			.createHeader("fullName", "address", "city", "telephone")
			.createRows(Arrays.asList(george, joe))
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=simpleObjectWrite.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Nested object write")
	@PostMapping(value = "nestedObjectWrite", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void nestedObjectWrite(HttpServletResponse response) throws Exception {
		Owner george = new Owner();
		george.setFullName("George Franklin");
		george.setAddress("110 W. Liberty St.");
		george.setCity("Madison");
		george.setTelephone("6085551023");
		george.addPet(new Pet("dog1", "dog", new Date(), 50, null));
		george.addPet(new Pet("dog2", "dog", new Date(), 95, null));
		george.addPet(new Pet("dog3", "dog", new Date(), 100, null));

		Owner joe = new Owner();
		joe.setFullName("Joe Bloggs");
		joe.setAddress("123 Caramel Street");
		joe.setCity("London");
		joe.setTelephone("01616291589");
		joe.addPet(new Pet("cat1", "cat", new Date(), 20, null));
		joe.addPet(new Pet("cat2", "cat", new Date(), 90, null));

		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(30)
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.rowType(Owner.class)
			// All fields
//			.createHeader()
			// Partial fields
			.createHeader("fullName", "address", "city", "telephone", "pets.name", "pets.birthday", "pets.type")
			.createRows(Arrays.asList(george, joe))
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=nestedObjectWrite.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Nested object write 2")
	@PostMapping(value = "nestedObjectWrite2", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void nestedObjectWrite2(HttpServletResponse response) throws Exception {
		Owner george = new Owner();
		george.setFullName("George Franklin");
		george.setAddress("110 W. Liberty St.");
		george.setCity("Madison");
		george.setTelephone("6085551023");
		george.addPet(new Pet("dog1", "dog", new Date(), 50, new HashSet<>(Arrays.asList(
			new Visit("visit1", new Date(), "..."),
			new Visit("visit2", new Date(), "..."),
			new Visit("visit3", new Date(), "...")
		))));
		george.addPet(new Pet("dog2", "dog", new Date(), 95, new HashSet<>(Arrays.asList(
			new Visit("visit4", new Date(), "...."),
			new Visit("visit5", new Date(), "...")
		))));
		george.addPet(new Pet("dog3", "dog", new Date(), 100, new HashSet<>(Arrays.asList(
			new Visit("visit6", new Date(), "..."),
			new Visit("visit7", new Date(), "..."),
			new Visit("visit8", new Date(), "...")
		))));

		Owner joe = new Owner();
		joe.setFullName("Joe Bloggs");
		joe.setAddress("123 Caramel Street");
		joe.setCity("London");
		joe.setTelephone("01616291589");
		joe.addPet(new Pet("cat1", "cat", new Date(), 20, new HashSet<>(Arrays.asList(
			new Visit("visit9", new Date(), "..."),
			new Visit("visit10", new Date(), "..."),
			new Visit("visit11", new Date(), "...")
		))));
		joe.addPet(new Pet("cat2", "cat", new Date(), 90, new HashSet<>(Arrays.asList(
			new Visit("visit12", new Date(), "..."),
			new Visit("visit13", new Date(), "..."),
			new Visit("visit14", new Date(), "...")
		))));

		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(30)
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.matchingCell(cell -> {
				// match health column
				if (cell == null || cell.getColumnIndex() != 7) {
					return false;
				}
				// health < 60
				String cellValue = CellUtils.getCellValue(cell);
				if (NumberUtils.isCreatable(cellValue)) {
					return Integer.parseInt(cellValue) < 60;
				}
				return false;
			})
			.setStrikeout(true)
			.setFontHeight(30)
			.setFontColor(IndexedColors.RED.getIndex())
			.addCellStyle()
			.rowType(Owner.class)
			// All fields
			.createHeader()
			.createRows(Arrays.asList(george, joe))
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=nestedObjectWrite2.xlsx");
			workbook.write(out);
		}
	}
}
