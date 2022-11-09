package com.jeeapp.excel.demo.controller;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URL;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DataValidation.ErrorStyle;
import org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
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
import com.jeeapp.excel.builder.SheetBuilder;
import com.jeeapp.excel.builder.WorkbookBuilder;
import com.jeeapp.excel.demo.entity.Owner;
import com.jeeapp.excel.demo.entity.Pet;
import com.jeeapp.excel.demo.entity.Visit;
import com.jeeapp.excel.util.CellUtils;

/**
 * @author Justice
 */
@Tag(name = "Excel write")
@Controller
@RequestMapping("excelWrite")
public class ExcelWriteController {

	@Operation(summary = "Create sheet")
	@PostMapping(value = "createSheet", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createSheet(HttpServletResponse response) throws Exception {
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
			response.setHeader("Content-disposition", "attachment; filename=createSheet.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Add default style")
	@PostMapping(value = "addDefaultStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void addDefaultCellStyle(HttpServletResponse response) throws Exception {
		// Default cell style
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(50)
			.setDefaultColumnWidth(50)
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
			response.setHeader("Content-disposition", "attachment; filename=addDefaultStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Create picture")
	@PostMapping(value = "createPicture", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createPicture(HttpServletResponse response) throws Exception {
		byte[] bytes = IOUtils.toByteArray(new URL("https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png")
			.openStream());
		Workbook workbook = new WorkbookBuilder(new XSSFWorkbook())
			.setDefaultRowHeight(100)
			.createSheet("Sheet 1")
			.addCellRange(0, 0, 0, 5)
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.insert()
			.end()
			.createCell(null)
			.setRowHeight(50)
			.matchingActiveCell()
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.insert()
			.end()
			.createCell(5, 5, null)
			.setRowHeight(50)
			.matchingActiveCell()
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.setSize(2, 2)
			.insert()
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createPicture.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Create data validation")
	@PostMapping(value = "createDataValidation", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createDataValidation(HttpServletResponse response) throws Exception {
		Workbook workbook = new WorkbookBuilder(new XSSFWorkbook())
			.createSheet("Sheet 1")
			.matchingRegion(0, 0, 0, 0)
			.createExplicitListConstraint("a", "b")
			.showErrorBox("error", "wrong data")
			.showPromptBox("hint", "select a or b")
			.setErrorStyle(ErrorStyle.INFO)
			.matchingRegion(0, 0, 1, 1)
			.createDateConstraint(OperatorType.BETWEEN, "2022-01-01", "2022-12-30", "yyyy-MM-dd")
			.showErrorBox("error", "wrong date format")
			.showPromptBox("hint", "yyyy-MM-dd")
			.matchingRegion(0, 0, 2, 2)
			.createIntegerConstraint(OperatorType.BETWEEN, "50", "100")
			.showErrorBox("error", "wrong number")
			.showPromptBox("hint", "must be 50~100")
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createDataValidation.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Add cell style")
	@PostMapping(value = "addCellStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void addCellStyle(HttpServletResponse response) throws Exception {
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
			.setRowHeight(100)
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
			response.setHeader("Content-disposition", "attachment; filename=addCellStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "add column style")
	@PostMapping(value = "addColumnStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void customColumnStyle(HttpServletResponse response) throws Exception {
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(40)
			.createSheet("Sheet 1")
			.matchingColumn(0)
			.setDataFormat("yyyy-MM-dd")
			.addCellStyle()
			.matchingColumn(1)
			.setDataFormat("#.##00")
			.setFontColor(IndexedColors.RED)
			.addCellStyle()
			.matchingColumn(2)
			.setDataFormat("[=1]\"male\";[=2]\"female\"")
			.addCellStyle()
			.createRow(new Object[]{new Date(), 22.121f, 1})
			.createRow(new Object[]{new Date(), 123.1d, 2})
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=addColumnStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Merge cell range")
	@PostMapping(value = "mergeCellRange", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void mergeCellRange(HttpServletResponse response) throws Exception {
		Workbook workbook = new WorkbookBuilder(new SXSSFWorkbook())
			.createSheet("Sheet 1")
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.createRow(new Object[]{"cell1", "cell2", "cell3", "cell4"})
			.addCellRange(1, 2, 1, 2)
			.setCellValue("cool")
			.merge()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=mergeCellRange.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Create bean row")
	@PostMapping(value = "createBeanRow", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createBeanRow(HttpServletResponse response) throws Exception {
		List<Owner> owners = createOwners();
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.rowType(Owner.class)
			// Partial fields
			.createHeader("fullName", "address", "city", "telephone")
			.createRows(owners)
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createBeanRow.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Create bean row with partial fields")
	@PostMapping(value = "createBeanRowWithPartialFields", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createBeanRowWithPartialFields(HttpServletResponse response) throws Exception {
		List<Owner> owners = createOwners();
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(30)
			.setDefaultColumnWidth(30)
			.createSheet("Sheet 1")
			.rowType(Owner.class)
			// partial fields
			.createHeader("fullName", "address", "city", "telephone", "pets.name", "pets.birthday", "pets.type")
			.createRows(owners)
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createBeanRowWithPartialFields.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Create nested bean row")
	@PostMapping(value = "createNestedBeanRow", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createNestedBeanRow(HttpServletResponse response) throws Exception {
		List<Owner> owners = createOwners();
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
			.createRows(owners)
			.end()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createNestedBeanRow.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "create nested bean sheet")
	@PostMapping(value = "createNestedBeanSheet", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void createNestedBeanSheet(HttpServletResponse response) throws Exception {
		List<Owner> owners = createOwners();
		WorkbookBuilder workbookBuilder = WorkbookBuilder.builder()
			.setDefaultRowHeight(30)
			.setDefaultColumnWidth(30);
		SheetBuilder sheetBuilder = null;
		for (Owner owner : owners) {
			sheetBuilder = workbookBuilder.createSheet(owner.getFullName())
				.rowType(Owner.class)
				.createHeader("fullName", "address", "city", "telephone")
				.createRow(owner)
				.end()
				.createRow()
				.createRow(new Object[]{"Pets"})
				.addCellRange(3, 3, 0, 6)
				.merge()
				.setRowHeight(50)
				.rowType(Pet.class)
				.createHeader()
				.createRows(owner.getPets())
				.end();
		}
		if (sheetBuilder != null) {
			Workbook workbook = sheetBuilder.build();
			try (ServletOutputStream out = response.getOutputStream()) {
				response.setHeader("Content-disposition", "attachment; filename=createNestedBeanSheet.xlsx");
				workbook.write(out);
			}
		}
	}

	private List<Owner> createOwners() {
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
		return Arrays.asList(george, joe);
	}
}
