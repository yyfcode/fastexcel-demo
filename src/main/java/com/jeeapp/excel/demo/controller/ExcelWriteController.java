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
import com.jeeapp.excel.builder.WorkbookBuilder;
import com.jeeapp.excel.demo.entity.Food;
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
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.createSheet("Sheet 2")
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.createSheet("Sheet 3")
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createSheet.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "Add default style")
	@PostMapping(value = "addDefaultStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void addDefaultStyle(HttpServletResponse response) throws Exception {
		// Default cell style
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultRowHeight(50)
			.setDefaultColumnWidth(50)
			.createSheet("Sheet 1")
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.matchingLastRow()
			.setRowHeight(100)
			.addCellStyle()
			.createSheet("Sheet 2")
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.createSheet("Sheet 3")
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
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
			.setDefaultRowHeight(60)
			.createSheet("Sheet 1")
			.setDefaultColumnWidth(25)
			.matchingRegion(0, 0, 0, 5)
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.addMergedRegion()
			.matchingLastCell()
			.createCellComment("aaa")
			.addCellStyle()
			.matchingLastRow()
			.setRowHeight(30)
			.addCellStyle()
			.matchingCell(0, 6)
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.addCellStyle()
			.createCell(5, 5)
			.matchingLastRow()
			.setRowHeight(45)
			.addCellStyle()
			.matchingLastCell()
			.createPicture(bytes, Workbook.PICTURE_TYPE_PNG)
			.addCellStyle()
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
			.matchingCell(0, 0)
			.createExplicitListConstraint("a", "b")
			.showErrorBox("error", "wrong data")
			.showPromptBox("hint", "select a or b")
			.setErrorStyle(ErrorStyle.INFO)
			.addValidationData()
			.addCellStyle()
			.matchingCell(0, 1)
			.createDateConstraint(OperatorType.BETWEEN, "Date(2022,01,01)", "Date(2022,12,31)", "yyyy-MM-dd")
			.showErrorBox("错误", "日期错误")
			.showPromptBox("提示", "日期需要在20220101~20221231范围内")
			.addValidationData()
			// 设置格式
			.setDataFormat("yyyy-MM-dd")
			// 设置格式必须创建单元格，否则单元格填值时不能格式化
			.setBlank()
			.matchingCell(0, 2)
			.createIntegerConstraint(OperatorType.BETWEEN, "50", "100")
			.showErrorBox("error", "wrong number")
			.showPromptBox("hint", "must be 50~100")
			.addValidationData()
			.addCellStyle()
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
			.setDefaultRowHeight(50)
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
			.setDefaultColumnWidth(15)
			// Set cell style of Sheet 1
			.matchingAll()
			.setFontColor(IndexedColors.RED)
			.addCellStyle()
			.createRow("cell1", "cell2", "cell3")
			.matchingLastRow()
			.setRowHeight(200)
			.addCellStyle()
			.matchingColumn(0)
			.setColumnBreak()
			.setColumnWidth(20)
			.addCellStyle()
			.createRow("cell1", "cell2", "cell3")
			.createSheet("Sheet 2")
			// Set cell style of Sheet 2
			.matchingAll()
			.setFontColor(IndexedColors.BLUE)
			.addCellStyle()
			// Set cell style of Sheet 2 and row number > 0
			.matchingCell(cell -> cell.getRowIndex() > 0)
			.setFillForegroundColor(IndexedColors.GREY_25_PERCENT)
			.addCellStyle()
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.matchingColumn(2)
			.setColumnHidden(true)
			.addCellStyle()
			.createSheet("Sheet 3")
			.setDefaultRowHeight(30)
			// Set cell style of Sheet 3
			.matchingAll()
			.setFontColor(IndexedColors.GREEN)
			.setItalic(true)
			.addCellStyle()
			.createRow("cell1", "cell2", "cell3")
			.createRow("cell1", "cell2", "cell3")
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=addCellStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "add row style")
	@PostMapping(value = "addRowStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void addRowStyle(HttpServletResponse response) throws Exception {
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(20)
			.createSheet("Sheet 1")
			.matchingRow(0)
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.GREY_25_PERCENT)
			.setBorder(BorderStyle.THIN)
			.setBottomBorderColor(IndexedColors.RED)
			.addCellStyle()
			.matchingRow(1)
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.LIGHT_YELLOW)
			.setBorder(BorderStyle.THIN)
			.setBottomBorderColor(IndexedColors.BROWN)
			.addCellStyle()
			.matchingRow(2)
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.LIGHT_BLUE)
			.setBorder(BorderStyle.THIN)
			.setBottomBorderColor(IndexedColors.AQUA)
			.addCellStyle()
			.matchingRow(3)
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.LIGHT_ORANGE)
			.setBorder(BorderStyle.THIN)
			.setBottomBorderColor(IndexedColors.VIOLET)
			.addCellStyle()
			.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=addRowStyle.xlsx");
			workbook.write(out);
		}
	}

	@Operation(summary = "add column style")
	@PostMapping(value = "addColumnStyle", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE)
	public void addColumnStyle(HttpServletResponse response) throws Exception {
		Workbook workbook = WorkbookBuilder.builder()
			.setDefaultColumnWidth(40)
			.createSheet("Sheet 1")
			.matchingColumn(0)
			.setDataFormat("yyyy-MM-dd")
			.addCellStyle()
			.matchingRegion(1, 5, 0, 0)
			.setDataFormat("yyyy/MM/dd")
			.fillUndefinedCells()
			.matchingColumn(1)
			.setDataFormat("#.##00")
			.setFontColor(IndexedColors.RED)
			.addCellStyle()
			.matchingColumn(2)
			.setDataFormat("[=1]\"male\";[=2]\"female\"")
			.addCellStyle()
			.createRow(new Date(), 22.121f, 1)
			.createRow(new Date(), 123.1d, 2)
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
			.createRow("cell1", "cell2", "cell3", "cell4")
			.createRow("cell1", "1111", "cell3", "cell4")
			.createRow("cell1", "cell4", "cell5", "cell4")
			.createRow("cell1", "cell2", "cell3", "cell4")
			.matchingRegion(1, 2, 1, 2)
			.setBorder(BorderStyle.DOUBLE)
			.setBorderColor(IndexedColors.GOLD)
			.setFillPattern(FillPatternType.SOLID_FOREGROUND)
			.setFillForegroundColor(IndexedColors.VIOLET)
			.setFillBackgroundColor(IndexedColors.WHITE)
			.setAlignment(HorizontalAlignment.CENTER)
			.setVerticalAlignment(VerticalAlignment.CENTER)
			.addMergedRegion()
			.matchingCell(0, 0)
			.setBorderColor(IndexedColors.DARK_RED)
			.setBorder(BorderStyle.DOUBLE)
			.setFontHeight(20)
			.addCellStyle()
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
			.createHeader("fullName", "address", "city", "telephone", "pets.visits", "pets.birthday", "pets.type")
			.createRows(owners)
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
			.matchingAll()
			.setBorder(BorderStyle.DOUBLE)
			.setBorderColor(IndexedColors.CORNFLOWER_BLUE)
			.addCellStyle()
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
		for (Owner owner : owners) {
			workbookBuilder = workbookBuilder.createSheet(owner.getFullName())
				.rowType(Owner.class)
				.createHeader("fullName", "address", "city", "telephone")
				.createRow(owner)
				.createRow()
				.createRow("Pets")
				.matchingRegion(3, 3, 0, 8)
				.setBorderColor(IndexedColors.RED)
				.setBorder(BorderStyle.DOUBLE)
				.setFillForegroundColor(IndexedColors.GOLD)
				.setFillPattern(FillPatternType.SOLID_FOREGROUND)
				.addMergedRegion()
				.rowType(Pet.class)
				.createHeader()
				.createRows(owner.getPets())
				.end();
		}
		Workbook workbook = workbookBuilder.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=createNestedBeanSheet.xlsx");
			workbook.write(out);
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
		)),
			Arrays.asList(
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4)
			)));
		george.addPet(new Pet("dog2", "dog", new Date(), 95, new HashSet<>(Arrays.asList(
			new Visit("visit4", new Date(), "...."),
			new Visit("visit5", new Date(), "..."))),
			Arrays.asList(
				new Food("1", 1),
				new Food("4", 4)
			)));
		george.addPet(new Pet("dog3", "dog", new Date(), 100, new HashSet<>(Arrays.asList(
			new Visit("visit6", new Date(), "..."),
			new Visit("visit7", new Date(), "..."),
			new Visit("visit8", new Date(), "..."))),
			Arrays.asList(
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4),
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4),
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4)
			)));

		Owner joe = new Owner();
		joe.setFullName("Joe Bloggs");
		joe.setAddress("123 Caramel Street");
		joe.setCity("London");
		joe.setTelephone("01616291589");
		joe.addPet(new Pet("cat1", "cat", new Date(), 20, new HashSet<>(Arrays.asList(
			new Visit("visit9", new Date(), "..."),
			new Visit("visit10", new Date(), "..."),
			new Visit("visit11", new Date(), "..."))),
			Arrays.asList(
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4),
				new Food("1", 1),
				new Food("2", 2),
				new Food("3", 3),
				new Food("4", 4)
			)));
		joe.addPet(new Pet("cat2", "cat", new Date(), 90, new HashSet<>(Arrays.asList(
			new Visit("visit12", new Date(), "..."),
			new Visit("visit13", new Date(), "..."),
			new Visit("visit14", new Date(), "..."))),
			Arrays.asList(
				new Food("1", 1),
				new Food("2", 2)
			)));
		return Arrays.asList(george, joe);
	}
}
