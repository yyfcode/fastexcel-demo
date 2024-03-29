package com.jeeapp.excel.demo.controller;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.convert.ConversionService;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Controller;
import org.springframework.validation.Validator;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import com.jeeapp.excel.builder.SheetBuilder;
import com.jeeapp.excel.builder.TableBuilder;
import com.jeeapp.excel.builder.WorkbookBuilder;
import com.jeeapp.excel.demo.entity.Store;
import com.jeeapp.excel.model.Comment;
import com.jeeapp.excel.model.Row;
import com.jeeapp.excel.rowset.AnnotationBasedRowSetMapper;
import com.jeeapp.excel.rowset.MappingResult;
import com.jeeapp.excel.rowset.RowSet;
import com.jeeapp.excel.rowset.RowSetReader;

/**
 * @author Justice
 */
@Slf4j
@Tag(name = "Excel read")
@Controller
@RequestMapping("excelRead")
public class ExcelReadController {

	@Autowired
	private ConversionService conversionService;

	@Autowired
	private Validator validator;

	@ResponseBody
	@Operation(summary = "Simple read")
	@PostMapping(value = "simpleRead", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
	public void simpleRead(@RequestPart("file") MultipartFile file) throws Exception {
		RowSetReader rowSetReader = RowSetReader.open(file.getInputStream());
		for (RowSet rowSet : rowSetReader) {
			System.out.println(rowSet);
		}
	}

	@ResponseBody
	@Operation(summary = "Object read")
	@PostMapping(value = "objectRead", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
	public void objectRead(@RequestPart("file") MultipartFile file) throws Exception {
		AnnotationBasedRowSetMapper<Store> rowSetMapper = new AnnotationBasedRowSetMapper<>(Store.class);
		rowSetMapper.setConversionService(conversionService);
		rowSetMapper.setValidator(validator);
		RowSetReader rowSetReader = RowSetReader.open(file.getInputStream());
		for (RowSet rowSet : rowSetReader) {
			Store store = rowSetMapper.mapRowSet(rowSet);
			System.out.println(store);
		}
	}

	@ResponseBody
	@Operation(summary = "Object read 2")
	@PostMapping(value = "objectRead2", produces = MediaType.APPLICATION_OCTET_STREAM_VALUE,
		consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
	public void objectRead2(HttpServletResponse response,
		@RequestPart("file") MultipartFile file) throws Exception {
		List<Row> errorRows = new ArrayList<>();
		AnnotationBasedRowSetMapper<Store> rowSetMapper = new AnnotationBasedRowSetMapper<>(Store.class);
		rowSetMapper.setConversionService(conversionService);
		rowSetMapper.setValidator(validator);
		RowSetReader rowSetReader = RowSetReader.open(file.getInputStream());
		for (RowSet rowSet : rowSetReader) {
			if (rowSet.getRow().getRowNum() > 0) {
				MappingResult<Store> mappingResult = rowSetMapper.getMappingResult(rowSet);
				if (!mappingResult.hasErrors()) {
					System.out.println(mappingResult.getTarget());
				} else {
					errorRows.add(rowSet.getRow());
				}
			}
		}

		TableBuilder<Store> tableBuilder = WorkbookBuilder.builder()
			.createSheet("objectRead2Result")
			.rowType(Store.class)
			.createHeader();
		for (Row errorRow : errorRows) {
			SheetBuilder sheetBuilder = tableBuilder.createRow((Object[]) errorRow.getCellValues());
			Set<Comment> comments = errorRow.getComments();
			if (CollectionUtils.isNotEmpty(comments)) {
				for (Comment comment : comments) {
					sheetBuilder.matchingLastRowCell(comment.getColNum())
						.createCellComment(comment.getText(), comment.getAuthor())
						.addCellStyle();
				}
			}
		}
		Workbook workbook = tableBuilder.build();
		try (ServletOutputStream out = response.getOutputStream()) {
			response.setHeader("Content-disposition", "attachment; filename=objectRead2Result.xlsx");
			workbook.write(out);
		}
	}
}
