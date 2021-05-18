package com.github.yjgbg.fun.excel;

import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;

@UtilityClass
public class API {
	public Workbook Workbook(Sheet... sheets) {
		final var res = new XSSFWorkbook();
		Arrays.stream(sheets).forEachOrdered(sheet -> sheet.toSheet(res));
		return res;
	}

	public Sheet Sheet(String name, Row... rows) {
		return Arrays.stream(rows).reduce(Sheet.sheet(name), Sheet::addRow, (a, b) -> null);
	}

	public Sheet Sheet(Row... rows) {
		return Arrays.stream(rows).reduce(Sheet.create(), Sheet::addRow, (a, b) -> null);
	}

	public Sheet Sheet(int index,Row... rows) {
		return Arrays.stream(rows).reduce(Sheet.sheet(index), Sheet::addRow, (a, b) -> null);
	}

	public Row Row(Cell... cells) {
		return Arrays.stream(cells).reduce(Row.create(), Row::addCell, (a, b) -> null);
	}

	public Row Row(int rowNum,Cell... cells) {
		return Arrays.stream(cells).reduce(Row.row(rowNum), Row::addCell, (a, b) -> null);
	}

	public Cell Cell() {
		return Cell.create();
	}

	public Cell Cell(String text) {
		return Cell.create().content(RichText.of(text));
	}


	public Cell Cell(RichText text) {
		return Cell.create().content(text);
	}

	public Cell Cell(LocalDateTime dateTime) {
		return Cell.create().content(dateTime);
	}

	public Cell Cell(LocalDate date,String fmt) {
		return Cell.create().content(date,fmt);
	}

	public Cell Cell(int cellNum) {
		return Cell.create(cellNum);
	}
}
