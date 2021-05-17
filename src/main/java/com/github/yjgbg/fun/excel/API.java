package com.github.yjgbg.fun.excel;

import lombok.experimental.UtilityClass;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Arrays;

@UtilityClass
public class API {
	public Workbook Workbook(Sheet... sheets) {
		final var res = new HSSFWorkbook();
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
		return Cell.create().setText(text);
	}

	public Cell Cell(int cellNum) {
		return Cell.cell(cellNum);
	}
}
