package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Sheet;

public interface Row {
	org.apache.poi.ss.usermodel.Row toRow(Sheet sheet);
	static Row create() {
		return sheet -> {
			final var last = sheet.getLastRowNum();
			return sheet.createRow(last+1);
		};
	}

	static Row row(int rowNum) {
		return sheet -> {
			final var res = sheet.getRow(rowNum);
			return res!=null ? res : sheet.createRow(rowNum);
		};
	}

	default Row addCell(Cell cell) {
		final var that = this;
		return sheet -> {
			final var res = that.toRow(sheet);
			cell.toCell(res);
			return res;
		};
	}
}
