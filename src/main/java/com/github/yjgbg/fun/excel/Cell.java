package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Row;

public interface Cell {
	org.apache.poi.ss.usermodel.Cell toCell(Row row);

	static Cell create() {
		return row -> {
			final var lastCellNum = row.getLastCellNum();
			return row.createCell(lastCellNum== -1 ? 0 : lastCellNum);
		};
	}

	static Cell cell(int cellNum) {
		if (cellNum < 0) throw new IllegalArgumentException();
		return row -> {
			final var res = row.getCell(cellNum);
			return res != null ? res : row.createCell(cellNum);
		};
	}

	default Cell setText(String text) {
		final var that = this;
		return row -> {
			final var res = that.toCell(row);
			res.setCellValue(text);
			return res;
		};
	}

	default Cell setCellStyle(CellStyle style) {
		final var that = this;
		return row -> {
			final var res = that.toCell(row);
			final var workBook = res.getRow().getSheet().getWorkbook();
			res.setCellStyle(style.toCellStyle(workBook));
			return res;
		};
	}
}
