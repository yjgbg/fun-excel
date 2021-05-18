package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.Arrays;

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
	default Row rowStyle(CellStyle... styles) {
		final var that = this;
		return sheet -> {
			final var res = that.toRow(sheet);
			final var style = Arrays.stream(styles).reduce(CellStyle.create(),CellStyle::plus);
			res.setRowStyle(style.toCellStyle(sheet.getWorkbook()));
			return res;
		};
	}
	default Row setHeight(short height) {
		final var that = this;
		return sheet -> {
			final var res = that.toRow(sheet);
			res.setHeight(height);
			return res;
		};
	}
	default Row heightInPoints(float heightInPoints) {
		final var that = this;
		return sheet -> {
			final var res = that.toRow(sheet);
			res.setHeightInPoints(heightInPoints);
			return res;
		};
	}
	default Row zeroHeight(boolean zeroHeight) {
		final var that = this;
		return sheet -> {
			final var res = that.toRow(sheet);
			res.setZeroHeight(zeroHeight);
			return res;
		};
	}
}
