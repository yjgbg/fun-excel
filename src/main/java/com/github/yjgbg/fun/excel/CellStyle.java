package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Workbook;

public interface CellStyle {
	org.apache.poi.ss.usermodel.CellStyle toCellStyle(Workbook workbook);
	static CellStyle create() {
		return Workbook::createCellStyle;
	}

	default CellStyle setFont(Font font) {
		final var that = this;
		return workbook -> {
			final var res = that.toCellStyle(workbook);
			res.setFont(font.toFont(workbook));
			return res;
		};
	}

	static CellStyle bold() {
		return CellStyle.create().setFont(Font.create().setBold(true));
	}
}
