package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.*;

public interface CellStyle {
	default org.apache.poi.ss.usermodel.CellStyle toCellStyle(Workbook workbook) {
		final var res = workbook.createCellStyle();
		post(res,workbook);
		return res;
	}

	static CellStyle create() {
		return (cellStyle,workbook) -> cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}

	static CellStyle alignment(HorizontalAlignment alignment) {
		return (cellStyle,workbook) -> cellStyle.setAlignment(alignment);
	}
	static CellStyle color(IndexedColors color) {
		return (cellStyle,workbook) -> cellStyle.setFillForegroundColor(color.getIndex());
	}

	static CellStyle wrapText() {
		return (cellStyle,workbook) -> cellStyle.setWrapText(true);
	}

	static CellStyle borderColor(IndexedColors color) {
		return (cellStyle,workbook) -> {
			cellStyle.setLeftBorderColor(color.getIndex());
			cellStyle.setRightBorderColor(color.getIndex());
			cellStyle.setTopBorderColor(color.getIndex());
			cellStyle.setBottomBorderColor(color.getIndex());
		};
	}

	static CellStyle border(BorderStyle borderStyle) {
		return (cellStyle,workbook) -> {
			cellStyle.setBorderLeft(borderStyle);
			cellStyle.setBorderRight(borderStyle);
			cellStyle.setBorderTop(borderStyle);
			cellStyle.setBorderBottom(borderStyle);
		};
	}


	static CellStyle plus(CellStyle one, CellStyle another) {
		return (cellStyle,workbook) -> {
			one.post(cellStyle,workbook);
			another.post(cellStyle,workbook);
		};
	}

	void post(org.apache.poi.ss.usermodel.CellStyle cellStyle,Workbook workbook);
}
