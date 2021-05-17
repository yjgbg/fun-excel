package com.github.yjgbg.fun.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public interface Sheet {
	org.apache.poi.ss.usermodel.Sheet toSheet(Workbook workbook);

	static Sheet sheet(String name) {
		return workbook -> {
			final var res = workbook.getSheet(name);
			return res != null ? res : workbook.createSheet(name);
		};
	}

	static Sheet create() {
		return Workbook::createSheet;
	}

	static Sheet sheet(int index) {
		if (index <=0) throw new IllegalArgumentException();
		return workbook -> {
			final var res = workbook.getSheetAt(index);
			return res != null ? res : workbook.createSheet();
		};
	}

	default Sheet addRow(Row row) {
		final var that = this;
		return workbook -> {
			final var res = that.toSheet(workbook);
			row.toRow(res);
			return res;
		};
	}

	default Workbook asWorkbook() {
		Workbook workbook = new HSSFWorkbook();
		toSheet(workbook);
		return workbook;
	}
}
