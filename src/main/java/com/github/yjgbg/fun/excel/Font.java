package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.jetbrains.annotations.Contract;

public interface Font {
	org.apache.poi.ss.usermodel.Font toFont(Workbook workbook);

	@Contract(pure = true)
	static Font create() {
		return Workbook::createFont;
	}

	default Font setBold(boolean bold) {
		final var that = this;
		return workbook -> {
			final var res =  that.toFont(workbook);
			res.setBold(bold);
			return res;
		};
	}

	static Font bold() {
		return create().setBold(true);
	}
}
