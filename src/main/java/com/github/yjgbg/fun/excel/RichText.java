package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

public interface RichText {
	RichTextString toRichText(Workbook workbook);

	static RichText of(String value) {
		return workbook -> new XSSFRichTextString(value);
	}

	default RichText applyFont(Font font) {
		return workbook -> {
			final var res = toRichText(workbook);
			res.applyFont(font.toFont(workbook));
			return res;
		};
	}

	default RichText applyFont(int start, int length, Font font) {
		return workbook -> {
			final var res = toRichText(workbook);
			res.applyFont(start, start + length, font.toFont(workbook));
			return res;
		};
	}

	default RichText bold(int start, int length) {
		return this.applyFont(start, length, Font.create().setBold(true));
	}

	default RichText bold() {
		return this.applyFont(Font.create().setBold(true));
	}

	default RichText color(int start, int length, IndexedColors colors) {
		return this.applyFont(start, length, Font.create().setColor(colors));
	}

	default RichText color(IndexedColors colors) {
		return this.applyFont(Font.create().setColor(colors));
	}
}
