package com.github.yjgbg.fun.excel;

import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 主要是让String变身为RichText
 */
@UtilityClass
public class LbkExtFunExcel {
	public static RichText rich(String origin) {
		return RichText.of(origin);
	}
	public static RichText bold(String origin,int start,int length) {
		return RichText.of(origin).bold(start, length);
	}
	public static RichText bold(String origin) {
		return RichText.of(origin).bold();
	}
	public static RichText color(String origin, int start, int length, IndexedColors colors) {
		return RichText.of(origin).color(start, length, colors);
	}
	public static RichText color(String origin, IndexedColors colors) {
		return RichText.of(origin).color(colors);
	}
}
