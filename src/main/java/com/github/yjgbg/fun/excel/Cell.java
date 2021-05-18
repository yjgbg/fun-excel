package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.function.Consumer;

public interface Cell {
	org.apache.poi.ss.usermodel.Cell toCell(Row row);

	static Cell create() {
		return row -> {
			final var lastCellNum = row.getLastCellNum();
			return row.createCell(lastCellNum== -1 ? 0 : lastCellNum);
		};
	}

	static Cell create(int cellNum) {
		if (cellNum < 0) throw new IllegalArgumentException();
		return row -> {
			final var res = row.getCell(cellNum);
			return res != null ? res : row.createCell(cellNum);
		};
	}

	default Cell content(RichText text) {
		return post(res -> res.setCellValue(text.toRichText(res.getRow().getSheet().getWorkbook())));
	}


	default Cell content(LocalDateTime localDateTime) {
		return post(res -> res.setCellValue(localDateTime));
	}

	default Cell content(LocalDate localDate, String fmt) {
		return post(res -> res.setCellValue(localDate)).style((cellStyle, workbook) -> {
			final var index = workbook.createDataFormat().getFormat(fmt);
			cellStyle.setDataFormat(index);
		});
	}

	default Cell style(CellStyle... styles) {
		return post(res -> {
			final var workBook = res.getRow().getSheet().getWorkbook();
			final var style = Arrays.stream(styles).reduce(CellStyle.create(),CellStyle::plus);
			res.setCellStyle(style.toCellStyle(workBook));
		});
	}

	default Cell comment(Comment comment) {
		return post(res -> res.setCellComment(comment.toComment(res)));
	}

	private Cell post(Consumer<org.apache.poi.ss.usermodel.Cell> consumer) {
		return row -> {
			final var res = toCell(row);
			consumer.accept(res);
			return res;
		};
	}
}
