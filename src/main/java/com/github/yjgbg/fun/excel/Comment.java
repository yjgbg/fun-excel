package com.github.yjgbg.fun.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

public interface Comment {
	default org.apache.poi.ss.usermodel.Comment toComment(Cell cell) {
		final var sheet = cell.getSheet();
		final var candidateDrawing = sheet.getDrawingPatriarch();
		final var drawing = candidateDrawing!=null ? candidateDrawing : sheet.createDrawingPatriarch();
		final var comment = drawing.createCellComment(new XSSFClientAnchor());
		comment.setAddress(cell.getRowIndex(), cell.getColumnIndex());
		post(comment, cell);
		return comment;
	}

	void post(org.apache.poi.ss.usermodel.Comment comment, Cell cell);

	static Comment simple(RichText text) {
		return (comment, cell) -> comment.setString(text.toRichText(cell.getSheet().getWorkbook()));
	}

	static Comment simple(String text) {
		return simple(RichText.of(text));
	}

	default Comment author(String author) {
		return (comment, cell) -> {
			this.post(comment,cell);
			comment.setAuthor(author);
		};
	}
}
