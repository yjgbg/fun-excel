package com.github.yjgbg.demo0;

import com.github.yjgbg.fun.excel.CellStyle;
import com.github.yjgbg.fun.excel.Font;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import static com.github.yjgbg.fun.excel.API.*;

@RestController
@RequiredArgsConstructor
public class SampleCtl {

	@SneakyThrows
	public static void main(String[] args) {
		final var workbook = Workbook(
				Sheet(
						Row(
								Cell("sheet1.row1.cell1").setCellStyle(CellStyle.bold()),
								Cell("sheet1.row1.cell2").setCellStyle(CellStyle.bold()),
								Cell("sheet1.row1.cell3").setCellStyle(CellStyle.bold())
						),
						Row(
								Cell("sheet1.row2.cell1"),
								Cell("sheet1.row2.cell2"),
								Cell("sheet1.row2.cell3")
						)
				),
				Sheet(
						Row(
								Cell("sheet2.row1.cell1"),
								Cell("sheet2.row1.cell2"),
								Cell("sheet2.row1.cell3")
						)
				)
		);
		final var now = LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMddHHmmss"));
		workbook.write(new FileOutputStream("C:\\Users\\12394\\Desktop\\test."+now+".xlsx"));
	}
}
