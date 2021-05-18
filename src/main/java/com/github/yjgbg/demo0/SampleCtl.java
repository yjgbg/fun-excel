package com.github.yjgbg.demo0;

import com.github.yjgbg.fun.excel.Comment;
import com.github.yjgbg.fun.excel.LbkExtFunExcel;
import lombok.SneakyThrows;
import lombok.experimental.ExtensionMethod;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import static com.github.yjgbg.fun.excel.API.*;
import static com.github.yjgbg.fun.excel.CellStyle.*;
import static org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER;
import static org.apache.poi.ss.usermodel.IndexedColors.RED;

@ExtensionMethod({LbkExtFunExcel.class})
public class SampleCtl {

	@SneakyThrows
	public static void main(String[] args) {
		final var workbook = Workbook(
				Sheet("表1",
						Row(
								Cell("sheet1.row1.cell1".rich().bold(0,6).color(0,6,RED)),
								Cell("sheet1\n.row1\n.cell2")
										.style(color(RED),wrapText(),alignment(CENTER)),
								Cell("sheet1.row1.cell3"),
								Cell(LocalDate.now(),"yyyyMMdd")
						),
						Row(
								Cell("sheet1.row2.cell1"),
								Cell("sheet1.row2.cell2"),
								Cell("sheet1.row2.cell3")
										.style(color(RED))
										.comment(Comment.simple("这是个批注").author("weicl"))
						)
				),
				Sheet("表2"),
				Sheet("表3",
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
