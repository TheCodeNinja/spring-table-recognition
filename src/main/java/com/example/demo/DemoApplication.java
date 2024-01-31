package com.example.demo;

import com.spire.pdf.PdfDocument;
import com.spire.pdf.utilities.PdfTable;
import com.spire.pdf.utilities.PdfTableExtractor;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

// 说明：
//1、数据抽取的基础要求就是各字段内容互不影响，尤其是单元格内换行导致上线行内容错乱
//2、抽取字段间需要有明确的起始标记，此处以|分割字段值，以$分割行
//3、抽取时另需考虑值为空或非法格式等特殊场景

// 注意事项：
//1、spire.pdf不能读取所有pdf文本中的表格，部分场景直接报错（个人不涉及）
//2、spire.pdf依赖包模块化基础较差，单个包就达到50M，远超其他组件包
//3、spire.pdf只能识读PDF中的表格数据，表格外内容无法处理（图片OCR一般会兼容）
//4、spire.pdf免费版本维护周期较大，已知BUG有行尾有标点符号，则该标点符号会丢失，数据内容就会错乱，有些场景造成的问题很大

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(DemoApplication.class, args);
	}

}
