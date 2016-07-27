
package shujuchuli;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 操作Excel表格的功能类
 */
public class Jiexi {
	private POIFSFileSystem fs;
	private HSSFWorkbook wb;
	private XSSFSheet sheet;
	private XSSFRow row;
	private String fm;

	/**
	 * 读取Excel表格表头的内容
	 * 
	 * @param InputStream
	 * @return String 表头内容的数组
	 */
	public String[] readExcelTitle(File file) {
		// HSSFWorkbook workBook = new HSSFWorkbook();// 创建 一个excel文档对象
		XSSFWorkbook wb = null;
		try {
			System.out.println(file.getAbsolutePath());
			System.out.println(file.getAbsolutePath());
			wb = new XSSFWorkbook(new FileInputStream(file));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		XSSFSheet sheet = wb.createSheet();// 创建一个工作薄对象
		// fs = new POIFSFileSystem(is);
		// wb = new HSSFWorkbook(fs);
		System.out.println("____________________");
		sheet = wb.getSheetAt(0);// 得到Excle工作表的行
		row = sheet.getRow(0);// 得到Excle工作表指定行的单元格
		System.out.println("++++++++++++++++++++++");
		// 标题总列数
		int colNum = row.getPhysicalNumberOfCells();
		System.out.println("colNum:" + colNum);
		String[] title = new String[colNum];
		for (int i = 0; i < colNum; i++) {
			// title[i] = getStringCellValue(row.getCell((short) i));
			title[i] = getCellFormatValue(row.getCell((short) i));
		}
		return title;
	}

	/**
	 * 读取Excel数据内容
	 * 
	 * @param InputStream
	 * @return Map 包含单元格数据内容的Map对象
	 */
	public Map<Integer, String> readExcelContent(File file) {
		Map<Integer, String> content = new HashMap<Integer, String>();
		String str = " ";
		/*
		 * try { fs = new POIFSFileSystem(is);//
		 * POIFSFileSystem是apache提供的poi支持包里的类，专门用来解析Excel的。得到Excel工作薄对象 wb = new
		 * HSSFWorkbook(fs);// 得到Excel工作表对象 } catch (IOException e) {
		 * e.printStackTrace(); }
		 */
		// HSSFWorkbook workBook = new HSSFWorkbook();// 创建 一个excel文档对象
		XSSFWorkbook wb = null;
		try {
			System.out.println(file.getAbsolutePath());
			wb = new XSSFWorkbook(new FileInputStream(file));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		sheet = wb.getSheetAt(0);
		// 得到总行数
		int rowNum = sheet.getLastRowNum();
		row = sheet.getRow(0);
		System.out.println(file.getAbsolutePath());
		int colNum = row.getPhysicalNumberOfCells();
		//System.out.println(colNum);
		// 正文内容应该从第二行开始,第一行为表头的标题
		for (int i = 1; i <rowNum-1; i++) {
			row = sheet.getRow(i);
			int j = 0;
			while (j < colNum) {
				// 每个单元格的数据内容用"-"分割开，以后需要时用String类的replace()方法还原数据
				// 也可以将每个单元格的数据设置到一个javabean的属性中，此时需要新建一个javabean
				// str += getStringCellValue(row.getCell((short) j)).trim() +
				// "-";
				//System.out.println(j);
				str += getCellFormatValue(row.getCell(j)).trim();
				System.out.println(str);
				j++;
			}
			content.put(i, str);
			str = "";
		}
		return content;
	}

	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	private String getStringCellValue(HSSFCell cell) {
		String strCell = "";
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			strCell = String.valueOf(cell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			strCell = "";
			break;
		default:
			strCell = "";
			break;
		}
		if (strCell.equals("") || strCell == null) {
			return "";
		}
		if (cell == null) {
			return "";
		}
		return strCell;
	}

	/**
	 * 获取单元格数据内容为日期类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	private String getDateCellValue(HSSFCell cell) {
		String result = "";
		try {
			int cellType = cell.getCellType();
			if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
				Date date = cell.getDateCellValue();
				result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1) + "-" + date.getDate();
			} else if (cellType == HSSFCell.CELL_TYPE_STRING) {
				String date = getStringCellValue(cell);
				result = date.replaceAll("[年月]", "-").replace("日", "").trim();
			} else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
				result = "";
			}
		} catch (Exception e) {
			System.out.println("日期格式不正确!");
			e.printStackTrace();
		}
		return result;
	}

	public class GetFileName {
		public String[] fm(String path) {
			File file = new File(path);
			String[] fileName = file.list();
			return fileName;
		}

		public void sl(String path, Set<File> fileSet) {
			File file = new File(path);
			File[] files = file.listFiles();
			String[] names = file.list();
			if (names != null)
				for (File a : files) {
					// System.out.println(a.getAbsolutePath());
					fileSet.add(a);
					if (a.isDirectory()) {
						sl(a.getAbsolutePath(), fileSet);
					}
				}
		}
	}

	/**
	 * 根据HSSFCell类型设置数据
	 * 
	 * @param xssfCell
	 * @return
	 */
	private String getCellFormatValue(XSSFCell xssfCell) {
		String cellvalue = "";
		if (xssfCell != null) {
			// 判断当前Cell的Type
			switch (xssfCell.getCellType()) {
			// 如果当前Cell的Type为NUMERIC
			case HSSFCell.CELL_TYPE_NUMERIC:
			case HSSFCell.CELL_TYPE_FORMULA: {
				// 判断当前的cell是否为Date
				if (HSSFDateUtil.isCellDateFormatted(xssfCell)) {
					// 如果是Date类型则，转化为Data格式

					// 方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
					// cellvalue = cell.getDateCellValue().toLocaleString();

					// 方法2：这样子的data格式是不带带时分秒的：2011-10-12
					Date date = xssfCell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellvalue = sdf.format(date);

				}
				// 如果是纯数字
				else {
					// 取得当前Cell的数值
					cellvalue = String.valueOf(xssfCell.getNumericCellValue());
				}
				break;
			}
				// 如果当前Cell的Type为STRIN
			case HSSFCell.CELL_TYPE_STRING:
				// 取得当前的Cell字符串
				cellvalue = xssfCell.getRichStringCellValue().getString();
				break;
			// 默认的Cell值
			default:
				cellvalue = " ";
			}
		} else {
			cellvalue = "";
		}
		return cellvalue;

	}

	public static void main(String[] args) {
		GetFileName gn = new Jiexi().new GetFileName();
		String[] fileName = gn.fm("/Users/apple/Desktop/原始数据/");
		for (String name : fileName) {
			System.out.println(name);
		}
		System.out.println("--------------------------------");
		ArrayList<String> listFileName = new ArrayList<String>();
		Set<File> fileSet = new LinkedHashSet();
		gn.sl("/Users/apple/Desktop/原始数据/", fileSet);
		for (File file1 : fileSet) {
			if (file1.getAbsolutePath().endsWith("xlsx")) {
				System.out.println(file1.getAbsolutePath());
				try {
					FileInputStream is = new FileInputStream(file1.getAbsolutePath());
					Jiexi excelReader = new Jiexi();
					String[] title = excelReader.readExcelTitle(file1);
					System.out.println("获得Excel表格的标题:");
					for (String s : title) {
						System.out.print(s + " ");
					}

					// 对读取Excel表格内容测试
					// FileInputStream is2 = new
					// FileInputStream(file1.getAbsolutePath());
					Map<Integer, String> map = excelReader.readExcelContent(file1);
					System.out.println("获得Excel表格的内容:");
					for (int i = 1; i <= map.size(); i++) {
						System.out.println(map.get(i));
					}

				} catch (FileNotFoundException e) {
					System.out.println("未找到指定路径的文件!");
					e.printStackTrace();
				}
				// end of catch

				// }
				// }
			}
		}
	}
}
