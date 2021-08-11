package org.cp;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
public class Test
{

	public static void main(String[] args) throws IOException
	{
		String path = "C:\\Users\\goodl\\Desktop\\abc\\报表\\FO_02012_门店毛利日报\\";
		String a = "良品_FO_02012_门店毛利日报_20201125.xlsx";
		String b = "云徙_FO_02012_门店毛利日报_20201125.xlsx";
		int keyLength = 10;

		String[] keys = getKeys(path + a, "", keyLength);

		List<Map<String, String>> xx = readExcel(new FileInputStream(path + a), "");

		List<Map<String, String>> xx2 = readExcel(new FileInputStream(path + b), "");
		System.out.println("找出数据不匹配");
		writeNotIn(keys, xx, xx2, path + "良品NotIn云徙.xlsx", path + "云徙NotIn良品.xlsx");

		System.out.println("对比数据差异");
		whriteSame(keys, xx, xx2, "良品", "云徙", path + "良品In云徙.xlsx");
		System.out.println("处理完成");

	}

	private static void writeNotIn(String[] keys, List<Map<String, String>> xx, List<Map<String, String>> xx2, String aNotInBFile, String bNotInAFile)
	{
		List<Map<String, String>> aNotInB = new ArrayList<>();
		for (Map<String, String> row : xx)
		{
			String[] values = getValues(keys, row);

			int index = getDataIndex(keys, values, xx2);
			if (index == -1)
			{
				aNotInB.add(row);
			}
		}
		List<Map<String, String>> bNotInA = new ArrayList<>();
		for (Map<String, String> row : xx2)
		{
			String[] values = getValues(keys, row);

			int index = getDataIndex(keys, values, xx);
			if (index == -1)
			{
				bNotInA.add(row);

			}
		}

		if (!aNotInB.isEmpty()) writeExcel(aNotInB, aNotInBFile);
		if (!bNotInA.isEmpty()) writeExcel(bNotInA, bNotInAFile);
	}
	private static void whriteSame(String[] keys, List<Map<String, String>> xx, List<Map<String, String>> xx2, String a, String b, String saveFile)
	{
		List<Map<String, String>> aInB = new ArrayList<>();
		for (Map<String, String> row : xx)
		{
			String[] values = getValues(keys, row);

			int index = getDataIndex(keys, values, xx2);
			if (index != -1)
			{
				aInB.add(getNewRow(row, xx2.get(index), keys, a, b));
			}
		}

		if (!aInB.isEmpty()) writeExcel(aInB, saveFile);

	}

	public static void print(Map<String, String> row)
	{
		for (Map.Entry<String, String> entry : row.entrySet())
		{
			String mapKey = entry.getKey();
			String mapValue = entry.getValue();
			System.out.println(mapKey + ":" + mapValue);
		}
	}
	/**
	 * 用key值找数据
	 * 
	 * @param keys
	 * @param values
	 * @param data
	 * @return
	 */
	public static int getDataIndex(String[] keys, String[] values, List<Map<String, String>> datas)
	{

		for (int i = 0, size = datas.size(); i < size; i++)
		{
			Map<String, String> data = datas.get(i);
			if (sameKey(keys, values, data)) return i;
		}
		return -1;
	}

	/**
	 * 获得keys对应值
	 */
	public static String[] getValues(String[] keys, Map<String, String> data)
	{

		int length = keys.length;
		String[] value = new String[length];
		for (int i = 0; i < length; i++)
		{
			value[i] = data.get(keys[i]);

		}
		return value;
	}

	/**
	 * 判断数据是否是这个key
	 */
	private static boolean sameKey(String[] keys, String[] values, Map<String, String> data)
	{

		for (int i = 0, length = keys.length; i < length; i++)
		{
			String v = data.get(keys[i]);
			if (values[i] == v || values[i].equals(v)) continue;
			return false;
		}
		return true;
	}

	public static String[] getKeys(String fileName, String sheetName, int length)
	{
		// 定义工作簿
		XSSFWorkbook xssfWorkbook = null;
		try
		{
			xssfWorkbook = new XSSFWorkbook(new FileInputStream(fileName));
		}
		catch (Exception e)
		{
			System.out.println("Excel data file cannot be found!");
		}

		// 定义工作表
		XSSFSheet xssfSheet;
		if (sheetName.equals(""))
		{
			// 默认取第一个子表
			xssfSheet = xssfWorkbook.getSheetAt(0);
		}
		else
		{
			xssfSheet = xssfWorkbook.getSheet(sheetName);
		}

		String[] args = new String[length];

		// 定义行
		// 默认第一行为标题行，index = 0
		XSSFRow titleRow = xssfSheet.getRow(0);

		// 循环取每个单元格(cell)的数据
		for (int cellIndex = 0; cellIndex < length; cellIndex++)
		{
			XSSFCell titleCell = titleRow.getCell(cellIndex);
			args[cellIndex] = getCellValue(titleCell);
		}
		return args;

	}
	/**
	 * 把内容写入Excel
	 * 
	 * @param list
	 *            传入要写的内容，此处以一个List内容为例，先把要写的内容放到一个list中
	 * @param outputStream
	 *            把输出流怼到要写入的Excel上，准备往里面写数据
	 */
	public static void writeExcel(List<Map<String, String>> list, String outputName)
	{
		// 创建工作簿

		XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

		// 创建工作表
		XSSFSheet xssfSheet = xssfWorkbook.createSheet();
		// 创建行
		XSSFRow xssfRow;

		// 创建列，即单元格Cell
		XSSFCell xssfCell;
		// 创建第一行
		Map<String, String> title = list.get(0);
		String[] keys = title.keySet().toArray(new String[0]);
		int c = title.size();
		xssfRow = xssfSheet.createRow(0);
		for (int j = 0; j < c; j++)
		{
			xssfCell = xssfRow.createCell(j); // 创建单元格
			if (keys[j].endsWith("_差异--"))
				xssfCell.setCellValue("差异"); // 设置单元格内容
			else
				xssfCell.setCellValue(keys[j]); // 设置单元格内容
		}

		XSSFCellStyle isInteger = xssfWorkbook.createCellStyle();

		isInteger.setDataFormat(xssfWorkbook.createDataFormat().getFormat("##0"));// 数据格式只显示整数

		XSSFCellStyle doubleNum = xssfWorkbook.createCellStyle();

		doubleNum.setDataFormat(xssfWorkbook.createDataFormat().getFormat("###0.####"));// 4位小数

		XSSFCellStyle general = xssfWorkbook.createCellStyle();

		general.setDataFormat(xssfWorkbook.createDataFormat().getFormat("General"));// 数据格式只显示整数

		// 把List里面的数据写到excel中
		for (int i = 0; i < list.size(); i++)
		{
			// 从第一行开始写入
			xssfRow = xssfSheet.createRow(i + 1);
			// 创建每个单元格Cell，即列的数据

			Map<String, String> data = list.get(i);
			for (int j = 0; j < c; j++)
			{
				xssfCell = xssfRow.createCell(j); // 创建单元格

				String value = data.get(keys[j]);
				xssfCell.setCellValue(value);
				// setCellValue(value, xssfCell, isInteger, doubleNum,general);

			}
		}

		// 用输出流写到excel
		try
		{
			xssfWorkbook.write(new FileOutputStream(outputName));
			xssfWorkbook.close();

		}
		catch (IOException e)
		{
			e.printStackTrace();

		}

	}

	private static void setCellValue(String data, Cell contentCell, XSSFCellStyle i, XSSFCellStyle d, XSSFCellStyle c)
	{

		Boolean isNum = false;// data是否为数值型
		Boolean isInteger = false;// data是否为整数
		Boolean isPercent = false;// data是否为百分数
		if (data != null || "".equals(data))
		{
			// 判断data是否为数值型
			isNum = data.toString().matches("^(-?\\d+)(\\.\\d+)?$");
			// 判断data是否为整数（小数部分是否为0）
			isInteger = data.toString().matches("^[-\\+]?[\\d]*$");
			// 判断data是否为百分数（是否包含“%”）
			isPercent = data.toString().contains("%");
		}

		// 如果单元格内容是数值类型，涉及到金钱（金额、本、利），则设置cell的类型为数值型，设置data的类型为数值类型
		if (isNum && !isPercent)
		{
			contentCell.setCellType(CellType.NUMERIC);
			// contentCell.setCellStyle(c);
			contentCell.setCellValue(data.toString());
			// if (isInteger)
			// {
			// contentCell.setCellStyle(i);
			//
			// }
			// else
			// {
			// contentCell.setCellStyle(d);
			// }
			// // 设置单元格格式
			//
			// // 设置单元格内容为double类型
			// contentCell.setCellValue(Double.parseDouble(data.toString()));
		}
		else
		{

			// 设置单元格内容为字符型
			contentCell.setCellValue(data.toString());
		}
	}
	/**
	 * 读取Excel文件的内容
	 * 
	 * @param inputStream
	 *            excel文件，以InputStream的形式传入
	 * @param sheetName
	 *            sheet名字
	 * @return 以List返回excel中内容
	 */
	public static List<Map<String, String>> readExcel(InputStream inputStream, String sheetName)
	{

		// 定义工作簿
		XSSFWorkbook xssfWorkbook = null;
		try
		{
			xssfWorkbook = new XSSFWorkbook(inputStream);
		}
		catch (Exception e)
		{
			System.out.println("Excel data file cannot be found!");
		}

		// 定义工作表
		XSSFSheet xssfSheet;
		if (sheetName.equals(""))
		{
			// 默认取第一个子表
			xssfSheet = xssfWorkbook.getSheetAt(0);
		}
		else
		{
			xssfSheet = xssfWorkbook.getSheet(sheetName);
		}

		List<Map<String, String>> list = new ArrayList<Map<String, String>>();

		// 定义行
		// 默认第一行为标题行，index = 0
		XSSFRow titleRow = xssfSheet.getRow(0);

		// 循环取每行的数据
		for (int rowIndex = 1; rowIndex < xssfSheet.getPhysicalNumberOfRows(); rowIndex++)
		{
			XSSFRow xssfRow = xssfSheet.getRow(rowIndex);
			if (xssfRow == null)
			{
				continue;
			}

			Map<String, String> map = new LinkedHashMap<String, String>();
			// 循环取每个单元格(cell)的数据
			for (int cellIndex = 0; cellIndex < xssfRow.getPhysicalNumberOfCells(); cellIndex++)
			{
				XSSFCell titleCell = titleRow.getCell(cellIndex);
				XSSFCell xssfCell = xssfRow.getCell(cellIndex);
				map.put(getCellValue(titleCell), getCellValue(xssfCell));
			}
			list.add(map);
		}
		return list;
	}

	public static String getCellValue(Cell cell)
	{

		if (cell.getCellType() == CellType.NUMERIC)
		{
			String value = String.valueOf(cell.getNumericCellValue());
			if (value.indexOf('E') != -1)
			{
				cell.setCellType(CellType.STRING);
				return cell.getStringCellValue();
			}
			return value.endsWith(".0") ? value.substring(0, value.length() - 2) : value;
		}

		if (cell.getCellType() == CellType.BOOLEAN) return String.valueOf(cell.getBooleanCellValue());

		if (cell.getCellType() == CellType.STRING) return String.valueOf(cell.getStringCellValue());

		if (cell.getCellType() == CellType.FORMULA)

			return String.valueOf(cell.getCellFormula());

		if (cell.getCellType() == CellType.BLANK)

			return "";

		return "ERROR";
	}

	public static Map<String, String> getNewRow(Map<String, String> arow, Map<String, String> brow, String[] keys, String a, String b)
	{
		Map<String, String> map = new LinkedHashMap<String, String>();

		// 先写主键
		for (String key : keys)
		{
			map.put(key, arow.get(key));
		}
		for (String key : arow.keySet())
		{
			if (map.containsKey(key)) continue;
			String avalue = arow.get(key);
			map.put(a + "_" + key, avalue);
			// 写b表数据
			String bvalue = brow.get(key);
			map.put(b + "_" + key, bvalue);
			// 写差异
			map.put(key + "_差异--", getDif(avalue, bvalue));
		}
		// 补下b表
		for (String key : brow.keySet())
		{
			if (map.containsKey(key) || map.containsKey(b + "_" + key)) continue;

			String bvalue = brow.get(key);

			map.put(b + "_" + key, bvalue);

		}
		return map;
	}
	public static String getDif(String avalue, String bvalue)
	{
		if (avalue == bvalue || avalue.equals(bvalue)) return "0";
		if (isInteger(avalue) && isInteger(bvalue))
		{
			return String.valueOf(sub(Double.parseDouble(avalue), Double.parseDouble(bvalue)));
		}
		return "1";
	}

	public static Double sub(Double v1, Double v2)
	{
		BigDecimal b1 = new BigDecimal(v1.toString());
		BigDecimal b2 = new BigDecimal(v2.toString());
		return b1.subtract(b2).doubleValue();
	}
	public static boolean isInteger(String str)
	{
		if (str == null) return false;
		return str.matches("[+-]?[0-9]+(\\.[0-9]{1,4})?");
	}
}
