import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Execute {

	static String inputFileName = "�x�Ѥ���20110103-20110225.xls";
	static String inputDirName = "/Users/Brian/Desktop/�ҵ{/���ݵ{�]/���W����k/�x�Ѥ���/�x��xls/";
	static String outputFileName = "�x�Ѥ���20110103-20110225.txt";
	static String outputDirName = "/Users/Brian/Desktop/�ҵ{/���ݵ{�]/���W����k/�x�Ѥ���/�x��txt/";
	static String savedDirName = "/Users/Brian/Desktop/�ҵ{/���ݵ{�]/���W����k/�x�Ѥ���/�x��xls_�w�ഫ/";
	static String[] valueAry = { "�}�L��", "�̰���", "�̧C��", "���L��", "����q" };
	static List<String> dateList = new ArrayList<String>();
	static Map<String, Integer> dateValueMap = new HashMap<String, Integer>();

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			File inputDir = new File(inputDirName);
			for (File inputFile : inputDir.listFiles()) {
				inputFileName = inputFile.getName();
				if (!inputFileName.endsWith(".xls"))
					continue;
				outputFileName = inputFileName + ".txt";
				Workbook workbook = Workbook.getWorkbook(inputFile);
				Sheet sheet = workbook.getSheet(0);
				initDataTable(sheet);

				BufferedWriter bufWriter = new BufferedWriter(new FileWriter(
						outputDirName + outputFileName));
				String printLine = "�Ѳ��N�X/�Ѳ��W��/����ɶ�/";
				for (int j = 0; j < valueAry.length; j++) {
					printLine += valueAry[j] + "/";
				}
				System.out.println(printLine);
				bufWriter.write(printLine);
				bufWriter.newLine();

				int i = 5;
				while (!("").equals(sheet.getCell(0, i).getContents())) {
					// 1.���o�C�ڢٴΪ��ɶ��P�U�ػ���
					for (String date : dateList) {
						boolean hasValue = true;
						printLine = "";
						// 2.���o�Ѳ��N�X�ΦW��
						String symbolID = sheet.getCell(0, i).getContents();
						String symbolName = sheet.getCell(1, i).getContents();
						printLine += symbolID + "/" + symbolName + "/";
						printLine += date + "/";
						// 2.1���o�},��,�C,��,����q
						for (int j = 0; j < valueAry.length; j++) {
							String date_valueType = date + valueAry[j];
							// 2.2���o�ӺؼƭȪ�column value
							int col = dateValueMap.get(date_valueType);
							String date_value = sheet.getCell(col, i)
									.getContents();
							if(("").equals(date_value)){
								hasValue = false;
								break;
							}
							printLine += date_value + "/";
						}
						// 2.3�p�G�Ӥ���S���ȡA�h���L
						if(hasValue == false)
							continue;
						System.out.println(printLine);
						bufWriter.write(printLine);
						bufWriter.newLine();
					}
					i++;
					if (sheet.getRows() <= i)
						break;
				}
				bufWriter.close();
				workbook.close();
				File newFile = new File(savedDirName + inputFileName);
				newFile.createNewFile();
				inputFile.renameTo(newFile);
			}
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void initDataTable(Sheet sheet) {
		dateList = new ArrayList<String>();
		dateValueMap = new HashMap<String, Integer>();
		int i = 2;
		String date = "";
		while (!("").equals((date = sheet.getCell(i, 1).getContents()))) {
			if (!dateList.contains(date))
				dateList.add(date);
			else
				break;
			i++;
			if (sheet.getColumns() <= i)
				break;
		}

		i = 2;
		date = "";
		while (!("").equals((date = sheet.getCell(i, 2).getContents()))) {
			if (!dateValueMap.containsKey(date))
				dateValueMap.put(date, i);
			i++;
			if (sheet.getColumns() <= i)
				break;
		}
	}
}
