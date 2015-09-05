import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelVeriOkuma {

	public int sayfaSecme(String sayfa) {
		if (sayfa.equals("MW TEST01"))

			return 0;
		if (sayfa.equals("MW TEST02"))

			return 1;
		if (sayfa.equals("MW TEST03"))

			return 2;
		return 0;
	}

	public String veriOkuma(int sayfaBelirteci, String hatIsmi) {
		try {
			List<String> liste = new ArrayList<String>();
			int uygunKayitSayisi = 0;
			// Excel Dosya Yolu
			FileInputStream file = new FileInputStream(new File("ssss.xltx"));

			// Excel Dosyamizi Temsil Eden Workbook Nesnesi
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Excel Dosyasýnýn Hangi Sayfasý Ýle Çalýþacaðýmýzý Seçelim.
			XSSFSheet sheet = workbook.getSheetAt(sayfaBelirteci);

			// Belirledigimiz sayfa icerisinde tum satirlari tek tek dolasacak
			// iterator nesnesi
			Iterator<Row> rowIterator = sheet.iterator();
			Cell cell;
			// Okunacak Satir Oldugu Surece
			while (rowIterator.hasNext()) {

				String[] tempArray = new String[9];
				// Excel içerisindeki satiri temsil eden nesne
				Row row = rowIterator.next();

				// Her bir satir icin tum hucreleri dolasacak iterator nesnesi
				Iterator<Cell> cellIterator = row.cellIterator();
				int counter = 0;
				while (cellIterator.hasNext()) {

					// Excel icerisindeki hucreyi temsil eden nesne
					cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						tempArray[counter] = String.valueOf(cell
								.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						tempArray[counter] = cell.getStringCellValue()
								.toString();
						break;
					}

					counter++;
				}
				if (tempArray[0].equals(hatIsmi)
						&& tempArray[7].equals("Aktif")) {
					uygunKayitSayisi++;
					for (int i = 0; i < 9; i++) {
						liste.add(tempArray[i]);
					}
				}

			}
			Random random = new Random();

			boolean sayiSecme = true;
			int rastgeleSayi = 0;
			while (sayiSecme == true) {
				rastgeleSayi = random.nextInt(uygunKayitSayisi - 1);
				if (rastgeleSayi != 0) {
					System.out.println(rastgeleSayi);
					sayiSecme = false;
				}
			}

			String deger = "";
			for (int j = (rastgeleSayi - 1) * 9; j < ((rastgeleSayi - 1) * 9) + 9; j++) {
				deger = deger + liste.get(j) + " ";
			}

			file.close();

			return deger;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException er) {
			er.printStackTrace();
		}
		return "hata";

	}

	public void veriEslestirme(int sayfaBelirleyici, String veri) {
		try {
			FileInputStream file = new FileInputStream(new File("ssss.xltx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			XSSFSheet sheet = workbook.getSheetAt(sayfaBelirleyici);

			Iterator<Row> rowIterator = sheet.iterator();
			Cell cell;
			int satýrSayici = 1;
			String tempVeri = "";
			while (rowIterator.hasNext()) {

				Row row = rowIterator.next();
				tempVeri = "";
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					cell = cellIterator.next();
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						tempVeri = tempVeri + " "
								+ String.valueOf(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						tempVeri = tempVeri + " " + cell.getStringCellValue();
						break;
					}

				}

				if (tempVeri.trim().equals(veri.trim())) {
					System.out.println(tempVeri);
					System.out.println(veri);
					System.out.println(satýrSayici);
					break;
				}
				satýrSayici++;
			}

			file.close();
			veriYaz(satýrSayici, sayfaBelirleyici);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException er) {
			er.printStackTrace();
		}
	}

	public static void veriYaz(int satir, int sayfaBelirleyici) {
		satir = satir - 1;
		try {
			FileInputStream file = new FileInputStream(new File("ssss.xltx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(sayfaBelirleyici);
			Cell cell = sheet.getRow(satir).getCell(7);
			cell.setCellValue("Pasif");
			FileOutputStream outFile = new FileOutputStream(new File(
					"ssss.xltx"));
			workbook.write(outFile);
			outFile.close();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
