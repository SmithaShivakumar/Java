import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;



public class ExcelMultiRead {

	/**
	 * @param args
	 *            the command line arguments
	 */
	public static void main(String[] args) {

		String FILE_NAME = "PJM_V4.xlsx";
		String targetDir = "out";
		String zipDir = "out";

		File file = new File(FILE_NAME);
		if (file.exists()) {
			Convertfile(file, targetDir, zipDir);
		}

	}

	public static void Convertfile(File inputFile, String targetDir, String zipDir) {
		InputStream inp = null;
		byte[] buffer = new byte[1024];
		try {
			inp = new FileInputStream(inputFile);
			Workbook wb = WorkbookFactory.create(inp);
			String filename = getFileName(inputFile);

			FileOutputStream zipOut = new FileOutputStream(zipDir + File.separatorChar + filename + ".zip");

			ZipOutputStream zos = new ZipOutputStream(zipOut);
			String CSV_FILE = wb.getSheetAt(7).getSheetName();
			System.out.println(CSV_FILE) ;
			File fout = new File(targetDir + File.separatorChar + "Zones.csv");

			FileOutputStream fos = new FileOutputStream(fout);
			OutputStreamWriter osw = new OutputStreamWriter(fos);

			echoAsCSV(wb.getSheetAt(7), osw);

			osw.close();
			ZipEntry ze = new ZipEntry(fout.getName());
			zos.putNextEntry(ze);
			FileInputStream in = new FileInputStream(fout.getAbsolutePath());

			int len;
			while ((len = in.read(buffer)) > 0) {
				zos.write(buffer, 0, len);
			}

			in.close();
			zos.closeEntry();


			zos.close();

		} catch (InvalidFormatException ex) {
			Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		} catch (FileNotFoundException ex) {
			Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		} catch (IOException ex) {
			Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
		} finally {
			try {
				inp.close();
			} catch (IOException ex) {
				Logger.getLogger(ExcelMultiRead.class.getName()).log(Level.SEVERE, null, ex);
			}
		}

	}

	public static void echoAsCSV(Sheet sheet, OutputStreamWriter osw) throws IOException {

		Row row = null;
		String header = "ZoneID, Description, Reserve(%load)";
		osw.write(header + "\n");

		for (int rowIndex = 1	; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			row = sheet.getRow(rowIndex);

			if (row != null) {
				Cell mappedLoadZoneID = row.getCell(6);
				if (mappedLoadZoneID.getNumericCellValue() != 0.0) {
					Cell zoneName = row.getCell(1) ;
					if (notSeenBefore(mappedLoadZoneID, zoneName )) {
						osw.write(zoneName + ", " + mappedLoadZoneID + ", 0.0\n");
						seenCellNow(mappedLoadZoneID, zoneName);
					}
				}
			}
		}
		osw.write("\n");
	}

	static class ZoneNameZoneIDPair {
		String zoneName;
		double mappedZoneID;

		@Override
		public boolean equals(Object o) {
			if (this == o) return true;
			if (o == null || getClass() != o.getClass()) return false;

			ZoneNameZoneIDPair that = (ZoneNameZoneIDPair) o;

			if (Double.compare(that.mappedZoneID, mappedZoneID) != 0) return false;
			return zoneName != null ? zoneName.equals(that.zoneName) : that.zoneName == null;
		}

		@Override
		public int hashCode() {
			int result;
			long temp;
			result = zoneName != null ? zoneName.hashCode() : 0;
			temp = Double.doubleToLongBits(mappedZoneID);
			result = 31 * result + (int) (temp ^ (temp >>> 32));
			return result;
		}
	}
	static Set seenZoneName = new HashSet<ZoneNameZoneIDPair>();

	static boolean notSeenBefore(Cell mappedLoadZoneID, Cell zoneName) {
		ZoneNameZoneIDPair zoneNamePair = new ZoneNameZoneIDPair();
		zoneNamePair.zoneName = zoneName.getStringCellValue();
		zoneNamePair.mappedZoneID = mappedLoadZoneID.getNumericCellValue();
		return !seenZoneName.contains(zoneNamePair);
	}

	static void seenCellNow(Cell mappedLoadZoneID, Cell zoneName) {
		ZoneNameZoneIDPair zoneNamePair = new ZoneNameZoneIDPair();
		zoneNamePair.zoneName = zoneName.getStringCellValue();
		zoneNamePair.mappedZoneID = mappedLoadZoneID.getNumericCellValue();
		seenZoneName.add(zoneNamePair);
	}


	private static String getFileName(File file) {
		String fileName = file.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(0, fileName.lastIndexOf("."));
		else
			return "";
	}

	private static String getFileExtension(File file) {
		String fileName = file.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

}
