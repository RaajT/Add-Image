
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;

public class AddImage {
	public static HSSFWorkbook my_workbook = null;

	public static void main(String[] args) throws Exception {
		/* Write changes to the workbook */
		FileOutputStream out = new FileOutputStream(new File("C:\\Automation\\excel_insert_image_example.xls"));
		/* Create a Workbook and Worksheet */
		my_workbook = new HSSFWorkbook();
		HSSFSheet my_sheet = my_workbook.createSheet("MyBanner");
		/* Read the input image into InputStream */
		InputStream my_banner_image = new FileInputStream("C:\\Automation\\images\\correct16.png");
		/* Convert Image to byte array */
		byte[] bytes = IOUtils.toByteArray(my_banner_image);
		/* Add Picture to workbook and get a index for the picture */
		int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
		/* Close Input Stream */
		my_banner_image.close();
		/* Create the drawing container */
		HSSFPatriarch drawing;
		/* Create an anchor point */
		ClientAnchor my_anchor;
		HSSFPicture my_picture;
		/* Define top left corner, and we can resize picture suitable from there */
		for (int i = 1; i <= 10; i++) {

			drawing = my_sheet.createDrawingPatriarch();
			my_anchor = new HSSFClientAnchor();
			my_anchor.setCol1(i);
			my_anchor.setRow1(1);

			/* Invoke createPicture and pass the anchor point and ID */
			my_picture = drawing.createPicture(my_anchor, my_picture_id);
			/* Call resize method, which resizes the image */
			my_picture.resize();
			my_workbook.write(out);
		}

		
		out.close();
	}

}