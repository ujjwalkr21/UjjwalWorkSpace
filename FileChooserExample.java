     
	import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Random;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class FileChooserExample {

	private static  String FILENAME = "E:\\Contact.vcf";
	static BufferedWriter bw = null;
	static FileWriter fw = null;
	
		public static void main(String[] args) throws BiffException, IOException {

			JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

			int returnValue = jfc.showOpenDialog(null);
			// int returnValue = jfc.showSaveDialog(null);

			if (returnValue == JFileChooser.APPROVE_OPTION) {
				File selectedFile = jfc.getSelectedFile();
				System.out.println(selectedFile.getAbsolutePath());
				
				Random rand = new Random();
				int value = rand.nextInt(1000000000);
				String FilePath = "F:\\Contact"+value+".xls";
				
				FileInputStream fs = new FileInputStream(selectedFile.getAbsolutePath());
				System.out.println("In readExcel()... Execution Starts Here..!! ");
				Workbook wb = Workbook.getWorkbook(fs);

				// TO get the access to the sheet
				Sheet sh = wb.getSheet("Sheet1");
				
				// To get the number of rows present in sheet
				int totalNoOfRows = sh.getRows();
				System.out.println(" totalNoOfRows:- "+totalNoOfRows);
				// To get the number of columns present in sheet
				int totalNoOfCols = sh.getColumns();
				System.out.println(" totalNoOfCols:- "+totalNoOfCols+"\n----------------------------");
				
				try {

					String content = "This is the content to write into file\n";
					fw = new FileWriter(FILENAME);
					bw = new BufferedWriter(fw);
					//fw.write(content);
					
					for (int row = 0; row < totalNoOfRows; row++) {

						for (int col = 0; col < totalNoOfCols-2; col++) {
							/*
							System.out.print("BEGIN:VCARD\nVERSION:3.0\nN:;");
							System.out.print(sh.getCell(1, row).getContents());	//Print first Column
							System.out.print(";;;\nFN:");
							System.out.print(sh.getCell(1, row).getContents() + "\n"); //Print first Column
							System.out.print("UID:");///TEL;TYPE=PREF,mobile:
							System.out.print(sh.getCell(0, row).getContents());
							System.out.print("\nTEL;TYPE=PREF,mobile:");
							System.out.print(sh.getCell(2, row).getContents());
							System.out.print("\nEND:VCARD");	
							*/
							//Perform Same thing for.vcf file
							fw.write("BEGIN:VCARD\nVERSION:3.0\nN:;");
							fw.write(sh.getCell(1, row).getContents());	//Print first Column
							fw.write(";;;\nFN:");
							fw.write(sh.getCell(1, row).getContents() + "\n"); //Print first Column
							fw.write("UID:");///TEL;TYPE=PREF,mobile:
							fw.write(sh.getCell(0, row).getContents());
							fw.write("\nTEL;TYPE=PREF,mobile:");
							fw.write(sh.getCell(2, row).getContents());
							fw.write("\nEND:VCARD");
						}
						fw.write("\n\n");
						//System.out.println("\n");
						if(row == totalNoOfRows-1){
							System.out.println("VCF Created Succesfully..!!");
						}
					}
					

				} catch (IOException e) {

					e.printStackTrace();
			}
				 finally {

						try {

							if (bw != null)
								bw.close();

							if (fw != null)
								fw.close();

						} catch (IOException ex) {

							ex.printStackTrace();

						}

					}
				
			}

		}

	}