package co.edu.ufps.convertidor;

import java.io.File;

import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;

/**
 * Hello world!
 *
 */
public class App {
	
	public static final String DOCX = "docx";
	public static final String ODT = "odt";
	public static final String XLSX = "xlsx";
	public static final String ODS = "ods";
	public static final String PPTX = "pptx";
	public static final String ODP = "odp";
	
	
	private Document document;
	private Workbook book;
	private Presentation presentation;

	public App() throws Exception {
		License lic = new License();
		lic.setLicense("Aspose.Total.Java.lic");
	}

	public static void main(String[] args) {
		String[] args1 = { "/home/miguel/Documentos/demo.odt"};

		try {
			App app;
			app = new App();
			app.convertir(args1);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void convertir(String[] args1) throws Exception {
		final int countArgumentos = args1.length;

		if (countArgumentos > 0) {
			final String ubicacion = args1[0];
			File file = new File(ubicacion);
			
			if(!file.exists()) {
				System.out.println("No existe el fichero:");
				System.out.println(ubicacion);
				return;
			}
			if(!file.isFile()) {
				System.out.println("No es un fichero:");
				System.out.println(ubicacion);
				return;
			}
			final String fileName = file.getName();
			final String directorio = file.getParent();

			// obtener la extension del archivo
			int i = fileName.lastIndexOf('.');
			if (i > 0) {
				// obtener nombre del archivo sin extencion

				final String fileNameNoExtension = fileName.substring(0, fileName.lastIndexOf('.'));
				final String extension = fileName.substring(i + 1);
				final String fileNameFinal = directorio+File.separator+fileNameNoExtension;
				
				System.out.println("ubicacion: "+ubicacion);
				System.out.println("fileNameFinal: "+fileNameFinal);
				System.out.println("extension: "+extension);
				
				switch (extension) {
				case DOCX:
					convertidorDocumento(ubicacion, fileNameFinal+"."+ODT, SaveFormat.ODT);
					break;
				case ODT:
					convertidorDocumento(ubicacion, fileNameFinal+"."+DOCX, SaveFormat.DOCX);
					break;
				case XLSX:
					convertidorHojaCalculo(ubicacion, fileNameFinal+"."+ODS, com.aspose.cells.SaveFormat.ODS);
					break;
				case ODS:
					convertidorHojaCalculo(ubicacion, fileNameFinal+"."+XLSX, com.aspose.cells.SaveFormat.XLSX);
					break;
				case PPTX:
					convertidorPresentacion(ubicacion, fileNameFinal+"."+ODP, com.aspose.slides.SaveFormat.Odp);
					break;
				case ODP:
					convertidorPresentacion(ubicacion, fileNameFinal+"."+PPTX, com.aspose.slides.SaveFormat.Pptx);
					break;

				default:
					System.out.println("No se reconoce el formato !!!");
					System.out.println("Abortando .. ");
				}
			}

		}else {
			System.out.println("Es requeria la ubicacion del archivo");
		}
	}

	private void convertidorDocumento(String fileName, String fileNameDestino, int formato) throws Exception {
		document = new Document(fileName);
		document.save(fileNameDestino, formato);
	}

	private void convertidorHojaCalculo(String fileName, String fileNameDestino, int formato) throws Exception {
		book = new Workbook(fileName);
		book.save(fileNameDestino, formato);

	}

	private void convertidorPresentacion(String fileName, String fileNameDestino, int formato) throws Exception {
		presentation = new Presentation(fileName);
		presentation.save(fileNameDestino, formato);

	}

}
