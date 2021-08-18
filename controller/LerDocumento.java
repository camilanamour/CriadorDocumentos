package controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class LerDocumento {
	private XWPFDocument documento;
	private FileInputStream file;
	private XWPFWordExtractor extrator; 
	
	
	/* Para usar o extractor para (.doc):
	 * HWPFDocument documento = new HWPFDocument(fileCarregado);
	 * WordExtractor extrator = new WordExtractor(documento);
	 * String[] dadosArquivo = extrator.getParagraphText();	--> paga parágrafos
	 * String texto = extrator.getText();					--> pega tudo
	 * 
	 */

	public LerDocumento(File arquivo) throws InvalidFormatException, IOException {
		this.file = new FileInputStream(arquivo);
		documento = new XWPFDocument(OPCPackage.open(file));
	}

	// TEXTO COMPLETO - EXTRATOR.
	public void leituraTextoCompleto() throws Exception {
		extrator = new XWPFWordExtractor(documento);
		System.out.println(extrator.getText());
	}

	// CABEÇALHO E RODAPÉS - getDefaultHeader(); getDefaultFooter();
	public void leituraCabecalhoRodape() {
		XWPFHeaderFooterPolicy politica = new XWPFHeaderFooterPolicy(documento);

		XWPFHeader cabecalho = politica.getDefaultHeader();
		if (cabecalho != null) {
			System.out.println(cabecalho.getText());
		}
		XWPFFooter rodape = politica.getDefaultFooter();
		if (rodape != null) {
			System.out.println(rodape.getText());
		}
	}

	// PARÁGRAFOS - getParagraphs();
	public void leituraParagrafos() {
		List<XWPFParagraph> paragrafos = documento.getParagraphs();

		for (XWPFParagraph paragrafo : paragrafos) {
			if(paragrafo.getText().length() != 0){
				System.out.println("Texto: " + paragrafo.getText()); 
				System.out.println("Alinhamento: " + paragrafo.getAlignment()); 
				System.out.println("linhas: " + paragrafo.getRuns().size()); 
				System.out.println("Estilo: " + paragrafo.getStyle()); 
				System.out.println("Formato de Numeração: " + paragrafo.getNumFmt()); 
				System.out.println();
			}
		}
	}

	// TABELAS - XWPFTable
	public void leituraTabelas() {
		Iterator<IBodyElement> iterator = documento.getBodyElementsIterator(); 	// Pegar os elementos do corpo do texto
		
		// (1) Achar a tabela...
		while (iterator.hasNext()) { 										   	// Tem próximo?!
			IBodyElement elemento = (IBodyElement) iterator.next(); 			// Caminha;
			if ("TABLE".equalsIgnoreCase(elemento.getElementType().name())) {	// Nome do tipo do elemento é "TABLE"?
				List<XWPFTable> listaTabela = elemento.getBody().getTables();	// Pega a tabela.
				
				// (2) Tratamento com a tabela
				for (XWPFTable tabela : listaTabela) {
					System.out.println("Número total de linhas:" + tabela.getNumberOfRows());
					int linhasTotal = tabela.getRows().size();
									
					for (int i = 0; i < linhasTotal; i++) {
						for (int j = 0; j < tabela.getRow(i).getTableCells().size(); j++) {
							System.out.println(tabela.getRow(i).getCell(j).getText()); // Pegar linha(i) + coluna(j).
						}
					}
				}
			}
		}
	}

	// IMAGENS - XWPFPictureDatas
	public void leituraImagem() {
		List<XWPFPictureData> pic = documento.getAllPictures(); // Todas as imagens do arquivo
		if (!pic.isEmpty()) {									// Não está vazia?!
			System.out.println(pic.get(0).getPictureType());    // O tipo de imagem interna do POI, 0 se o tipo de imagem desconhecido.
			System.out.println(pic.get(0).getData());			// Obtém dados da imagem.
			// Dimension imgSize = getImageDimension(new ByteArrayInputStream(data.getData()), data.getPictureType());
		}
	}
}