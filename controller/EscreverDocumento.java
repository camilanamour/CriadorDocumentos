package controller;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;;

public class EscreverDocumento {
	
	private XWPFDocument documento; // HWPFDocument (.doc);
	private FileOutputStream file;
	private XWPFHeaderFooterPolicy politica;
	
//	CRIAR DOCUMENTO
	public EscreverDocumento(File arquivo) throws FileNotFoundException{
		this.documento = new XWPFDocument(); 			// (1) Cria documento em branco.
		this.file = new FileOutputStream(arquivo); 		// (2) Criar arquivo de escrita.
		this.politica = new XWPFHeaderFooterPolicy(documento);  // (3) Verificar cabeçalho e rodapé
	}								// (4) Inserir o texto.
									// (5) Colocar o texto no documento. (GERAR ARQUIVO)
	
//      MUDAR A ORIENTAÇÃO
	public void mudarOrientacao(String orientacao){
		CTDocument1 doc = documento.getDocument();
	    CTBody corpo = doc.addNewBody();
	    corpo.addNewSectPr();
	    
	    CTSectPr sessao = corpo.getSectPr();
	    if(!sessao.isSetPgSz()) {
	    	sessao.addNewPgSz(); // Adiciona nova página.
	    }
	    
	    CTPageSz pagina = sessao.getPgSz();	// Pega a página. (A4 595x842)
	    if(orientacao.equalsIgnoreCase("paisagem")){
	    	pagina.setOrient(STPageOrientation.LANDSCAPE);
	    	pagina.setW(BigInteger.valueOf(842 * 20));
	    	pagina.setH(BigInteger.valueOf(595 * 20));
	    }
	    else{
	    	pagina.setOrient(STPageOrientation.PORTRAIT);
	    	pagina.setH(BigInteger.valueOf(842 * 20));
	    	pagina.setW(BigInteger.valueOf(595 * 20));
	    }
	}
	
//	CABEÇALHO
	public void inserirCabecalho(){
		XWPFHeader header = politica.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
		XWPFParagraph paragrafo = header.getParagraphArray(0); //ou apenas header.createParagraph();
		
		if(paragrafo == null) {
			paragrafo = header.createParagraph();
		}
		
		paragrafo.setAlignment(ParagraphAlignment.RIGHT);
		XWPFRun cabecalho = paragrafo.createRun();
		cabecalho.setText("Coloquei o cabeçalho!");		
	}
	
//	TÍTULO
	public void inserirTitulo(){
		XWPFParagraph titulo1 = documento.createParagraph();	//  Cria págrafo em branco.
		titulo1.setAlignment(ParagraphAlignment.CENTER);	// -> alinhamento no centro;
		XWPFRun titulo = titulo1.createRun();			// Criar o título do texto.
		this.EstiloTituloRun(titulo, "Título", "Showcard Gothic", "A9A9A9", 18);// Colocar estilo.
	}
	
	private void EstiloTituloRun(XWPFRun run, String texto, String fonte, String color, int tamanho){
		run.setBold(true);		// Negrito;
		run.setItalic(true);    	// Itálico;
		run.setText(texto);     	// Escrevo o texto;
		run.setFontFamily(fonte);	// Tipo de fonte;
		run.setFontSize(tamanho); 	// Tamanho da fonte;
		run.setColor(color); 		// Atribuir cor;
		run.setTextPosition(10);	// Espaçamento inferior;
		run.addBreak(); 		// Quebra linha.
	}
	
//	PARÁGRAFO
	public void inserirParagrafo(){
		XWPFParagraph paragrafo = documento.createParagraph(); 	// Criar parágrafo em branco.
		paragrafo.setAlignment(ParagraphAlignment.BOTH); 	// -> alinhamento justificado.
		XWPFRun corpoTexto = paragrafo.createRun();		// Alimenta o documento com um paragrafo
		
		corpoTexto.setFontFamily("Arial");
		corpoTexto.setFontSize(12);
		
		corpoTexto.addTab(); 					// -> adicionar espaçamento <tab>.
		corpoTexto.setText("Escrevendo um parágrafo...");	// -> USAR BUFFERSTRING converte para STRING!
	}
	
//	IMAGENS	
	public void inserirImagens(String path) throws InvalidFormatException, IOException{ 
		XWPFParagraph paragrafo = documento.createParagraph();
		paragrafo.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun imagem = paragrafo.createRun();
		
	 // Colocar imagem no documento.
		String imgFile = path;
		FileInputStream img = new FileInputStream(imgFile);
		imagem.addPicture(img, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(400), Units.toEMU(200)); // 400(larg)x200(altura) pixels
		img.close();
		
	}
	
//      RODAPÉ
	public void inserirRodape(){
		XWPFFooter footer = politica.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
		XWPFParagraph paragrafo = footer.createParagraph();
		paragrafo.setAlignment(ParagraphAlignment.RIGHT);
		XWPFRun rodape = paragrafo.createRun();
		rodape.setText("Coloquei o rodapé!");	
	}
	
//	TABELA
	public void inserirTabela(){
		XWPFTable tabela = documento.createTable();
		tabela.setWidth(500);
		
		// Primeira linha
		XWPFTableRow linha0 = tabela.getRow(0);
		linha0.getCell(0).setText("linha 0, coluna0");
		linha0.addNewTableCell().setText("linha 0, coluna1");
		
		//Segunda linha
		XWPFTableRow linha1 = tabela.createRow();
		linha1.getCell(0).setText("linha 1, coluna0");
		linha1.getCell(1).setText("linha 1, coluna1");
	}
	
//      GERAR O ARQUIVO.
	public void gerarArquivo() throws IOException{	
		documento.write(file);							
		file.close();
	}	
	
//      ABRIR O ARQUIVO	
	public void abrirArquivo(File arq) throws IOException{
		if(arq.exists() && arq.isFile()){
			Desktop desktop = Desktop.getDesktop();
			desktop.open(arq);
		} else {
			System.err.println("Arquivo inválido");
		}
	}

}
