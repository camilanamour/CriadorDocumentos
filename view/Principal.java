package view;

import java.io.File;

import javax.swing.JOptionPane;

import controller.EscreverDocumento;
import controller.LerDocumento;

public class Principal {

	public static void main(String[] args) throws Exception {
		
		File arquivo = new File("documento.docx"); //Path documento.
		EscreverDocumento gerar = null;
		LerDocumento exibir = null;
		
		try {
			JOptionPane.showMessageDialog(null, escreveDocumentos(gerar, arquivo));
			JOptionPane.showMessageDialog(null, lerDocumentos(exibir, arquivo));
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, e.getMessage());
		}
	}
	
	public static String escreveDocumentos(EscreverDocumento gerar, File arquivo) throws Exception{
		gerar = new EscreverDocumento(arquivo);
		gerar.mudarOrientacao("paisagem");
		gerar.inserirCabecalho();
		gerar.inserirTitulo();
		gerar.inserirParagrafo();
		gerar.inserirImagens("imagem.jpg"); //Path imagem.
		gerar.inserirRodape();
		gerar.inserirTabela();
		gerar.gerarArquivo();
		gerar.abrirArquivo(arquivo);
		return "Documento criado!";
	}
	
	public static String lerDocumentos(LerDocumento exibir, File arquivo) throws Exception{
		exibir = new LerDocumento(arquivo);
		exibir.leituraCabecalhoRodape();
		exibir.leituraParagrafos();
		exibir.leituraTabelas();
		exibir.leituraImagem();
		System.out.println("\nO texto todo:");
		exibir.leituraTextoCompleto();
		return "Documento lido!";
	}
	
	

}
