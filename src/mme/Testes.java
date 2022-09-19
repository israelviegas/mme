package mme;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class Testes {

	public static void main(String[] args) {

		
		int contador = 0;
		int indice = 0;
		
		   DateFormat df = new SimpleDateFormat ("dd/MM/yyyy");  
		   
		   Calendar cal = Calendar.getInstance();
		   cal.setTime(new Date());
			 
		     if (cal.get(Calendar.DAY_OF_MONTH) == 1){
		    	 cal.add (Calendar.DATE, -2); 
		     } else if (cal.get(Calendar.DAY_OF_WEEK) == 2){
		    	 cal.add (Calendar.DATE, -3); 
		     } else {
		    	 cal.add (Calendar.DATE, -1); 
		     }
		
    	String arquivo = "Cost Transaction Extract(Everything Report)_20210807_1200.xlsx";
    	int contadorParaAlterarNomeArquivos = 0;
    	contadorParaAlterarNomeArquivos = contadorParaAlterarNomeArquivos + 1;
    	String[] partesNomeArquivo = arquivo.split("\\.");
    	
    	if (partesNomeArquivo != null && partesNomeArquivo.length > 0) {
    		String nome = partesNomeArquivo[0];
    		nome = nome + "_" + contadorParaAlterarNomeArquivos;
    		String extensao = partesNomeArquivo[1];
    		arquivo = nome + "." + extensao;
    	}
		
		String espaco = "Client Delivery & Operations            1";
		System.out.println(espaco.trim());
		
		indice = contador++;
		System.out.println(indice);
		
		indice = contador++;
		System.out.println(indice);
		
		indice = contador++;
		System.out.println(indice);
		
		String teste = "folder\ninsert_drive_file\n9940191116 SW Factories\nContract and Project with sourcing(Open), with Delivery Reporting(Retired) Entity\nSelectable";
		
		//System.out.println(teste);
		
		String mes = "05";
		
		System.out.println("Mês: " + Integer.parseInt(mes));
		
		String[] partes = teste.split("aaaaaa");
		
		
		String teste2 = teste.replaceAll("\\n", "teste");
		
		//System.out.println(partes[2]);
		
		
		String mesAtual = new SimpleDateFormat("MM").format(new Date());
		
		
		System.out.println(mesAtual);
		
		String somenteNumeros = "Master Active - Julio.2021-";
		
		System.out.println("somente números: " + somenteNumeros.replaceAll("[^0-9]", ""));
		

		
		
		
		
		
	}

}
