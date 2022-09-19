package mme;

import java.util.Scanner;

public class Teste {

public static void main(String[] args) throws InterruptedException {
		
	iniciarInteligenciaArtificial();

}


public static void iniciarInteligenciaArtificial () throws InterruptedException {
	
	Scanner input = new Scanner(System.in);
	boolean yes, no;
	String resposta;
	
	//----------------------------------------------------------------------------------------------------
	
	
	System.out.println("Ola meu nome eh Millsz e irei falar um pouco com voce deseja comecar ? (yes/no)");
	resposta = input.next();
	
	if(resposta.equals("yes")) {
		
		respostaSim();
		
	} else if (resposta.equals("no")) {
		System.out.println("Ah que pena, volte mais tarde");
	} else {
		System.out.println("Opcao invalida, digite yes ou no na pergunta abaixo");
		Thread.sleep(5000);
		iniciarInteligenciaArtificial();
	}
	
	
}

public static  void respostaSim () {
	
	System.out.println("Vamos comecar entao");
	
	// Ações para o Sim

}
	

}
