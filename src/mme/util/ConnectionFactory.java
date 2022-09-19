package mme.util;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionFactory {
	
	public Connection getConnection(String sid) throws IOException {
		try {
			
			String host = Util.getValor("host");
			String porta = Util.getValor("porta");
			//String sid = Util.getValor("sid");
			String usuario = Util.getValor("usuarioBanco");
			String senha = Util.getValor("senhaBanco");
			String urlConexao = "jdbc:sqlserver://" + host + ":" + porta + ";databaseName=" + sid;
			// spring.datasource.url=jdbc:sqlserver://12.12.12.115:1433;databaseName=XXXXX
			
			return DriverManager.getConnection(
					urlConexao,usuario,senha);
			} catch (SQLException e) {
				throw new RuntimeException(e);
			}
	}
	
}