package mme;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;

import mme.util.ConnectionFactory;
import mme.util.Util;

public class HistoricoExecucaoDao {
	
	private Connection connection;
	
	 public HistoricoExecucaoDao() throws IOException {
		 
		 String sid = Util.getValor("sidHistoricoExecucao");
	
		 this.connection = new ConnectionFactory().getConnection(sid);
	 
	 }
	 
	 public void inserirStatusExecucao(String servico, String dataHora, String status) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO Tb_Historico_Execucao_Robos   (               ";
			 sql += "Servico,                                                        ";
			 sql += "Data_Inicial_Extracao,                                          ";
			 sql += "Data_Final_Extracao,                                            ";
			 sql += "Status                                                          ";
			 sql +=  ") VALUES (                                                     ";
			 sql += "?,                                                              ";
			 sql += "?,                                                              ";
			 sql += "?,                                                              ";			 
			 sql += "?                                                               ";
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 String dataAtualFimExecucao = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
			 statement.setString(1, servico);
			 statement.setString(2, dataHora);
			 statement.setString(3, dataAtualFimExecucao);
			 statement.setString(4, status);
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		 
	 }

}
