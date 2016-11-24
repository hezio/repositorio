package br.com.bancoamazonia.persistencia;

/**
 * Geração do Relatório de Capitalização (Campanha Capitalização)
 * @author 6330 - Hezio Silva
 * @version 1.0
 * @since 18/11/216
 * @see http://www.w3ii.com/pt/apache_poi/apache_poi_database.html 
 */

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SisVendasCap {
	public static void main(String[] args) throws Exception {
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		Connection connect = DriverManager
				.getConnection("jdbc:sqlserver://srv1007:1434;databaseName=SISICATU;user=pd_case;password=pdbasa");
		Statement statement = connect.createStatement();
		ResultSet resultSet = statement
				.executeQuery("SELECT" + "	CAST(AG.AGEN_SUPER AS VARCHAR(10)) + ' - ' + SUP.SUP_DES 'SUPER',"
						+ "	CAST(PRO.AGEN_COD AS VARCHAR(10)) 'AGENCIA'," + "	AG.AGEN_DESC 'AGENCIA_NOME',"
						+ "	PRO.SUBS_CPFCNPJ 'CPF/CNPJ'," + "	SUB.SUBS_NOME 'NOME'," + "	PRD_DES 'PRODUTO',"
						+ "	PRO.PRO_NUM 'PROPOSTA'," + "	EVE.STAT_EVENTO + ' - ' + ST.STAT_DESC 'SITUACAO',"
						+ "	RIGHT(CONVERT(VARCHAR(10),PRO.PRO_DTASSIN),2) + '/' + SUBSTRING(CONVERT(VARCHAR(10),PRO.PRO_DTASSIN),5,2) + '/' + "
						+ "	LEFT(CONVERT(VARCHAR(10),PRO.PRO_DTASSIN),4) 'DATA ASSINATURA',"
						+ "	PRO.FUNC_MATR +' - ' + FUNC.FUNC_NOME 'ANGARIADOR',"
						+ "	RIGHT(CONVERT(VARCHAR(10),EVE.EVEN_DAT),2) + '/' + SUBSTRING(CONVERT(VARCHAR(10),EVE.EVEN_DAT),5,2) + '/' + LEFT(CONVERT(VARCHAR(10),EVE.EVEN_DAT),4)  'DATA PREVISTA DEBITO',"
						+ "		(case EVE.STAT_EVENTO"
						+ "		when 'PG' then (RIGHT(CONVERT(VARCHAR(10),EVE.EVEN_DATEFETIVA),2) + '/' + SUBSTRING(CONVERT(VARCHAR(10),EVE.EVEN_DATEFETIVA),5,2) + '/' + LEFT(CONVERT(VARCHAR(10),EVE.EVEN_DATEFETIVA),4))"
						+ "		else null" + "		end)'DATA EFETIVA DEBITO CONTA - ARCE',"
						+ "	RIGHT(CONVERT(VARCHAR(10),EVE.EVEN_DATENVICA),2) + '/' + SUBSTRING(CONVERT(VARCHAR(10),EVE.EVEN_DATENVICA),5,2) + '/' + LEFT(CONVERT(VARCHAR(10),EVE.EVEN_DATENVICA),4) 'DATA RETORNO ICATU',  "
						+ "	PRO.AGEN_CC 'CONTA'," + "	REPLACE(CAST(EVE.EVEN_VALOR AS VARCHAR),'.',',') 'VALOR'" + " "
						+ "FROM  TAB_STATUS ST, TAB_PROPOSTA PRO "
						+ "INNER JOIN TAB_FUNCIONARIO FUNC ON FUNC.FUNC_MATR = PRO.FUNC_MATR "
						+ "INNER JOIN TAB_AGENCIA AG ON AG.AGEN_COD = PRO.AGEN_COD "
						+ "INNER JOIN TAB_SUPER SUP ON SUP.SUP_COD = AG.AGEN_SUPER "
						+ "INNER JOIN TAB_PRODUTO PRD ON PRO.PRD_COD = PRD.PRD_COD "
						+ "INNER JOIN TAB_EVENTO EVE  ON PRO.PRD_COD = EVE.PRD_COD AND PRO.PLN_COD = EVE.PLN_COD AND PRO.PRO_NUM = EVE.PRO_NUM "
						+ "INNER JOIN TAB_SUBSCRITOR SUB ON SUB.SUBS_CPFCNPJ = PRO.SUBS_CPFCNPJ " + " " + "WHERE "
						+ "	PRO.PRD_COD IN (27,28,29) " + "	AND PRO.PRO_DTAPROPOSTA BETWEEN '20161116' AND '20161123' "
						+ "	AND EVE.TIP_EVENTO = 'P1' " + "	AND PRO.PRO_STPROPOSTA = 'AV' "
						+ "	AND ST.STAT_COD = EVE.STAT_EVENTO " + "ORDER BY PRO.PRO_STPROPOSTA,PRO.PRO_DTAPROPOSTA");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("SISVENDAS");
		XSSFRow row = spreadsheet.createRow(0);
		XSSFCell cell;
		cell = row.createCell(0);
		cell.setCellValue("SUPER");
		cell = row.createCell(1);
		cell.setCellValue("AGENCIA");
		cell = row.createCell(2);
		cell.setCellValue("AGENCIA_NOME");
		cell = row.createCell(3);
		cell.setCellValue("CPF/CNPJ");
		cell = row.createCell(4);
		cell.setCellValue("NOME");
		cell = row.createCell(5);
		cell.setCellValue("PRODUTO");
		cell = row.createCell(6);
		cell.setCellValue("PROPOSTA");
		cell = row.createCell(7);
		cell.setCellValue("SITUACAO");
		cell = row.createCell(8);
		cell.setCellValue("DATA ASSINATURA");
		cell = row.createCell(9);
		cell.setCellValue("ANGARIADOR");
		cell = row.createCell(10);
		cell.setCellValue("DATA PREVISTA DEBITO");
		cell = row.createCell(11);
		cell.setCellValue("DATA EFETIVA DEBITO CONTA - ARCE");
		cell = row.createCell(12);
		cell.setCellValue("DATA RETORNO ICATU");
		cell = row.createCell(13);
		cell.setCellValue("CONTA");
		cell = row.createCell(14);
		cell.setCellValue("VALOR");

		int i = 1;
		while (resultSet.next()) {
			row = spreadsheet.createRow(i);
			cell = row.createCell(0);
			cell.setCellValue(resultSet.getString("SUPER"));
			cell = row.createCell(1);
			cell.setCellValue(resultSet.getString("AGENCIA"));
			cell = row.createCell(2);
			cell.setCellValue(resultSet.getString("AGENCIA_NOME"));
			cell = row.createCell(3);
			cell.setCellValue(resultSet.getString("CPF/CNPJ"));
			cell = row.createCell(4);
			cell.setCellValue(resultSet.getString("NOME"));
			cell = row.createCell(5);
			cell.setCellValue(resultSet.getString("PRODUTO"));
			cell = row.createCell(6);
			cell.setCellValue(resultSet.getString("PROPOSTA"));
			cell = row.createCell(7);
			cell.setCellValue(resultSet.getString("SITUACAO"));
			cell = row.createCell(8);
			cell.setCellValue(resultSet.getString("DATA ASSINATURA"));
			cell = row.createCell(9);
			cell.setCellValue(resultSet.getString("ANGARIADOR"));
			cell = row.createCell(10);
			cell.setCellValue(resultSet.getString("DATA PREVISTA DEBITO"));
			cell = row.createCell(11);
			cell.setCellValue(resultSet.getString("DATA EFETIVA DEBITO CONTA - ARCE"));
			cell = row.createCell(12);
			cell.setCellValue(resultSet.getString("DATA RETORNO ICATU"));
			cell = row.createCell(13);
			cell.setCellValue(resultSet.getString("CONTA"));
			cell = row.createCell(14);
			cell.setCellValue(resultSet.getString("VALOR"));

			i++;
		}

		SimpleDateFormat sdfd = new SimpleDateFormat("ddMMyyyy");
		String dataAtual = sdfd.format(new Date());

		SimpleDateFormat sdfh = new SimpleDateFormat("HHmm");
		Date hora = Calendar.getInstance().getTime();
		String horaFormatada = sdfh.format(hora);

		FileOutputStream out = new FileOutputStream(new File("sisvendas_" + dataAtual + "_" + horaFormatada + ".xlsx"));
		workbook.write(out);
		out.close();
		System.out.println("sisvendas_" + dataAtual + "_" + horaFormatada + ".xlsx" + " gerado com sucesso!");
	}
}