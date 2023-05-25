package connection;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Classe de Conexão com banco de dados Firebird
 *
 * @author user
 */
public class connectionFactory {

//Início da classe de conexão//
    private static final String DRIVER = "org.firebirdsql.jdbc.FBDriver";
    private static final String URL = "jdbc:firebirdsql://192.168.0.3/C:/Quatro14/ArqSql/MOISES_FINAL.AX1";
    private static final String USER = "SYSDBA";
    private static final String PASS = "masterkey";

    /**
     * Tenta Realizar conexão com base nos dados impostos...
     *
     * @return
     */
    public static Connection getConnection() {

        try {
            Class.forName(DRIVER);

            return DriverManager.getConnection(URL, USER, PASS);

        } catch (ClassNotFoundException | SQLException ex) {
            throw new RuntimeException("Erro na conexao", ex);
        }
    }

    /**
     * Encerra conexão com banco de dados..
     *
     * @param con
     */
    public static void closeConnection(Connection con) {

        if (con != null) {
            try {
                con.close();
            } catch (SQLException ex) {
                System.err.println("Erro: " + ex);
            }
        }
    }

    /**
     * Encerra conexão com banco de dados..
     *
     * @param con
     * @param pstmt
     */
    public static void closeConnection(Connection con, PreparedStatement pstmt) {

        if (pstmt != null) {
            try {
                con.close();
            } catch (SQLException ex) {
                System.err.println("Erro: " + ex);
            }
        }

        closeConnection(con);
    }

    /**
     * Encerra conexão com banco de dados..
     *
     * @param con
     * @param pstmt
     * @param rs
     */
    public static void closeConnection(Connection con, PreparedStatement pstmt, ResultSet rs) {

        if (rs != null) {
            try {
                rs.close();
            } catch (SQLException ex) {
                System.err.println("Erro: " + ex);
            }
        }

        closeConnection(con, pstmt);
    }
    /**
     * Teste de conexão
     */
    /*
    public static void main(String[] args) throws ClassNotFoundException {
    connectionFactory.getConnection();
    }
     */

}
