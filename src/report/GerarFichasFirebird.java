/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package report;

import connection.connectionFactory;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;
import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JRResultSetDataSource;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.view.JasperViewer;

/**
 * Gera Ficha Cadastral para impressao
 *
 * @author user
 */
public class GerarFichasFirebird {

    private String relatorio;
    private String tipo;
    private String arg;
    private Connection con = null;

    public void fichas(String relatorio, String tipo, String arg) throws JRException, Exception {
        this.relatorio = relatorio;
        // this.cli = relatorio;
        String argumento = tipo + " " + "like '%" + arg + "%'";

        con = connectionFactory.getConnection();

        Statement stm = con.createStatement();
        String query;

        query = "SELECT  "
                + "(select MAX(cli_fichao.FCL_RAZAO) from cli_fichao where cli_fichao.FCL_COD = clientes.K_COD ) "
                + "	as FCL_RAZAO, "
                + "(select MAX(cli_fichao.FCL_KNOMEVEND) from cli_fichao where cli_fichao.FCL_COD = clientes.K_COD ) "
                + "	as vendedor, "
                + "(select MAX(cli_fichao.FCL_PROSPECTOR) from cli_fichao where cli_fichao.FCL_COD = clientes.K_COD ) "
                + "	as prospector, "
                + "(select MAX(cli_fichao.FCL_ATENDENTE) from cli_fichao where cli_fichao.FCL_COD = clientes.K_COD ) "
                + "	as atendente, "
                + "clientes.K_COD as codigo,  "
                + "clientes.K_CGCFIS as cgccli, clientes.K_IEFIS as iecli, clientes.K_SUFRAMA as suframa, "
                + "clientes.K_RAZAO as razaocli, clientes.K_APELIDO as apelidoCli , clientes.K_EMBBOBINA as bobina, "
                + "clientes.K_EMBDISCO as disco, clientes.K_EMAIL as emailcli, clientes.K_FONECON as fonecli, "
                + "clientes.K_HOMEPAGE as sitecli, clientes.K_BIP as bip, clientes.K_CODMUNICIPIO as ibgecli, "
                + "clientes.K_CODMUNENT as ibgeent, clientes.K_ENDFIS as endfiscli, clientes.K_NUMFIS as numfiscli, "
                + "clientes.K_COMPLFIS as complifiscli, clientes.K_BAIFIS as bairrofiscli, clientes.K_CEPFIS as cepfiscli, "
                + "clientes.K_CEPFIS1 as cepfis1cli, clientes.K_MUNFIS as munfiscli, clientes.K_ESTFIS as estfiscli, "
                + "clientes.K_ENDCOB as endcobcli, clientes.K_NUMCOB as numcob, clientes.K_COMPLCOB as complcob, "
                + "clientes.K_BAICOB as baicob, clientes.K_CEPCOB as cepcob, clientes.K_CEPCOB1 cepcob1, "
                + "clientes.K_MUNCOB as muncob1, clientes.K_ESTCOB as estcob1, clientes.K_ENDENT as endent, "
                + "clientes.K_NUMENT as nument, clientes.K_COMPLENT as complent, clientes.K_BAIENT as baient, "
                + "clientes.K_CEPENT as cepent, clientes.K_CEPENT1 as cepent1, clientes.K_MUNENT as munent1, "
                + "clientes.K_ESTENT as estent, transpor.K_CGCFIS as cgctra, transpor.K_IEFIS as ietra, "
                + "transpor.K_RAZAO as razaotra, transpor.K_APELIDO as apelidotra, transpor.K_HOMEPAGE as sitetra, "
                + "transpor.K_FONECON as fonetra, transpor.K_EMAIL as emailtra, transpor.K_CODMUNICIPIO as ibgetra, "
                + "transpor.K_ENDFIS as endtra, transpor.K_NUMFIS as numtra, transpor.K_COMPLFIS as compltra, "
                + "transpor.K_BAIFIS as bairrotra, transpor.K_CEPFIS as ceptra, transpor.K_CEPFIS1 as cep1tra, "
                + "transpor.K_MUNFIS as muntra, transpor.K_ESTFIS as esttra "
                + "from clientes full join transpor on clientes.K_CODTRANS = transpor.K_COD  "
                + "where " + argumento + " ORDER BY codigo asc ";

        ResultSet rs = stm.executeQuery(query);

        JRResultSetDataSource jrRS = new JRResultSetDataSource(rs);

        Map parameters = new HashMap();

        JasperPrint jp = JasperFillManager.fillReport(relatorio, parameters, jrRS);

        JasperViewer.viewReport(jp, false);

    }
}
