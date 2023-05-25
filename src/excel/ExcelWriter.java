package excel;

import connection.connectionFactory;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import model.fundicao;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Classe responsável por gerar o banco de dados...
 *
 * @author user
 */
public class ExcelWriter {

    private static Connection con = null;

    /**
     * Recebe os dados da outra tela para consulta no banco de dados, inicia as
     * células (e o layout) que sera apresentando e Gera o Excel
     *
     * @param data1Entrada
     * @param data2Entrada
     * @param dataVe1
     * @param dataVe2
     * @throws IOException
     * @throws InvalidFormatException
     * @throws SQLException
     * @throws java.text.ParseException
     */
    public static void ExcelWriter(String data1Entrada, String data2Entrada, Date dataVe1, Date dataVe2) throws IOException, InvalidFormatException, SQLException, ParseException {
        /* CreationHelper helps us create instances of various things like DataFormat,
        Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        try ( // Create a Workbook
                Workbook workbook = new XSSFWorkbook() // Gerra arquivo .xls
                ) {
            /* CreationHelper helps us create instances of various things like DataFormat,
            Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */

            CreationHelper createHelper = workbook.getCreationHelper();
            // Cria a planilha 
            Sheet sheet = workbook.createSheet("FUNDIÇÂO");
            // ESTILOS UTILIZADOS
            // Estabele a Fonte utilizada
            Font headerFont = workbook.createFont();
            headerFont.setFontHeightInPoints((short) 10);

            // CELULA YELLOW FORMATADO NA HORIZONTAL
            CellStyle linha3Yellow = workbook.createCellStyle();
            linha3Yellow.setFont(headerFont);
            linha3Yellow.setVerticalAlignment(VerticalAlignment.CENTER);
            linha3Yellow.setAlignment(HorizontalAlignment.CENTER);
            linha3Yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            linha3Yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            linha3Yellow.setBorderTop(BorderStyle.THIN);
            linha3Yellow.setBorderRight(BorderStyle.THIN);
            linha3Yellow.setBorderLeft(BorderStyle.THIN);

            // CELULA YELLOW FORMATADO NA VERTICAL
            CellStyle linha4Yellow90 = workbook.createCellStyle();
            linha4Yellow90.setFont(headerFont);
            linha4Yellow90.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4Yellow90.setAlignment(HorizontalAlignment.CENTER);
            linha4Yellow90.setRotation((short) 90);
            linha4Yellow90.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            linha4Yellow90.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha4Yellow90.setBorderRight(BorderStyle.THIN);
            linha4Yellow90.setBorderLeft(BorderStyle.THIN);

            //CELULA BRIGHT GREEN FORMATADO NA HORIZONTAL
            CellStyle linha3BrightGreen = workbook.createCellStyle();
            linha3BrightGreen.setFont(headerFont);
            linha3BrightGreen.setAlignment(HorizontalAlignment.CENTER);
            linha3BrightGreen.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            linha3BrightGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            linha3BrightGreen.setBorderTop(BorderStyle.THIN);
            linha3BrightGreen.setBorderRight(BorderStyle.THIN);
            linha3BrightGreen.setBorderLeft(BorderStyle.THIN);

            //CELULA BRIGHT GREEN FORMATADO NA HORIZONTAL
            CellStyle linha3Pink = workbook.createCellStyle();
            linha3Pink.setFont(headerFont);
            linha3Pink.setAlignment(HorizontalAlignment.CENTER);
            linha3Pink.setFillForegroundColor(IndexedColors.PINK.getIndex());
            linha3Pink.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            linha3Pink.setBorderTop(BorderStyle.THIN);
            linha3Pink.setBorderRight(BorderStyle.THIN);
            linha3Pink.setBorderLeft(BorderStyle.THIN);

            //CELULA ROSE FORMATADO NA HORIZONTAL
            CellStyle linha3Rose = workbook.createCellStyle();
            linha3Rose.setFont(headerFont);
            linha3Rose.setVerticalAlignment(VerticalAlignment.CENTER);
            linha3Rose.setAlignment(HorizontalAlignment.CENTER);
            linha3Rose.setRotation((short) 90);
            linha3Rose.setFillForegroundColor(IndexedColors.ROSE.getIndex());
            linha3Rose.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha3Rose.setBorderRight(BorderStyle.THIN);
            linha3Rose.setBorderLeft(BorderStyle.THIN);

            //CELULA LIME FORMATADO NA HORZONTAL
            CellStyle linha3Lime = workbook.createCellStyle();
            linha3Lime.setFont(headerFont);
            linha3Lime.setAlignment(HorizontalAlignment.CENTER);
            linha3Lime.setAlignment(HorizontalAlignment.CENTER);
            linha3Lime.setAlignment(HorizontalAlignment.CENTER);
            linha3Lime.setFillForegroundColor(IndexedColors.LIME.getIndex());
            linha3Lime.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha3Lime.setBorderTop(BorderStyle.THIN);

            //CELULA LIME FORMATADO NA VERTICAL
            CellStyle linha4Lime90 = workbook.createCellStyle();
            linha4Lime90.setFont(headerFont);
            linha4Lime90.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4Lime90.setAlignment(HorizontalAlignment.CENTER);
            linha4Lime90.setRotation((short) 90);
            linha4Lime90.setFillForegroundColor(IndexedColors.LIME.getIndex());
            linha4Lime90.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha4Lime90.setBorderBottom(BorderStyle.THIN);
            linha4Lime90.setBorderRight(BorderStyle.THIN);
            linha4Lime90.setBorderLeft(BorderStyle.THIN);

            //CELULA WHITE FORMATADO NA VERTICAL
            CellStyle linha4White90 = workbook.createCellStyle();
            linha4White90.setFont(headerFont);
            linha4White90.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4White90.setAlignment(HorizontalAlignment.CENTER);
            linha4White90.setRotation((short) 90);
            linha4White90.setBorderBottom(BorderStyle.THIN);
            linha4White90.setBorderRight(BorderStyle.THIN);
            linha4White90.setBorderLeft(BorderStyle.THIN);

            //CELULA WHITE FORMATADO NA HORIZONTAL
            CellStyle linha4White = workbook.createCellStyle();
            linha4White.setFont(headerFont);
            linha4White.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4White.setAlignment(HorizontalAlignment.CENTER);
            linha4White.setBorderBottom(BorderStyle.THIN);
            linha4White.setBorderRight(BorderStyle.THIN);
            linha4White.setBorderLeft(BorderStyle.THIN);

            //CELULA TAN FORMATADO NA HORIZONTAL
            CellStyle linha4Tan = workbook.createCellStyle();
            linha4Tan.setFont(headerFont);
            linha4Tan.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4Tan.setAlignment(HorizontalAlignment.CENTER);
            linha4Tan.setFillForegroundColor(IndexedColors.TAN.getIndex());
            linha4Tan.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha4Tan.setBorderBottom(BorderStyle.THIN);
            linha4Tan.setBorderRight(BorderStyle.THIN);
            linha4Tan.setBorderLeft(BorderStyle.THIN);

            //CELULA PALE_BLUE FORMATADO NA VERTICAL
            CellStyle linha4PaleBlue90 = workbook.createCellStyle();
            linha4PaleBlue90.setFont(headerFont);
            linha4PaleBlue90.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4PaleBlue90.setAlignment(HorizontalAlignment.CENTER);
            linha4PaleBlue90.setRotation((short) 90);
            linha4PaleBlue90.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
            linha4PaleBlue90.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha4PaleBlue90.setBorderBottom(BorderStyle.THIN);
            linha4PaleBlue90.setBorderRight(BorderStyle.THIN);
            linha4PaleBlue90.setBorderLeft(BorderStyle.THIN);

            //CELULA RED FORMATADO NA VERTICAL
            CellStyle linha4Red90 = workbook.createCellStyle();
            linha4Red90.setFont(headerFont);
            linha4Red90.setVerticalAlignment(VerticalAlignment.CENTER);
            linha4Red90.setAlignment(HorizontalAlignment.CENTER);
            linha4Red90.setRotation((short) 90);
            linha4Red90.setFillForegroundColor(IndexedColors.RED.getIndex());
            linha4Red90.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            linha4Red90.setBorderBottom(BorderStyle.THIN);
            linha4Red90.setBorderRight(BorderStyle.THIN);
            linha4Red90.setBorderLeft(BorderStyle.THIN);

            //CELULA FORMATANDO CELULAS RECEBIDAS DO BANCO DE DADOS
            CellStyle alignDataStyle = workbook.createCellStyle();
            alignDataStyle.setAlignment(HorizontalAlignment.CENTER);
            alignDataStyle.setBorderBottom(BorderStyle.THIN);
            alignDataStyle.setBorderTop(BorderStyle.THIN);
            alignDataStyle.setBorderRight(BorderStyle.THIN);
            alignDataStyle.setBorderLeft(BorderStyle.THIN);
            alignDataStyle.setFont(headerFont);

            //CELULA FORMATANDO DATAS RECEBIDAS DO BANCO DE DADOS
            CellStyle dateCellStyle = workbook.createCellStyle();
            dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yy"));
            dateCellStyle.setAlignment(HorizontalAlignment.CENTER);
            dateCellStyle.setAlignment(HorizontalAlignment.CENTER);
            dateCellStyle.setBorderBottom(BorderStyle.THIN);
            dateCellStyle.setBorderTop(BorderStyle.THIN);
            dateCellStyle.setBorderRight(BorderStyle.THIN);
            dateCellStyle.setBorderLeft(BorderStyle.THIN);
            dateCellStyle.setFont(headerFont);

            //CELULA FORMATANDO NUMEROS RECEBIDAS DO BANCO DE DADOS
            CellStyle numCellStyle = workbook.createCellStyle();
            numCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
            numCellStyle.setAlignment(HorizontalAlignment.CENTER);
            numCellStyle.setAlignment(HorizontalAlignment.CENTER);
            numCellStyle.setBorderBottom(BorderStyle.THIN);
            numCellStyle.setBorderTop(BorderStyle.THIN);
            numCellStyle.setBorderRight(BorderStyle.THIN);
            numCellStyle.setBorderLeft(BorderStyle.THIN);
            numCellStyle.setFont(headerFont);

            //CELULA FORMATANDO PORCENTAGEM RECEBIDAS DO BANCO DE DADOS
            CellStyle porCellStyle = workbook.createCellStyle();
            porCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("#,##0%"));
            porCellStyle.setAlignment(HorizontalAlignment.CENTER);
            porCellStyle.setAlignment(HorizontalAlignment.CENTER);
            porCellStyle.setBorderBottom(BorderStyle.THIN);
            porCellStyle.setBorderTop(BorderStyle.THIN);
            porCellStyle.setBorderRight(BorderStyle.THIN);
            porCellStyle.setBorderLeft(BorderStyle.THIN);
            porCellStyle.setFont(headerFont);

            // CONEXÃO E MAPEAMENTO COM BANCO DE DADOS...
            con = connectionFactory.getConnection();
            PreparedStatement pstmt = con.prepareStatement(
                    "SELECT"
                    //Variáveis da tabela fundicao
                    + " F_PEDIDO, F_EMISSAO, F_DATAAPROVFK, F_DATAISO, F_DATAPEDIDO, F_DATAOS, F_PRAZOFUNDIR, F_NUMPLACASMAIOR90, F_DATAENVFUND, F_LOTENUM,"
                    + " F_MEDIDA, F_DUREZA, F_OSNUM, F_CLIENTE, F_F0, F_F1, F_F2, F_F3, F_SUPER, F_3105, F_OUTRASLIGAS, F_PLACANUM, F_LARGREALOS, F_KGCALC, F_LARGFUND,"
                    + " F_NUMBOBINA, F_DURSOLICOS, F_SETOR, F_ATENDENTE, F_TRATADOR, F_DATAFUND, F_DEFFUND, F_QTDEFUND, F_OPERADORFUND, F_DATALQ, F_DEFLQ, F_QTDELQ,"
                    + " F_OPERADORLQ, F_DATASLIT, F_DEFSLIT, F_QTDESLIT, F_OPERADORSLIT, F_DATAEG, F_ESPMM, F_DATALF, F_DEFLF, F_QTDELF, F_OPERADORLF, F_DATALFF2,"
                    + " F_DEFLFF2, F_QTDELFF2, F_OPERADORLFF2, F_DATALFF1, F_DEFLFF1, F_QTDELFF1, F_OPERADORLFF1, F_DATAE, F_DEFE, F_QTDEE, F_OPERADORE, F_DATARLBUHL,"
                    + " F_DATARLLFF2, F_DATARLLFF1, F_DATARLLA9, F_DEFLA9, F_QTDELA9, F_OPERADORLA9, F_DEFRL, F_QTDERL, F_OPERADORRL, F_DATAROT, F_DEFROT, F_QTDEROT,"
                    + " F_OPERADORROT, F_DATASF, F_DEFSF, F_QTDESF, F_OPERADORSF, F_DATASFI, F_DEFSFI, F_QTDESFI, F_OPERADORSFI, F_DATAP, F_DEFP, F_QTDEP, F_OPERADORP,"
                    + " F_DATAGUI, F_DEFGUI, F_QTDEGUI, F_OPERADORGUI, F_DATAGUIF, F_DEFGUIF, F_QTDEGUIF, F_OPERADORGUIF, F_DATAR, F_DEFR, F_QTDER, F_OPERADORR, F_DATAPE,"
                    + " F_DEFPE, F_QTDEPE, F_OPERADORPE, F_DATAC, F_DEFC, F_QTDEC, F_OPERADORC, F_DATAESC, F_DEFESC, F_QTDEESC, F_OPERADORESC, F_PRAZOENT, F_KGTEOR, F_QUANT,"
                    + " F_KGEST, F_DATAS1, F_S1, F_DATAS2, F_S2, F_DATAS3, F_S3, F_DATAS4, F_S4, F_DATAS5, F_S5, F_DATAS6, F_S6, F_DATAS7, F_S7, F_DATAS8, F_S8, F_DATAS9,"
                    + " F_S9, F_DATAS10, F_S10, F_DATAS11, F_S11, F_DATAS12, F_S12, F_DATAS13, F_S13, F_DATAS14, F_S14, F_DATAS15, F_S15, F_DATAS16, F_S16, F_DATAS17, F_S17,"
                    + " F_DATAS18, F_S18, F_DATAS19, F_S19, F_DATAS20, F_S20, F_TOTAL, F_PERC, F_SEQ,"
                    //variáveis da tabela FUND_DEFEITOS
                    + " FD_SEQPF, FD_ESPESSURA, FD_OXIDACAO, FD_ONDULACAO, FD_CASQUINHA, FD_RASTRO_TRATOR, FD_MANCHA_OLEO, FD_LINHA_DO_LQ, FD_RISCO_DA_PRENSA, FD_SUJEIRA,"
                    + " FD_NATA, FD_CAVACO, FD_LINHA_PRETA, FD_BOLHA_CAB, FD_BOLHA_TUB, FD_BOLHA_DES, FD_RISCO, FD_QUEBRA_DE_BOB, FD_BURACO, FD_MANCHA_BRANCA, FD_MANCHA_MARROM,"
                    + " FD_BATIMENTO_LATERAL, FD_BOLHA_AQUECIMENTO, FD_GRAO, FD_DUREZADEF, FD_SOBRA, FD_FLECHA, FD_FALTOU_PESO, FD_TEMPERATURA, FD_OS_PERDIDA, FD_PE_DE_PLACA,"
                    + " FD_REBARBA, FD_MANCHA_D_AGUA, FD_OLHO, FD_ZINABRO, FD_DESALINHADA, FD_SUMIU, FD_TRINCADA, FD_BOB_ESTREITA, FD_MARCA_CILINDRO, FD_SEM_OS,"
                    + " FD_LATERAL_AMASSADA, FD_QUEBRA_PROCED, FD_NFL, FD_FUND_MED_MENOS, FD_FUND_MED_MAIS, FD_PLACA_BOLHA_PQ, FD_MANCHA_PRETA, FD_POZINHO, FD_BOB_FROUCHA,"
                    + " FD_BORRA_CHAPA, FD_MANCHA_ESTUFA, FD_MAL_ENROLADA, FD_MAT_FALTA_LUB, FD_EXCESSO_BORRA, FD_BOB_MAL_SOLDADA, FD_FALTA_TESTE_MAT, FD_MAT_MARCADO,"
                    + " FD_EXCESSO_OLEO, FD_FALTA_LIMPEZA, FD_MATERIAL_FIAPO, FD_MATERIAL_CHUVISCO, FD_MAT_FORA_ESQUADRO, FD_MAQUINA_QUEBRADA, FD_FALTA_ATENCAO,"
                    + " FD_ERRO_MEDIDA, FD_FALTA_TESTE_MAQ, FD_NAO_FEZ_MANUT, FD_FALTA_REVISAO, FD_GASTO_EXCESSIVO, FD_MANUT_SEM_EFIC, FD_FALTA_MANUT, FD_PROD_DANIFICADO,"
                    + " FD_FORA_BERCOS, FD_ATRASOU_SERVICO, FD_PRODUCAO_BAIXA, FD_FALTOU_DIA, FD_TROCA_PAPELETA, FD_FALTA_DISCIPLINA, FD_SEM_PROT_AURICULAR, FD_SEM_OCULOS,"
                    + " FD_SEM_LUVA, FD_SEM_SAPATO_PROT, FD_CORREDOR_OBST, FD_QUEBRA_NORMA, FD_MAT_TEMP_ALTA, FD_MAT_TEMP_BAIXA, FD_RISCO_LQ, FD_RISCO_SG, FD_RISCO_LF,"
                    + " FD_RISCO_SF, FD_PEDIDO_CANCEL, FD_BOB_DESVIADA, FD_MEIA_LUA, FD_REFILE_MAIOR, FD_ERRO_DE_MEDIA,"
                    //COLUNAS DA TABELA CONTROLE_DISCAPROV
                    + " DATA_LEVANTAMENTO, NUMDISC, CULPADO "

                    + " FROM FUNDICAO LEFT JOIN FUND_DEFEITOS on F_SEQ = FD_SEQPF"
                    + " LEFT JOIN CONTROLE_DISCAPROV on F_SEQ = SEQPF"
                    + " where F_DATAFUND between  '" + data1Entrada + "' and '" + data2Entrada + "' "
                    + " ORDER by F_DATAFUND asc ");
            ResultSet rs = pstmt.executeQuery();
            //CRIA LISTAGEM ARRAY LIST
            ArrayList<fundicao> arrayList = new ArrayList<>();
            //ADICIONA AS INFORMAÇÕES DO BANCO DE DADOS NO Array List
            while (rs.next()) {
                fundicao Fundicao = new fundicao();

                Fundicao.setF_PEDIDO(rs.getString("F_PEDIDO"));
                Fundicao.setF_EMISSAO(rs.getDate("F_EMISSAO"));
                Fundicao.setF_DATAAPROVFK(rs.getDate("F_DATAAPROVFK"));
                Fundicao.setF_DATAISO(rs.getDate("F_DATAISO"));
                Fundicao.setF_DATAPEDIDO(rs.getDate("F_DATAPEDIDO"));
                Fundicao.setF_DATAOS(rs.getDate("F_DATAOS"));
                Fundicao.setF_PRAZOFUNDIR(rs.getString("F_PRAZOFUNDIR"));
                Fundicao.setF_NUMPLACASMAIOR90(rs.getString("F_NUMPLACASMAIOR90"));
                Fundicao.setF_DATAENVFUND(rs.getDate("F_DATAENVFUND"));
                Fundicao.setF_LOTENUM(rs.getDouble("F_LOTENUM"));
                Fundicao.setF_MEDIDA(rs.getString("F_MEDIDA"));
                Fundicao.setF_DUREZA(rs.getString("F_DUREZA"));
                Fundicao.setF_OSNUM(rs.getString("F_OSNUM"));
                Fundicao.setF_CLIENTE(rs.getString("F_CLIENTE"));

                Fundicao.setF_F0(rs.getString("F_F0"));
                Fundicao.setF_F1(rs.getString("F_F1"));
                Fundicao.setF_F2(rs.getString("F_F2"));
                Fundicao.setF_F3(rs.getString("F_F3"));
                Fundicao.setF_SUPER(rs.getString("F_SUPER"));
                Fundicao.setF_3105(rs.getString("F_3105")); //F1-27
                Fundicao.setF_OUTRASLIGAS(rs.getString("F_OUTRASLIGAS"));

                Fundicao.setF_PLACANUM(rs.getString("F_PLACANUM"));
                Fundicao.setF_LARGREALOS(rs.getString("F_LARGREALOS"));
                Fundicao.setF_KGCALC(rs.getFloat("F_KGCALC"));
                Fundicao.setF_LARGFUND(rs.getString("F_LARGFUND"));
                Fundicao.setF_NUMBOBINA(rs.getString("F_NUMBOBINA"));
                Fundicao.setF_DURSOLICOS(rs.getString("F_DURSOLICOS"));
                Fundicao.setF_SETOR(rs.getString("F_SETOR"));
                Fundicao.setF_ATENDENTE(rs.getString("F_ATENDENTE"));
                Fundicao.setF_TRATADOR(rs.getString("F_TRATADOR"));

                Fundicao.setF_DATAFUND(rs.getDate("F_DATAFUND"));
                Fundicao.setF_DEFFUND(rs.getString("F_DEFFUND"));
                Fundicao.setF_QTDEFUND(rs.getFloat("F_QTDEFUND"));
                Fundicao.setF_OPERADORFUND(rs.getString("F_OPERADORFUND"));

                Fundicao.setF_DATALQ(rs.getDate("F_DATALQ"));
                Fundicao.setF_DEFLQ(rs.getString("F_DEFLQ"));
                Fundicao.setF_QTDELQ(rs.getFloat("F_QTDELQ"));
                Fundicao.setF_OPERADORLQ(rs.getString("F_OPERADORLQ"));

                Fundicao.setF_DATASLIT(rs.getDate("F_DATASLIT"));
                Fundicao.setF_DEFSLIT(rs.getString("F_DEFSLIT"));
                Fundicao.setF_QTDESLIT(rs.getFloat("F_QTDESLIT"));
                Fundicao.setF_OPERADORSLIT(rs.getString("F_OPERADORSLIT"));

                Fundicao.setF_DATAEG(rs.getDate("F_DATAEG"));
                Fundicao.setF_ESPMM(rs.getDouble("F_ESPMM"));

                Fundicao.setF_DATALF(rs.getDate("F_DATALF"));
                Fundicao.setF_DEFLF(rs.getString("F_DEFLF"));
                Fundicao.setF_QTDELF(rs.getFloat("F_QTDELF"));
                Fundicao.setF_OPERADORLF(rs.getString("F_OPERADORLF"));

                Fundicao.setF_DATALFF2(rs.getDate("F_DATALFF2"));
                Fundicao.setF_DEFLFF2(rs.getString("F_DEFLFF2"));
                Fundicao.setF_QTDELFF2(rs.getFloat("F_QTDELFF2"));
                Fundicao.setF_OPERADORLFF2(rs.getString("F_OPERADORLFF2"));

                Fundicao.setF_DATALFF1(rs.getDate("F_DATALFF1"));
                Fundicao.setF_DEFLFF1(rs.getString("F_DEFLFF1"));
                Fundicao.setF_QTDELFF1(rs.getFloat("F_QTDELFF1"));
                Fundicao.setF_OPERADORLFF1(rs.getString("F_OPERADORLFF1"));

                Fundicao.setF_DATAE(rs.getDate("F_DATAE"));
                Fundicao.setF_DEFE(rs.getString("F_DEFE"));
                Fundicao.setF_QTDEE(rs.getFloat("F_QTDEE"));
                Fundicao.setF_OPERADORE(rs.getString("F_OPERADORE"));

                Fundicao.setF_DATARLBUHL(rs.getDate("F_DATARLBUHL"));
                Fundicao.setF_DATARLLFF2(rs.getDate("F_DATARLLFF2"));
                Fundicao.setF_DATARLLFF1(rs.getDate("F_DATARLLFF1"));

                Fundicao.setF_DATARLLA9(rs.getDate("F_DATARLLA9"));
                Fundicao.setF_DEFLA9(rs.getString("F_DEFLA9"));
                Fundicao.setF_QTDELA9(rs.getFloat("F_QTDELA9"));
                Fundicao.setF_OPERADORLA9(rs.getString("F_OPERADORLA9"));

                Fundicao.setF_DEFRL(rs.getString("F_DEFRL"));
                Fundicao.setF_QTDERL(rs.getFloat("F_QTDERL"));
                Fundicao.setF_OPERADORRL(rs.getString("F_OPERADORRL"));

                Fundicao.setF_DATAROT(rs.getDate("F_DATAROT"));
                Fundicao.setF_DEFROT(rs.getString("F_DEFROT"));
                Fundicao.setF_QTDEROT(rs.getFloat("F_QTDEROT"));
                Fundicao.setF_OPERADORROT(rs.getString("F_OPERADORROT"));

                Fundicao.setF_DATASF(rs.getDate("F_DATASF"));
                Fundicao.setF_DEFSF(rs.getString("F_DEFSF"));
                Fundicao.setF_QTDESF(rs.getFloat("F_QTDESF"));
                Fundicao.setF_OPERADORSF(rs.getString("F_OPERADORSF"));

                Fundicao.setF_DATASF(rs.getDate("F_DATASF"));
                Fundicao.setF_DEFSF(rs.getString("F_DEFSF"));
                Fundicao.setF_QTDESF(rs.getFloat("F_QTDESF"));
                Fundicao.setF_OPERADORSF(rs.getString("F_OPERADORSF"));

                Fundicao.setF_DATASFI(rs.getDate("F_DATASFI"));
                Fundicao.setF_DEFSFI(rs.getString("F_DEFSFI"));
                Fundicao.setF_QTDESFI(rs.getFloat("F_QTDESFI"));
                Fundicao.setF_OPERADORSFI(rs.getString("F_OPERADORSFI"));

                Fundicao.setF_DATAP(rs.getDate("F_DATAP"));
                Fundicao.setF_DEFP(rs.getString("F_DEFP"));
                Fundicao.setF_QTDEP(rs.getFloat("F_QTDEP"));
                Fundicao.setF_OPERADORP(rs.getString("F_OPERADORP"));

                Fundicao.setF_DATAGUI(rs.getDate("F_DATAGUI"));
                Fundicao.setF_DEFGUI(rs.getString("F_DEFGUI"));
                Fundicao.setF_QTDEGUI(rs.getFloat("F_QTDEGUI"));
                Fundicao.setF_OPERADORGUI(rs.getString("F_OPERADORGUI"));

                Fundicao.setF_DATAGUIF(rs.getDate("F_DATAGUIF"));
                Fundicao.setF_DEFGUIF(rs.getString("F_DEFGUIF"));
                Fundicao.setF_QTDEGUIF(rs.getFloat("F_QTDEGUIF"));
                Fundicao.setF_OPERADORGUIF(rs.getString("F_OPERADORGUIF"));

                Fundicao.setF_DATAR(rs.getDate("F_DATAR"));
                Fundicao.setF_DEFR(rs.getString("F_DEFR"));
                Fundicao.setF_QTDER(rs.getFloat("F_QTDER"));
                Fundicao.setF_OPERADORR(rs.getString("F_OPERADORR"));

                Fundicao.setF_DATAPE(rs.getDate("F_DATAPE"));
                Fundicao.setF_DEFPE(rs.getString("F_DEFPE"));
                Fundicao.setF_QTDEPE(rs.getFloat("F_QTDEPE"));
                Fundicao.setF_OPERADORPE(rs.getString("F_OPERADORPE"));

                Fundicao.setF_DATAC(rs.getDate("F_DATAC"));
                Fundicao.setF_DEFC(rs.getString("F_DEFC"));
                Fundicao.setF_QTDEC(rs.getFloat("F_QTDEC"));
                Fundicao.setF_OPERADORC(rs.getString("F_OPERADORC"));

                Fundicao.setF_DATAESC(rs.getDate("F_DATAESC"));
                Fundicao.setF_DEFESC(rs.getString("F_DEFESC"));
                Fundicao.setF_QTDEESC(rs.getFloat("F_QTDEESC"));
                Fundicao.setF_OPERADORESC(rs.getString("F_OPERADORESC"));

                Fundicao.setF_PRAZOENT(rs.getDate("F_PRAZOENT"));
                Fundicao.setF_KGTEOR(rs.getFloat("F_KGTEOR"));
                Fundicao.setF_QUANT(rs.getFloat("F_QUANT"));
                Fundicao.setF_KGEST(rs.getFloat("F_KGEST"));

                Fundicao.setF_DATAS1(rs.getDate("F_DATAS1"));
                Fundicao.setF_S1(rs.getFloat("F_S1"));
                Fundicao.setF_DATAS2(rs.getDate("F_DATAS2"));
                Fundicao.setF_S2(rs.getFloat("F_S2"));
                Fundicao.setF_DATAS3(rs.getDate("F_DATAS3"));
                Fundicao.setF_S3(rs.getFloat("F_S3"));
                Fundicao.setF_DATAS4(rs.getDate("F_DATAS4"));
                Fundicao.setF_S4(rs.getFloat("F_S4"));
                Fundicao.setF_DATAS5(rs.getDate("F_DATAS5"));
                Fundicao.setF_S5(rs.getFloat("F_S5"));

                Fundicao.setF_DATAS6(rs.getDate("F_DATAS6"));
                Fundicao.setF_S6(rs.getFloat("F_S6"));
                Fundicao.setF_DATAS7(rs.getDate("F_DATAS7"));
                Fundicao.setF_S7(rs.getFloat("F_S7"));
                Fundicao.setF_DATAS8(rs.getDate("F_DATAS8"));
                Fundicao.setF_S8(rs.getFloat("F_S8"));
                Fundicao.setF_DATAS9(rs.getDate("F_DATAS9"));
                Fundicao.setF_S9(rs.getFloat("F_S9"));
                Fundicao.setF_DATAS10(rs.getDate("F_DATAS10"));
                Fundicao.setF_S10(rs.getFloat("F_S10"));

                Fundicao.setF_DATAS11(rs.getDate("F_DATAS11"));
                Fundicao.setF_S11(rs.getFloat("F_S11"));
                Fundicao.setF_DATAS12(rs.getDate("F_DATAS12"));
                Fundicao.setF_S12(rs.getFloat("F_S12"));
                Fundicao.setF_DATAS13(rs.getDate("F_DATAS13"));
                Fundicao.setF_S13(rs.getFloat("F_S13"));
                Fundicao.setF_DATAS14(rs.getDate("F_DATAS14"));
                Fundicao.setF_S14(rs.getFloat("F_S14"));
                Fundicao.setF_DATAS15(rs.getDate("F_DATAS15"));
                Fundicao.setF_S15(rs.getFloat("F_S15"));

                Fundicao.setF_DATAS16(rs.getDate("F_DATAS16"));
                Fundicao.setF_S16(rs.getFloat("F_S16"));
                Fundicao.setF_DATAS17(rs.getDate("F_DATAS17"));
                Fundicao.setF_S17(rs.getFloat("F_S17"));
                Fundicao.setF_DATAS18(rs.getDate("F_DATAS18"));
                Fundicao.setF_S18(rs.getFloat("F_S18"));
                Fundicao.setF_DATAS19(rs.getDate("F_DATAS19"));
                Fundicao.setF_S19(rs.getFloat("F_S19"));
                Fundicao.setF_DATAS20(rs.getDate("F_DATAS20"));
                Fundicao.setF_S20(rs.getFloat("F_S20"));

                Fundicao.setF_TOTAL(rs.getFloat("F_TOTAL"));
                Fundicao.setF_PERC(rs.getFloat("F_PERC"));

                Fundicao.setF_SEQ(rs.getFloat("F_SEQ"));
                Fundicao.setFD_SEQPF(rs.getFloat("FD_SEQPF"));

                //DEFEITOS
                Fundicao.setFD_ESPESSURA(rs.getFloat("FD_ESPESSURA"));
                Fundicao.setFD_OXIDACAO(rs.getFloat("FD_OXIDACAO"));
                Fundicao.setFD_ONDULACAO(rs.getFloat("FD_ONDULACAO"));

                Fundicao.setFD_CASQUINHA(rs.getFloat("FD_CASQUINHA"));
                Fundicao.setFD_RASTRO_TRATOR(rs.getFloat("FD_RASTRO_TRATOR"));
                Fundicao.setFD_MANCHA_OLEO(rs.getFloat("FD_MANCHA_OLEO"));

                Fundicao.setFD_LINHA_DO_LQ(rs.getFloat("FD_LINHA_DO_LQ"));
                Fundicao.setFD_RISCO_DA_PRENSA(rs.getFloat("FD_RISCO_DA_PRENSA"));
                Fundicao.setFD_SUJEIRA(rs.getFloat("FD_SUJEIRA"));

                Fundicao.setFD_NATA(rs.getFloat("FD_NATA"));
                Fundicao.setFD_CAVACO(rs.getFloat("FD_CAVACO"));
                Fundicao.setFD_LINHA_PRETA(rs.getFloat("FD_LINHA_PRETA"));

                Fundicao.setFD_BOLHA_CAB(rs.getFloat("FD_BOLHA_CAB"));
                Fundicao.setFD_BOLHA_TUB(rs.getFloat("FD_BOLHA_TUB"));
                Fundicao.setFD_BOLHA_DES(rs.getFloat("FD_BOLHA_DES"));

                Fundicao.setFD_RISCO(rs.getFloat("FD_RISCO"));
                Fundicao.setFD_QUEBRA_DE_BOB(rs.getFloat("FD_QUEBRA_DE_BOB"));
                Fundicao.setFD_BURACO(rs.getFloat("FD_BURACO"));

                Fundicao.setFD_MANCHA_BRANCA(rs.getFloat("FD_MANCHA_BRANCA"));
                Fundicao.setFD_MANCHA_MARROM(rs.getFloat("FD_MANCHA_MARROM"));

                Fundicao.setFD_BATIMENTO_LATERAL(rs.getFloat("FD_BATIMENTO_LATERAL"));
                Fundicao.setFD_BOLHA_AQUECIMENTO(rs.getFloat("FD_BOLHA_AQUECIMENTO"));
                Fundicao.setFD_GRAO(rs.getFloat("FD_GRAO"));

                Fundicao.setFD_DUREZADEF(rs.getFloat("FD_DUREZADEF"));
                Fundicao.setFD_SOBRA(rs.getFloat("FD_SOBRA"));
                Fundicao.setFD_FLECHA(rs.getFloat("FD_FLECHA"));

                Fundicao.setFD_FALTOU_PESO(rs.getFloat("FD_FALTOU_PESO"));
                Fundicao.setFD_TEMPERATURA(rs.getFloat("FD_TEMPERATURA"));
                Fundicao.setFD_OS_PERDIDA(rs.getFloat("FD_OS_PERDIDA"));

                Fundicao.setFD_PE_DE_PLACA(rs.getFloat("FD_PE_DE_PLACA"));
                Fundicao.setFD_REBARBA(rs.getFloat("FD_REBARBA"));
                Fundicao.setFD_MANCHA_D_AGUA(rs.getFloat("FD_MANCHA_D_AGUA"));

                Fundicao.setFD_OLHO(rs.getFloat("FD_OLHO"));
                Fundicao.setFD_ZINABRO(rs.getFloat("FD_ZINABRO"));
                Fundicao.setFD_DESALINHADA(rs.getFloat("FD_DESALINHADA"));

                Fundicao.setFD_SUMIU(rs.getFloat("FD_SUMIU"));
                Fundicao.setFD_TRINCADA(rs.getFloat("FD_TRINCADA"));
                Fundicao.setFD_BOB_ESTREITA(rs.getFloat("FD_BOB_ESTREITA"));

                Fundicao.setFD_MARCA_CILINDRO(rs.getFloat("FD_MARCA_CILINDRO"));
                Fundicao.setFD_SEM_OS(rs.getFloat("FD_SEM_OS"));
                Fundicao.setFD_LATERAL_AMASSADA(rs.getFloat("FD_LATERAL_AMASSADA"));

                Fundicao.setFD_QUEBRA_PROCED(rs.getFloat("FD_QUEBRA_PROCED"));
                Fundicao.setFD_NFL(rs.getFloat("FD_NFL"));

                Fundicao.setFD_FUND_MED_MENOS(rs.getFloat("FD_FUND_MED_MENOS"));
                Fundicao.setFD_FUND_MED_MAIS(rs.getFloat("FD_FUND_MED_MAIS"));

                Fundicao.setFD_PLACA_BOLHA_PQ(rs.getFloat("FD_PLACA_BOLHA_PQ"));
                Fundicao.setFD_MANCHA_PRETA(rs.getFloat("FD_MANCHA_PRETA"));
                Fundicao.setFD_POZINHO(rs.getFloat("FD_POZINHO"));

                Fundicao.setFD_BOB_FROUCHA(rs.getFloat("FD_BOB_FROUCHA"));
                Fundicao.setFD_BORRA_CHAPA(rs.getFloat("FD_BORRA_CHAPA"));
                Fundicao.setFD_MANCHA_ESTUFA(rs.getFloat("FD_MANCHA_ESTUFA"));

                Fundicao.setFD_MAL_ENROLADA(rs.getFloat("FD_MAL_ENROLADA"));
                Fundicao.setFD_MAT_FALTA_LUB(rs.getFloat("FD_MAT_FALTA_LUB"));
                Fundicao.setFD_EXCESSO_BORRA(rs.getFloat("FD_EXCESSO_BORRA"));

                Fundicao.setFD_BOB_MAL_SOLDADA(rs.getFloat("FD_BOB_MAL_SOLDADA"));
                Fundicao.setFD_FALTA_TESTE_MAT(rs.getFloat("FD_FALTA_TESTE_MAT"));
                Fundicao.setFD_MAT_MARCADO(rs.getFloat("FD_MAT_MARCADO"));

                Fundicao.setFD_EXCESSO_OLEO(rs.getFloat("FD_EXCESSO_OLEO"));
                Fundicao.setFD_FALTA_LIMPEZA(rs.getFloat("FD_FALTA_LIMPEZA"));
                Fundicao.setFD_MATERIAL_FIAPO(rs.getFloat("FD_MATERIAL_FIAPO"));

                Fundicao.setFD_MATERIAL_CHUVISCO(rs.getFloat("FD_MATERIAL_CHUVISCO"));
                Fundicao.setFD_MAT_FORA_ESQUADRO(rs.getFloat("FD_MAT_FORA_ESQUADRO"));
                Fundicao.setFD_MAQUINA_QUEBRADA(rs.getFloat("FD_MAQUINA_QUEBRADA"));

                Fundicao.setFD_FALTA_ATENCAO(rs.getFloat("FD_FALTA_ATENCAO"));

                Fundicao.setFD_ERRO_MEDIDA(rs.getFloat("FD_ERRO_MEDIDA"));
                Fundicao.setFD_FALTA_TESTE_MAQ(rs.getFloat("FD_FALTA_TESTE_MAQ"));
                Fundicao.setFD_NAO_FEZ_MANUT(rs.getFloat("FD_NAO_FEZ_MANUT"));

                Fundicao.setFD_FALTA_REVISAO(rs.getFloat("FD_FALTA_REVISAO"));
                Fundicao.setFD_GASTO_EXCESSIVO(rs.getFloat("FD_GASTO_EXCESSIVO"));
                Fundicao.setFD_MANUT_SEM_EFIC(rs.getFloat("FD_MANUT_SEM_EFIC"));

                Fundicao.setFD_FALTA_MANUT(rs.getFloat("FD_FALTA_MANUT"));
                Fundicao.setFD_PROD_DANIFICADO(rs.getFloat("FD_PROD_DANIFICADO"));
                Fundicao.setFD_FORA_BERCOS(rs.getFloat("FD_FORA_BERCOS"));

                Fundicao.setFD_ATRASOU_SERVICO(rs.getFloat("FD_ATRASOU_SERVICO"));
                Fundicao.setFD_PRODUCAO_BAIXA(rs.getFloat("FD_PRODUCAO_BAIXA"));
                Fundicao.setFD_FALTOU_DIA(rs.getFloat("FD_FALTOU_DIA"));

                Fundicao.setFD_TROCA_PAPELETA(rs.getFloat("FD_TROCA_PAPELETA"));
                Fundicao.setFD_FALTA_DISCIPLINA(rs.getFloat("FD_FALTA_DISCIPLINA"));
                Fundicao.setFD_SEM_PROT_AURICULAR(rs.getFloat("FD_SEM_PROT_AURICULAR"));

                Fundicao.setFD_SEM_OCULOS(rs.getFloat("FD_SEM_OCULOS"));
                Fundicao.setFD_SEM_LUVA(rs.getFloat("FD_SEM_LUVA"));
                Fundicao.setFD_SEM_SAPATO_PROT(rs.getFloat("FD_SEM_SAPATO_PROT"));

                Fundicao.setFD_CORREDOR_OBST(rs.getFloat("FD_CORREDOR_OBST"));
                Fundicao.setFD_QUEBRA_NORMA(rs.getFloat("FD_QUEBRA_NORMA"));

                Fundicao.setFD_MAT_TEMP_ALTA(rs.getFloat("FD_MAT_TEMP_ALTA"));
                Fundicao.setFD_MAT_TEMP_BAIXA(rs.getFloat("FD_MAT_TEMP_BAIXA"));

                Fundicao.setFD_RISCO_LQ(rs.getFloat("FD_RISCO_LQ"));
                Fundicao.setFD_RISCO_SG(rs.getFloat("FD_RISCO_SG"));
                Fundicao.setFD_RISCO_LF(rs.getFloat("FD_RISCO_LF"));
                Fundicao.setFD_RISCO_SF(rs.getFloat("FD_RISCO_SF"));

                Fundicao.setFD_PEDIDO_CANCEL(rs.getFloat("FD_PEDIDO_CANCEL"));
                Fundicao.setFD_BOB_DESVIADA(rs.getFloat("FD_BOB_DESVIADA"));

                Fundicao.setFD_MEIA_LUA(rs.getFloat("FD_MEIA_LUA"));
                Fundicao.setFD_REFILE_MAIOR(rs.getFloat("FD_REFILE_MAIOR"));
                Fundicao.setFD_ERRO_DE_MEDIA(rs.getFloat("FD_ERRO_DE_MEDIA"));

                Fundicao.setDATA_LEVANTAMENTO(rs.getDate("DATA_LEVANTAMENTO"));
                Fundicao.setCULPADO(rs.getString("CULPADO"));
                Fundicao.setNUMDISC(rs.getFloat("NUMDISC"));
       
                arrayList.add (Fundicao);
}
//DETERMINA A LINHA COM AS CELULAS COM AS INFORMAÇÔES DO BANCO DE DADOS
int rowNum = 4;
            for (fundicao obj : arrayList) {
                Row dataRow = sheet.createRow(rowNum++);

                Cell F_PEDIDO = dataRow.createCell(1);
                F_PEDIDO.setCellValue(obj.getF_PEDIDO());
                F_PEDIDO.setCellStyle(alignDataStyle);

                Cell F_EMISSAO = dataRow.createCell(2);
                F_EMISSAO.setCellValue(obj.getF_EMISSAO());
                F_EMISSAO.setCellStyle(dateCellStyle);

                Cell F_DATAAPROVFK = dataRow.createCell(3);
                F_DATAAPROVFK.setCellValue(obj.getF_DATAAPROVFK());
                F_DATAAPROVFK.setCellStyle(dateCellStyle);

                Cell F_DATAISO = dataRow.createCell(4);
                F_DATAISO.setCellValue(obj.getF_DATAISO());
                F_DATAISO.setCellStyle(dateCellStyle);

                Cell F_DATAPEDIDO = dataRow.createCell(5);
                F_DATAPEDIDO.setCellValue(obj.getF_DATAPEDIDO());
                F_DATAPEDIDO.setCellStyle(dateCellStyle);

                Cell F_DATAOS = dataRow.createCell(6);
                F_DATAOS.setCellValue(obj.getF_DATAOS());
                F_DATAOS.setCellStyle(dateCellStyle);

                Cell F_PRAZOFUNDIR = dataRow.createCell(7);
                F_PRAZOFUNDIR.setCellValue(obj.getF_PRAZOFUNDIR());
                F_PRAZOFUNDIR.setCellStyle(dateCellStyle);

                Cell F_NUMPLACASMAIOR90 = dataRow.createCell(8);
                F_NUMPLACASMAIOR90.setCellValue(obj.getF_NUMPLACASMAIOR90());
                F_NUMPLACASMAIOR90.setCellStyle(alignDataStyle);

                Cell F_DATAENVFUND = dataRow.createCell(9);
                F_DATAENVFUND.setCellValue(obj.getF_DATAENVFUND());
                F_DATAENVFUND.setCellStyle(dateCellStyle);

                Cell F_LOTENUM = dataRow.createCell(10);
                F_LOTENUM.setCellValue(obj.getF_LOTENUM());
                F_LOTENUM.setCellStyle(alignDataStyle);

                Cell F_MEDIDA = dataRow.createCell(11);
                F_MEDIDA.setCellValue(obj.getF_MEDIDA());
                F_MEDIDA.setCellStyle(alignDataStyle);
                sheet.addMergedRegion(CellRangeAddress.valueOf("L" + rowNum + ":P" + rowNum + ""));

                Cell F_DUREZA = dataRow.createCell(16);
                F_DUREZA.setCellValue(obj.getF_DUREZA());
                F_DUREZA.setCellStyle(alignDataStyle);
                sheet.addMergedRegion(CellRangeAddress.valueOf("Q" + rowNum + ":R" + rowNum + ""));

                Cell F_OSNUM = dataRow.createCell(18);
                F_OSNUM.setCellValue(obj.getF_OSNUM());
                F_OSNUM.setCellStyle(alignDataStyle);

                Cell F_CLIENTE = dataRow.createCell(19);
                F_CLIENTE.setCellValue(obj.getF_CLIENTE());
                F_CLIENTE.setCellStyle(alignDataStyle);

                Cell F_F0 = dataRow.createCell(20);
                F_F0.setCellValue(obj.getF_F0());
                F_F0.setCellStyle(alignDataStyle);

                Cell F_F1 = dataRow.createCell(21);
                F_F1.setCellValue(obj.getF_F1());
                F_F1.setCellStyle(alignDataStyle);

                Cell F_F2 = dataRow.createCell(22);
                F_F2.setCellValue(obj.getF_F2());
                F_F2.setCellStyle(alignDataStyle);

                Cell F_F3 = dataRow.createCell(23);
                F_F3.setCellValue(obj.getF_F3());
                F_F3.setCellStyle(alignDataStyle);

                Cell F_SUPER = dataRow.createCell(24);
                F_SUPER.setCellValue(obj.getF_SUPER());
                F_SUPER.setCellStyle(alignDataStyle);

                Cell F_3105 = dataRow.createCell(25); //F1-27
                F_3105.setCellValue(obj.getF_3105());
                F_3105.setCellStyle(alignDataStyle);

                Cell outrasLigas = dataRow.createCell(26);
                outrasLigas.setCellValue("");
                outrasLigas.setCellStyle(alignDataStyle);

                Cell F_PLACANUM = dataRow.createCell(27);
                F_PLACANUM.setCellValue(obj.getF_PLACANUM());
                F_PLACANUM.setCellStyle(alignDataStyle);

                Cell F_LARGREALOS = dataRow.createCell(28);
                F_LARGREALOS.setCellValue(obj.getF_LARGREALOS());
                F_LARGREALOS.setCellStyle(alignDataStyle);

                Cell F_KGCALC = dataRow.createCell(29);
                F_KGCALC.setCellValue(obj.getF_KGCALC());
                F_KGCALC.setCellStyle(numCellStyle);

                Cell F_LARGFUND = dataRow.createCell(30);
                F_LARGFUND.setCellValue(obj.getF_LARGFUND());
                F_LARGFUND.setCellStyle(alignDataStyle);

                Cell F_NUMBOBINA = dataRow.createCell(31);
                F_NUMBOBINA.setCellValue(obj.getF_NUMBOBINA());
                F_NUMBOBINA.setCellStyle(alignDataStyle);

                Cell F_DURSOLICOS = dataRow.createCell(32);
                F_DURSOLICOS.setCellValue(obj.getF_DURSOLICOS());
                F_DURSOLICOS.setCellStyle(alignDataStyle);

                Cell grao = dataRow.createCell(33);
                grao.setCellStyle(alignDataStyle);
                grao.setCellFormula("IF(OR(K" + rowNum + " =\"\", K" + rowNum + "=\"FUNDIR\"),\"AGUARDAR\",\"PREENCHER\")");

                Cell durezaH = dataRow.createCell(34);
                durezaH.setCellStyle(alignDataStyle);
                durezaH.setCellFormula("IF(OR(K" + rowNum + " =\"\", K" + rowNum + "=\"FUNDIR\"),\"AGUARDAR\",\"PREENCHER\")");

                Cell F_SETOR = dataRow.createCell(35);
                F_SETOR.setCellValue(obj.getF_SETOR());
                F_SETOR.setCellStyle(alignDataStyle);

                Cell F_ATENDENTE = dataRow.createCell(36);
                F_ATENDENTE.setCellValue(obj.getF_ATENDENTE());
                F_ATENDENTE.setCellStyle(alignDataStyle);

                Cell F_TRATADOR = dataRow.createCell(37);
                F_TRATADOR.setCellValue(obj.getF_TRATADOR());
                F_TRATADOR.setCellStyle(alignDataStyle);

                Cell F_DATAFUND = dataRow.createCell(38);
                F_DATAFUND.setCellValue(obj.getF_DATAFUND());
                F_DATAFUND.setCellStyle(dateCellStyle);

                Cell DATA_LEVANTAMENTO = dataRow.createCell(39);
                DATA_LEVANTAMENTO.setCellValue(obj.getDATA_LEVANTAMENTO());
                DATA_LEVANTAMENTO.setCellStyle(dateCellStyle);

                Cell NUMDISC = dataRow.createCell(40);
                NUMDISC.setCellValue(obj.getNUMDISC());
                NUMDISC.setCellStyle(alignDataStyle);

                Cell CULPADO = dataRow.createCell(41);
                CULPADO.setCellValue(obj.getCULPADO());
                CULPADO.setCellStyle(alignDataStyle);

                Cell F_DEFFUND = dataRow.createCell(42);
                F_DEFFUND.setCellValue(obj.getF_DEFFUND());
                F_DEFFUND.setCellStyle(alignDataStyle);

                Cell F_QTDEFUND = dataRow.createCell(43);
                F_QTDEFUND.setCellValue(obj.getF_DEFFUND());
                F_QTDEFUND.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEFUND()+")");
                F_QTDEFUND.setCellStyle(numCellStyle);

                Cell F_OPERADORFUND = dataRow.createCell(44);
                F_OPERADORFUND.setCellValue(obj.getF_OPERADORFUND());
                F_OPERADORFUND.setCellStyle(alignDataStyle);

                Cell EscolaFund = dataRow.createCell(45);
                EscolaFund.setCellValue("Escolinha");
                EscolaFund.setCellStyle(alignDataStyle);

                Cell F_DATALQ = dataRow.createCell(46);
                F_DATALQ.setCellValue(obj.getF_DATALQ());
                F_DATALQ.setCellStyle(dateCellStyle);

                Cell F_DEFLQ = dataRow.createCell(47);
                F_DEFLQ.setCellValue(obj.getF_DEFLQ());
                F_DEFLQ.setCellStyle(dateCellStyle);

                Cell F_QTDELQ = dataRow.createCell(48);
                F_QTDELQ.setCellValue(obj.getF_QTDELQ());
                F_QTDELQ.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDELQ()+")");
                F_QTDELQ.setCellStyle(numCellStyle);

                Cell F_OPERADORLQ = dataRow.createCell(49);
                F_OPERADORLQ.setCellValue(obj.getF_OPERADORLQ());
                F_OPERADORLQ.setCellStyle(alignDataStyle);

                Cell EscolaLq = dataRow.createCell(50);
                EscolaLq.setCellValue("Escolinha");
                EscolaLq.setCellStyle(alignDataStyle);

                Cell F_DATASLIT = dataRow.createCell(51);
                F_DATASLIT.setCellValue(obj.getF_DATASLIT());
                F_DATASLIT.setCellStyle(dateCellStyle);

                Cell F_DEFSLIT = dataRow.createCell(52);
                F_DEFSLIT.setCellValue(obj.getF_DEFSLIT());
                F_DEFSLIT.setCellStyle(dateCellStyle);

                Cell F_QTDESLIT = dataRow.createCell(53);
                F_QTDESLIT.setCellValue(obj.getF_QTDESLIT());
                F_QTDESLIT.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDESLIT()+")");
                F_QTDESLIT.setCellStyle(numCellStyle);

                Cell F_OPERADORSLIT = dataRow.createCell(54);
                F_OPERADORSLIT.setCellValue(obj.getF_OPERADORSLIT());
                F_OPERADORSLIT.setCellStyle(alignDataStyle);

                Cell EscolaSlit = dataRow.createCell(55);
                EscolaSlit.setCellValue("Escolinha");
                EscolaSlit.setCellStyle(alignDataStyle);

                Cell F_DATAEG = dataRow.createCell(56);
                F_DATAEG.setCellValue(obj.getF_DATAEG());
                F_DATAEG.setCellStyle(dateCellStyle);

                Cell F_ESPMM = dataRow.createCell(57);
                F_ESPMM.setCellValue(obj.getF_ESPMM());
                F_ESPMM.setCellStyle(alignDataStyle);

                Cell F_DATALF = dataRow.createCell(58);
                F_DATALF.setCellValue(obj.getF_DATALF());
                F_DATALF.setCellStyle(dateCellStyle);

                Cell F_DEFLF = dataRow.createCell(59);
                F_DEFLF.setCellValue(obj.getF_DEFLF());
                F_DEFLF.setCellStyle(dateCellStyle);

                Cell F_QTDELF = dataRow.createCell(60);
                F_QTDELF.setCellValue(obj.getF_QTDELF());
                F_QTDELF.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDELF()+")");
                F_QTDELF.setCellStyle(numCellStyle);

                Cell F_OPERADORLF = dataRow.createCell(61);
                F_OPERADORLF.setCellValue(obj.getF_OPERADORLF());
                F_OPERADORLF.setCellStyle(alignDataStyle);

                Cell EscolaLf = dataRow.createCell(62);
                EscolaLf.setCellValue("Escolinha");
                EscolaLf.setCellStyle(alignDataStyle);

                Cell F_DATALFF2 = dataRow.createCell(63);
                F_DATALFF2.setCellValue(obj.getF_DATALFF2());
                F_DATALFF2.setCellStyle(dateCellStyle);

                Cell F_DEFLFF2 = dataRow.createCell(64);
                F_DEFLFF2.setCellValue(obj.getF_DEFLFF2());
                F_DEFLFF2.setCellStyle(dateCellStyle);

                Cell F_QTDELFF2 = dataRow.createCell(65);
                F_QTDELFF2.setCellValue(obj.getF_QTDELFF2());
                F_QTDELFF2.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDELFF2()+")");
                F_QTDELFF2.setCellStyle(numCellStyle);

                Cell F_OPERADORLFF2 = dataRow.createCell(66);
                F_OPERADORLFF2.setCellValue(obj.getF_OPERADORLFF2());
                F_OPERADORLFF2.setCellStyle(alignDataStyle);

                Cell EscolaLff2 = dataRow.createCell(67);
                EscolaLff2.setCellValue("Escolinha");
                EscolaLff2.setCellStyle(alignDataStyle);

                Cell F_DATALFF1 = dataRow.createCell(68);
                F_DATALFF1.setCellValue(obj.getF_DATALFF1());
                F_DATALFF1.setCellStyle(dateCellStyle);

                Cell F_DEFLFF1 = dataRow.createCell(69);
                F_DEFLFF1.setCellValue(obj.getF_DEFLFF1());
                F_DEFLFF1.setCellStyle(dateCellStyle);

                Cell F_QTDELFF1 = dataRow.createCell(70);
                F_QTDELFF1.setCellValue(obj.getF_QTDELFF1());
                F_QTDELFF1.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDELFF1()+")");
                F_QTDELFF1.setCellStyle(numCellStyle);

                Cell F_OPERADORLFF1 = dataRow.createCell(71);
                F_OPERADORLFF1.setCellValue(obj.getF_OPERADORLFF1());
                F_OPERADORLFF1.setCellStyle(alignDataStyle);

                Cell EscolaLff1 = dataRow.createCell(72);
                EscolaLff1.setCellValue("Escolinha");
                EscolaLff1.setCellStyle(alignDataStyle);

                Cell F_DATAE = dataRow.createCell(73);
                F_DATAE.setCellValue(obj.getF_DATAE());
                F_DATAE.setCellStyle(dateCellStyle);

                Cell F_DEFE = dataRow.createCell(74);
                F_DEFE.setCellValue(obj.getF_DEFE());
                F_DEFE.setCellStyle(dateCellStyle);

                Cell F_QTDEE = dataRow.createCell(75);
                F_QTDEE.setCellValue(obj.getF_QTDEE());
                F_QTDEE.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEE()+")");
                F_QTDEE.setCellStyle(numCellStyle);

                Cell F_OPERADORE = dataRow.createCell(76);
                F_OPERADORE.setCellValue(obj.getF_OPERADORE());
                F_OPERADORE.setCellStyle(alignDataStyle);

                Cell EscolaE = dataRow.createCell(77);
                EscolaE.setCellValue("Escolinha");
                EscolaE.setCellStyle(alignDataStyle);

                Cell F_DATARLBUHL = dataRow.createCell(78);
                F_DATARLBUHL.setCellValue(obj.getF_DATARLBUHL());
                F_DATARLBUHL.setCellStyle(dateCellStyle);

                Cell F_DATARLLFF2 = dataRow.createCell(79);
                F_DATARLLFF2.setCellValue(obj.getF_DATARLLFF2());
                F_DATARLLFF2.setCellStyle(dateCellStyle);

                Cell F_DATARLLFF1 = dataRow.createCell(80);
                F_DATARLLFF1.setCellValue(obj.getF_DATARLLFF1());
                F_DATARLLFF1.setCellStyle(dateCellStyle);

                Cell F_DATARLLA9 = dataRow.createCell(81);
                F_DATARLLA9.setCellValue(obj.getF_DATARLLA9());
                F_DATARLLA9.setCellStyle(dateCellStyle);

                Cell F_DEFRL = dataRow.createCell(82);
                F_DEFRL.setCellValue(obj.getF_DEFRL());
                F_DEFRL.setCellStyle(dateCellStyle);

                Cell F_QTDERL = dataRow.createCell(83);
                F_QTDERL.setCellValue(obj.getF_QTDERL());
                F_QTDERL.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDERL()+")");
                F_QTDERL.setCellStyle(numCellStyle);

                Cell F_OPERADORRL = dataRow.createCell(84);
                F_OPERADORRL.setCellValue(obj.getF_OPERADORRL());
                F_OPERADORRL.setCellStyle(alignDataStyle);

                Cell EscolaRl = dataRow.createCell(85);
                EscolaRl.setCellValue("Escolinha");
                EscolaRl.setCellStyle(alignDataStyle);

                Cell getF_DATALA9 = dataRow.createCell(86);
                getF_DATALA9.setCellValue(obj.getF_DATALA9());
                getF_DATALA9.setCellStyle(dateCellStyle);

                Cell F_DEFLA9 = dataRow.createCell(87);
                F_DEFLA9.setCellValue(obj.getF_DEFLA9());
                F_DEFLA9.setCellStyle(dateCellStyle);

                Cell F_QTDELA9 = dataRow.createCell(88);
                F_QTDELA9.setCellValue(obj.getF_QTDELA9());
                F_QTDELA9.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDELA9()+")");
                F_QTDELA9.setCellStyle(numCellStyle);

                Cell F_OPERADORLA9 = dataRow.createCell(89);
                F_OPERADORLA9.setCellValue(obj.getF_OPERADORLA9());
                F_OPERADORLA9.setCellStyle(alignDataStyle);

                Cell EscolaLa9 = dataRow.createCell(90);
                EscolaLa9.setCellValue("Escolinha");
                EscolaLa9.setCellStyle(alignDataStyle);

                Cell F_DATAROT = dataRow.createCell(91);
                F_DATAROT.setCellValue(obj.getF_DATAROT());
                F_DATAROT.setCellStyle(dateCellStyle);

                Cell F_DEFROT = dataRow.createCell(92);
                F_DEFROT.setCellValue(obj.getF_DEFROT());
                F_DEFROT.setCellStyle(dateCellStyle);

                Cell F_QTDEROT = dataRow.createCell(93);
                F_QTDEROT.setCellValue(obj.getF_QTDEROT());
                F_QTDEROT.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEROT()+")");
                F_QTDEROT.setCellStyle(numCellStyle);

                Cell F_OPERADORROT = dataRow.createCell(94);
                F_OPERADORROT.setCellValue(obj.getF_OPERADORROT());
                F_OPERADORROT.setCellStyle(alignDataStyle);

                Cell EscolaRot = dataRow.createCell(95);
                EscolaRot.setCellValue("Escolinha");
                EscolaRot.setCellStyle(alignDataStyle);

                Cell F_DATASF = dataRow.createCell(96);
                F_DATASF.setCellValue(obj.getF_DATASF());
                F_DATASF.setCellStyle(dateCellStyle);

                Cell F_DEFSF = dataRow.createCell(97);
                F_DEFSF.setCellValue(obj.getF_DEFSF());
                F_DEFSF.setCellStyle(dateCellStyle);

                Cell F_QTDESF = dataRow.createCell(98);
                F_QTDESF.setCellValue(obj.getF_QTDESF());
                F_QTDESF.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDESF()+")");
                F_QTDESF.setCellStyle(numCellStyle);

                Cell F_OPERADORSF = dataRow.createCell(99);
                F_OPERADORSF.setCellValue(obj.getF_OPERADORSF());
                F_OPERADORSF.setCellStyle(alignDataStyle);

                Cell EscolaSf = dataRow.createCell(100);
                EscolaSf.setCellValue("Escolinha");
                EscolaSf.setCellStyle(alignDataStyle);
//
                Cell F_DATASFI = dataRow.createCell(101);
                F_DATASFI.setCellValue(obj.getF_DATASFI());
                F_DATASFI.setCellStyle(dateCellStyle);

                Cell F_DEFSFI = dataRow.createCell(102);
                F_DEFSFI.setCellValue(obj.getF_DEFSFI());
                F_DEFSFI.setCellStyle(dateCellStyle);

                Cell F_QTDESFI = dataRow.createCell(103);
                F_QTDESFI.setCellValue(obj.getF_QTDESFI());
                F_QTDESFI.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDESFI()+")");
                F_QTDESFI.setCellStyle(numCellStyle);

                Cell F_OPERADORSFI = dataRow.createCell(104);
                F_OPERADORSFI.setCellValue(obj.getF_OPERADORSFI());
                F_OPERADORSFI.setCellStyle(alignDataStyle);

                Cell EscolaSfi = dataRow.createCell(105);
                EscolaSfi.setCellValue("Escolinha");
                EscolaSfi.setCellStyle(alignDataStyle);

                Cell F_DATAP = dataRow.createCell(106);
                F_DATAP.setCellValue(obj.getF_DATAP());
                F_DATAP.setCellStyle(dateCellStyle);

                Cell F_DEFP = dataRow.createCell(107);
                F_DEFP.setCellValue(obj.getF_DEFP());
                F_DEFP.setCellStyle(dateCellStyle);

                Cell F_QTDEP = dataRow.createCell(108);
                F_QTDEP.setCellValue(obj.getF_QTDEP());
                F_QTDEP.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEP()+")");
                F_QTDEP.setCellStyle(numCellStyle);

                Cell F_OPERADORP = dataRow.createCell(109);
                F_OPERADORP.setCellValue(obj.getF_OPERADORP());
                F_OPERADORP.setCellStyle(alignDataStyle);

                Cell EscolaP = dataRow.createCell(110);
                EscolaP.setCellValue("Escolinha");
                EscolaP.setCellStyle(alignDataStyle);

                Cell F_DATAGUI = dataRow.createCell(111);
                F_DATAGUI.setCellValue(obj.getF_DATAGUI());
                F_DATAGUI.setCellStyle(dateCellStyle);

                Cell F_DEFGUI = dataRow.createCell(112);
                F_DEFGUI.setCellValue(obj.getF_DEFGUI());
                F_DEFGUI.setCellStyle(dateCellStyle);

                Cell F_QTDEGUI = dataRow.createCell(113);
                F_QTDEGUI.setCellValue(obj.getF_QTDEGUI());
                F_QTDEGUI.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEGUI()+")");
                F_QTDEGUI.setCellStyle(numCellStyle);

                Cell F_OPERADORGUI = dataRow.createCell(114);
                F_OPERADORGUI.setCellValue(obj.getF_OPERADORGUI());
                F_OPERADORGUI.setCellStyle(alignDataStyle);

                Cell EscolaGui = dataRow.createCell(115);
                EscolaGui.setCellValue("Escolinha");
                EscolaGui.setCellStyle(alignDataStyle);

                Cell F_DATAGUIF = dataRow.createCell(116);
                F_DATAGUIF.setCellValue(obj.getF_DATAGUIF());
                F_DATAGUIF.setCellStyle(dateCellStyle);

                Cell F_DEFGUIF = dataRow.createCell(117);
                F_DEFGUIF.setCellValue(obj.getF_DEFGUIF());
                F_DEFGUIF.setCellStyle(dateCellStyle);

                Cell F_QTDEGUIF = dataRow.createCell(118);
                F_QTDEGUIF.setCellValue(obj.getF_QTDEGUIF());
                F_QTDEGUIF.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEGUIF()+")");
                F_QTDEGUIF.setCellStyle(numCellStyle);

                Cell F_OPERADORGUIF = dataRow.createCell(119);
                F_OPERADORGUIF.setCellValue(obj.getF_OPERADORGUIF());
                F_OPERADORGUIF.setCellStyle(alignDataStyle);

                Cell EscolaGuiF = dataRow.createCell(120);
                EscolaGuiF.setCellValue("Escolinha");
                EscolaGuiF.setCellStyle(alignDataStyle);

                Cell F_DATAR = dataRow.createCell(121);
                F_DATAR.setCellValue(obj.getF_DATAR());
                F_DATAR.setCellStyle(dateCellStyle);

                Cell F_DEFR = dataRow.createCell(122);
                F_DEFR.setCellValue(obj.getF_DEFR());
                F_DEFR.setCellStyle(dateCellStyle);

                Cell F_QTDER = dataRow.createCell(123);
                F_QTDER.setCellValue(obj.getF_QTDER());
                F_QTDER.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDER()+")");
                F_QTDER.setCellStyle(numCellStyle);

                Cell F_OPERADORR = dataRow.createCell(124);
                F_OPERADORR.setCellValue(obj.getF_OPERADORR());
                F_OPERADORR.setCellStyle(alignDataStyle);

                Cell EscolaR = dataRow.createCell(125);
                EscolaR.setCellValue("Escolinha");
                EscolaR.setCellStyle(alignDataStyle);

                Cell F_DATAPE = dataRow.createCell(126);
                F_DATAPE.setCellValue(obj.getF_DATAPE());
                F_DATAPE.setCellStyle(dateCellStyle);

                Cell F_DEFPE = dataRow.createCell(127);
                F_DEFPE.setCellValue(obj.getF_DEFPE());
                F_DEFPE.setCellStyle(dateCellStyle);

                Cell F_QTDEPE = dataRow.createCell(128);
                F_QTDEPE.setCellValue(obj.getF_QTDEPE());
                F_QTDEPE.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEPE()+")");
                F_QTDEPE.setCellStyle(numCellStyle);

                Cell F_OPERADORPE = dataRow.createCell(129);
                F_OPERADORPE.setCellValue(obj.getF_OPERADORPE());
                F_OPERADORPE.setCellStyle(alignDataStyle);

                Cell EscolaPe = dataRow.createCell(130);
                EscolaPe.setCellValue("Escolinha");
                EscolaPe.setCellStyle(alignDataStyle);

                Cell F_DATAC = dataRow.createCell(131);
                F_DATAC.setCellValue(obj.getF_DATAC());
                F_DATAC.setCellStyle(dateCellStyle);

                Cell F_DEFC = dataRow.createCell(132);
                F_DEFC.setCellValue(obj.getF_DEFC());
                F_DEFC.setCellStyle(dateCellStyle);

                Cell F_QTDEC = dataRow.createCell(133);
                F_QTDEC.setCellValue(obj.getF_QTDEC());
                F_QTDEC.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEC()+")");
                F_QTDEC.setCellStyle(numCellStyle);

                Cell F_OPERADORC = dataRow.createCell(134);
                F_OPERADORC.setCellValue(obj.getF_OPERADORC());
                F_OPERADORC.setCellStyle(alignDataStyle);

                Cell EscolaC = dataRow.createCell(135);
                EscolaC.setCellValue("Escolinha");
                EscolaC.setCellStyle(alignDataStyle);

                Cell F_DATAESC = dataRow.createCell(136);
                F_DATAESC.setCellValue(obj.getF_DATAESC());
                F_DATAESC.setCellStyle(dateCellStyle);

                Cell realizado = dataRow.createCell(137);
                realizado.setCellValue("");
                realizado.setCellStyle(alignDataStyle);

                Cell F_DEFESC = dataRow.createCell(138);
                F_DEFESC.setCellValue(obj.getF_DEFESC());
                F_DEFESC.setCellStyle(dateCellStyle);

                Cell F_QTDEESC = dataRow.createCell(139);
                F_QTDEESC.setCellValue(obj.getF_QTDEESC());
                F_QTDEESC.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_QTDEESC()+")");
                F_QTDEESC.setCellStyle(numCellStyle);

                Cell F_OPERADORESC = dataRow.createCell(140);
                F_OPERADORESC.setCellValue(obj.getF_OPERADORESC());
                F_OPERADORESC.setCellStyle(alignDataStyle);

                Cell EscolaEsc = dataRow.createCell(141);
                EscolaEsc.setCellValue("Escolinha");
                EscolaEsc.setCellStyle(alignDataStyle);

                Cell embarque = dataRow.createCell(142);
                embarque.setCellValue("");
                embarque.setCellStyle(alignDataStyle);

                Cell manhaTarde = dataRow.createCell(143);
                manhaTarde.setCellValue("");
                manhaTarde.setCellStyle(alignDataStyle);

                Cell realizado1 = dataRow.createCell(144);
                realizado1.setCellValue("");
                realizado1.setCellStyle(alignDataStyle);

                Cell regiao = dataRow.createCell(145);
                regiao.setCellValue("");
                regiao.setCellStyle(alignDataStyle);

                Cell medio = dataRow.createCell(146);
                medio.setCellValue("");
                medio.setCellStyle(alignDataStyle);

                Cell grd = dataRow.createCell(147);
                grd.setCellValue("");
                grd.setCellStyle(alignDataStyle);

                Cell bebe = dataRow.createCell(148);
                bebe.setCellValue("");
                bebe.setCellStyle(alignDataStyle);

                Cell F_PRAZOENT = dataRow.createCell(149);
                F_PRAZOENT.setCellValue(obj.getF_PRAZOENT());
                F_PRAZOENT.setCellStyle(dateCellStyle);

                Cell F_KGTEOR = dataRow.createCell(150);
                F_KGTEOR.setCellValue(obj.getF_KGTEOR());
                F_KGTEOR.setCellStyle(alignDataStyle);
                float float1 = 3.64f;
                float float2 = 0.57f;
                float float3 = 0.7f;
                F_KGTEOR.setCellFormula("IF(L" + rowNum + " =\"\",(AC" + rowNum + "*" + float1 + ")*AB" + rowNum + "*" + float2 + ",(AC" + rowNum + "*" + float1 + ")*AB" + rowNum + "*" + float3 + ")");

                Cell F_QUANT = dataRow.createCell(151);
                F_QUANT.setCellValue(obj.getF_QUANT());
                F_QUANT.setCellStyle(numCellStyle);

                Cell F_KGEST = dataRow.createCell(152);
                F_KGEST.setCellValue(obj.getF_KGEST());
                F_KGEST.setCellStyle(numCellStyle);
                F_KGEST.setCellFormula("EU" + rowNum + "-EV" + rowNum);

                Cell F_DATAS1 = dataRow.createCell(153);
                F_DATAS1.setCellValue(obj.getF_DATAS1());
                F_DATAS1.setCellStyle(dateCellStyle);

                Cell F_S1 = dataRow.createCell(154);
                F_S1.setCellValue(obj.getF_S1());
                F_S1.setCellFormula("IF(ISNA(AND(VLOOKUP(K"+rowNum+",K"+(rowNum-1)+",1,TRUE),VLOOKUP(S"+rowNum+",S"+(rowNum-1)+",1,TRUE),VLOOKUP(AF"+rowNum+",AF"+(rowNum-1)+",1,TRUE))= TRUE),\"0,00\", "+obj.getF_S1()+")");
                F_S1.setCellStyle(numCellStyle);

                Cell F_DATAS2 = dataRow.createCell(155);
                F_DATAS2.setCellValue(obj.getF_DATAS2());
                F_DATAS2.setCellStyle(dateCellStyle);

                Cell F_S2 = dataRow.createCell(156);
                F_S2.setCellValue(obj.getF_S2());
                F_S2.setCellStyle(numCellStyle);

                Cell F_DATAS3 = dataRow.createCell(157);
                F_DATAS3.setCellValue(obj.getF_DATAS3());
                F_DATAS3.setCellStyle(dateCellStyle);

                Cell F_S3 = dataRow.createCell(158);
                F_S3.setCellValue(obj.getF_S3());
                F_S3.setCellStyle(numCellStyle);

                Cell F_DATAS4 = dataRow.createCell(159);
                F_DATAS4.setCellValue(obj.getF_DATAS4());
                F_DATAS4.setCellStyle(dateCellStyle);

                Cell F_S4 = dataRow.createCell(160);
                F_S4.setCellValue(obj.getF_S4());
                F_S4.setCellStyle(numCellStyle);

                Cell F_DATAS5 = dataRow.createCell(161);
                F_DATAS5.setCellValue(obj.getF_DATAS5());
                F_DATAS5.setCellStyle(dateCellStyle);

                Cell F_S5 = dataRow.createCell(162);
                F_S5.setCellValue(obj.getF_S5());
                F_S5.setCellStyle(numCellStyle);

                Cell F_DATAS6 = dataRow.createCell(163);
                F_DATAS6.setCellValue(obj.getF_DATAS6());
                F_DATAS6.setCellStyle(dateCellStyle);

                Cell F_S6 = dataRow.createCell(164);
                F_S6.setCellValue(obj.getF_S6());
                F_S6.setCellStyle(numCellStyle);

                Cell F_DATAS7 = dataRow.createCell(165);
                F_DATAS7.setCellValue(obj.getF_DATAS7());
                F_DATAS7.setCellStyle(dateCellStyle);

                Cell F_S7 = dataRow.createCell(166);
                F_S7.setCellValue(obj.getF_S7());
                F_S7.setCellStyle(numCellStyle);

                Cell F_DATAS8 = dataRow.createCell(167);
                F_DATAS8.setCellValue(obj.getF_DATAS8());
                F_DATAS8.setCellStyle(dateCellStyle);

                Cell F_S8 = dataRow.createCell(168);
                F_S8.setCellValue(obj.getF_S8());
                F_S8.setCellStyle(numCellStyle);

                Cell F_DATAS9 = dataRow.createCell(169);
                F_DATAS9.setCellValue(obj.getF_DATAS9());
                F_DATAS9.setCellStyle(dateCellStyle);

                Cell F_S9 = dataRow.createCell(170);
                F_S9.setCellValue(obj.getF_S9());
                F_S9.setCellStyle(numCellStyle);

                Cell F_DATAS10 = dataRow.createCell(171);
                F_DATAS10.setCellValue(obj.getF_DATAS10());
                F_DATAS10.setCellStyle(dateCellStyle);

                Cell F_S10 = dataRow.createCell(172);
                F_S10.setCellValue(obj.getF_S10());
                F_S10.setCellStyle(numCellStyle);

                Cell F_DATAS11 = dataRow.createCell(173);
                F_DATAS11.setCellValue(obj.getF_DATAS11());
                F_DATAS11.setCellStyle(dateCellStyle);

                Cell F_S11 = dataRow.createCell(174);
                F_S11.setCellValue(obj.getF_S11());
                F_S11.setCellStyle(numCellStyle);

                Cell F_DATAS12 = dataRow.createCell(175);
                F_DATAS12.setCellValue(obj.getF_DATAS12());
                F_DATAS12.setCellStyle(dateCellStyle);

                Cell F_S12 = dataRow.createCell(176);
                F_S12.setCellValue(obj.getF_S12());
                F_S12.setCellStyle(numCellStyle);

                Cell F_DATAS13 = dataRow.createCell(177);
                F_DATAS13.setCellValue(obj.getF_DATAS13());
                F_DATAS13.setCellStyle(dateCellStyle);

                Cell F_S13 = dataRow.createCell(178);
                F_S13.setCellValue(obj.getF_S13());
                F_S13.setCellStyle(numCellStyle);

                Cell F_DATAS14 = dataRow.createCell(179);
                F_DATAS14.setCellValue(obj.getF_DATAS14());
                F_DATAS14.setCellStyle(dateCellStyle);

                Cell F_S14 = dataRow.createCell(180);
                F_S14.setCellValue(obj.getF_S14());
                F_S14.setCellStyle(numCellStyle);

                Cell F_DATAS15 = dataRow.createCell(181);
                F_DATAS15.setCellValue(obj.getF_DATAS15());
                F_DATAS15.setCellStyle(dateCellStyle);

                Cell F_S15 = dataRow.createCell(182);
                F_S15.setCellValue(obj.getF_S15());
                F_S15.setCellStyle(numCellStyle);

                Cell F_DATAS16 = dataRow.createCell(183);
                F_DATAS16.setCellValue(obj.getF_DATAS16());
                F_DATAS16.setCellStyle(dateCellStyle);

                Cell F_S16 = dataRow.createCell(184);
                F_S16.setCellValue(obj.getF_S16());
                F_S16.setCellStyle(numCellStyle);

                Cell F_DATAS17 = dataRow.createCell(185);
                F_DATAS17.setCellValue(obj.getF_DATAS17());
                F_DATAS17.setCellStyle(dateCellStyle);

                Cell F_S17 = dataRow.createCell(186);
                F_S17.setCellValue(obj.getF_S17());
                F_S17.setCellStyle(numCellStyle);

                Cell F_DATAS18 = dataRow.createCell(187);
                F_DATAS18.setCellValue(obj.getF_DATAS18());
                F_DATAS18.setCellStyle(dateCellStyle);

                Cell F_S18 = dataRow.createCell(188);
                F_S18.setCellValue(obj.getF_S18());
                F_S18.setCellStyle(numCellStyle);

                Cell F_DATAS19 = dataRow.createCell(189);
                F_DATAS19.setCellValue(obj.getF_DATAS19());
                F_DATAS19.setCellStyle(dateCellStyle);

                Cell F_S19 = dataRow.createCell(190);
                F_S19.setCellValue(obj.getF_S19());
                F_S19.setCellStyle(numCellStyle);

                Cell F_DATAS20 = dataRow.createCell(191);
                F_DATAS20.setCellValue(obj.getF_DATAS20());
                F_DATAS20.setCellStyle(dateCellStyle);

                Cell F_S20 = dataRow.createCell(192);
                F_S20.setCellValue(obj.getF_S20());
                F_S20.setCellStyle(numCellStyle);

                Cell F_TOTAL = dataRow.createCell(193);
                F_TOTAL.setCellValue(obj.getF_TOTAL());
                F_TOTAL.setCellStyle(numCellStyle);
                F_TOTAL.setCellFormula("EU" + rowNum + "-(SUM(EY" + rowNum + "))"
                        + "-(SUM(FA" + rowNum + "))-(SUM(FC" + rowNum + "))-(SUM(FE" + rowNum + "))"
                        + "-(SUM(FG" + rowNum + "))-(SUM(FI" + rowNum + "))-(SUM(FK" + rowNum + "))"
                        + "-(SUM(FM" + rowNum + "))-(SUM(FO" + rowNum + "))-(SUM(FQ" + rowNum + "))"
                        + "-(SUM(FS" + rowNum + "))-(SUM(FU" + rowNum + "))-(SUM(FW" + rowNum + "))"
                        + "-(SUM(FY" + rowNum + "))-(SUM(GA" + rowNum + "))-(SUM(GC" + rowNum + "))"
                        + "-(SUM(GE" + rowNum + "))-(SUM(GG" + rowNum + "))-(SUM(GI" + rowNum + "))"
                        + "-(SUM(GK" + rowNum + "))");

                Cell porcen = dataRow.createCell(194);
                porcen.setCellValue(obj.getF_PERC());
                porcen.setCellStyle(porCellStyle);
                porcen.setCellFormula("GL" + rowNum + "/EU" + rowNum + "");

                Cell obs = dataRow.createCell(195);
                obs.setCellStyle(alignDataStyle);
                obs.setCellFormula("IF(GM" + rowNum + "<=10%,\"ok\",\"atenção\")");

                Cell kgrefugo = dataRow.createCell(196);
                kgrefugo.setCellValue("");
                kgrefugo.setCellStyle(alignDataStyle);
                kgrefugo.setCellFormula("(SUM(AR" + rowNum + "))"
                        + "+(SUM(AW" + rowNum + "))+(SUM(BB" + rowNum + "))"
                        + "+(SUM(BI" + rowNum + "))+(SUM(BN" + rowNum + "))"
                        + "+(SUM(BS" + rowNum + "))+(SUM(BX" + rowNum + "))"
                        + "+(SUM(CF" + rowNum + "))+(SUM(CK" + rowNum + "))"
                        + "+(SUM(CP" + rowNum + "))+(SUM(CU" + rowNum + "))"
                        + "+(SUM(CZ" + rowNum + "))+(SUM(DE" + rowNum + "))"
                        + "+(SUM(DJ" + rowNum + "))+(SUM(DO" + rowNum + "))"
                        + "+(SUM(DT" + rowNum + "))+(SUM(DY" + rowNum + "))"
                        + "+(SUM(ED" + rowNum + "))+(SUM(EJ" + rowNum + "))");

                Cell F_SEQ = dataRow.createCell(197);
                F_SEQ.setCellValue(obj.getF_SEQ());
                F_SEQ.setCellStyle(alignDataStyle);

                Cell FD_SEQPF = dataRow.createCell(198);
                FD_SEQPF.setCellValue(obj.getFD_SEQPF());
                FD_SEQPF.setCellStyle(alignDataStyle);

                Cell FD_ESPESSURA = dataRow.createCell(199);
                FD_ESPESSURA.setCellValue(obj.getFD_ESPESSURA());
                FD_ESPESSURA.setCellStyle(numCellStyle);

                Cell FD_OXIDACAO = dataRow.createCell(200);
                FD_OXIDACAO.setCellValue(obj.getFD_OXIDACAO());
                FD_OXIDACAO.setCellStyle(numCellStyle);

                Cell FD_ONDULACAO = dataRow.createCell(201);
                FD_ONDULACAO.setCellValue(obj.getFD_ONDULACAO());
                FD_ONDULACAO.setCellStyle(numCellStyle);

                Cell FD_CASQUINHA = dataRow.createCell(202);
                FD_CASQUINHA.setCellValue(obj.getFD_CASQUINHA());
                FD_CASQUINHA.setCellStyle(numCellStyle);

                Cell FD_RASTRO_TRATOR = dataRow.createCell(203);
                FD_RASTRO_TRATOR.setCellValue(obj.getFD_RASTRO_TRATOR());
                FD_RASTRO_TRATOR.setCellStyle(numCellStyle);

                Cell FD_MANCHA_OLEO = dataRow.createCell(204);
                FD_MANCHA_OLEO.setCellValue(obj.getFD_MANCHA_OLEO());
                FD_MANCHA_OLEO.setCellStyle(numCellStyle);

                Cell FD_LINHA_DO_LQ = dataRow.createCell(205);
                FD_LINHA_DO_LQ.setCellValue(obj.getFD_LINHA_DO_LQ());
                FD_LINHA_DO_LQ.setCellStyle(numCellStyle);

                Cell FD_RISCO_DA_PRENSA = dataRow.createCell(206);
                FD_RISCO_DA_PRENSA.setCellValue(obj.getFD_RISCO_DA_PRENSA());
                FD_RISCO_DA_PRENSA.setCellStyle(numCellStyle);

                Cell FD_SUJEIRA = dataRow.createCell(207);
                FD_SUJEIRA.setCellValue(obj.getFD_SUJEIRA());
                FD_SUJEIRA.setCellStyle(numCellStyle);

                Cell FD_NATA = dataRow.createCell(208);
                FD_NATA.setCellValue(obj.getFD_NATA());
                FD_NATA.setCellStyle(numCellStyle);

                Cell FD_CAVACO = dataRow.createCell(209);
                FD_CAVACO.setCellValue(obj.getFD_CAVACO());
                FD_CAVACO.setCellStyle(numCellStyle);

                Cell FD_LINHA_PRETA = dataRow.createCell(210);
                FD_LINHA_PRETA.setCellValue(obj.getFD_LINHA_PRETA());
                FD_LINHA_PRETA.setCellStyle(numCellStyle);

                Cell FD_BOLHA_CAB = dataRow.createCell(211);
                FD_BOLHA_CAB.setCellValue(obj.getFD_BOLHA_CAB());
                FD_BOLHA_CAB.setCellStyle(numCellStyle);

                Cell FD_BOLHA_TUB = dataRow.createCell(212);
                FD_BOLHA_TUB.setCellValue(obj.getFD_BOLHA_TUB());
                FD_BOLHA_TUB.setCellStyle(numCellStyle);

                Cell FD_BOLHA_DES = dataRow.createCell(213);
                FD_BOLHA_DES.setCellValue(obj.getFD_BOLHA_DES());
                FD_BOLHA_DES.setCellStyle(numCellStyle);

                Cell FD_RISCO = dataRow.createCell(214);
                FD_RISCO.setCellValue(obj.getFD_RISCO());
                FD_RISCO.setCellStyle(numCellStyle);

                Cell FD_QUEBRA_DE_BOB = dataRow.createCell(215);
                FD_QUEBRA_DE_BOB.setCellValue(obj.getFD_QUEBRA_DE_BOB());
                FD_QUEBRA_DE_BOB.setCellStyle(numCellStyle);

                Cell FD_BURACO = dataRow.createCell(216);
                FD_BURACO.setCellValue(obj.getFD_BURACO());
                FD_BURACO.setCellStyle(numCellStyle);

                Cell FD_MANCHA_BRANCA = dataRow.createCell(217);
                FD_MANCHA_BRANCA.setCellValue(obj.getFD_MANCHA_BRANCA());
                FD_MANCHA_BRANCA.setCellStyle(numCellStyle);

                Cell FD_MANCHA_MARROM = dataRow.createCell(218);
                FD_MANCHA_MARROM.setCellValue(obj.getFD_MANCHA_MARROM());
                FD_MANCHA_MARROM.setCellStyle(numCellStyle);

                Cell FD_BATIMENTO_LATERAL = dataRow.createCell(219);
                FD_BATIMENTO_LATERAL.setCellValue(obj.getFD_BATIMENTO_LATERAL());
                FD_BATIMENTO_LATERAL.setCellStyle(numCellStyle);

                Cell FD_BOLHA_AQUECIMENTO = dataRow.createCell(220);
                FD_BOLHA_AQUECIMENTO.setCellValue(obj.getFD_BOLHA_AQUECIMENTO());
                FD_BOLHA_AQUECIMENTO.setCellStyle(numCellStyle);

                Cell FD_GRAO = dataRow.createCell(221);
                FD_GRAO.setCellValue(obj.getFD_GRAO());
                FD_GRAO.setCellStyle(numCellStyle);

                Cell FD_DUREZADEF = dataRow.createCell(222);
                FD_DUREZADEF.setCellValue(obj.getFD_DUREZADEF());
                FD_DUREZADEF.setCellStyle(numCellStyle);

                Cell FD_SOBRA = dataRow.createCell(223);
                FD_SOBRA.setCellValue(obj.getFD_SOBRA());
                FD_SOBRA.setCellStyle(numCellStyle);

                Cell FD_FLECHA = dataRow.createCell(224);
                FD_FLECHA.setCellValue(obj.getFD_FLECHA());
                FD_FLECHA.setCellStyle(numCellStyle);

                Cell FD_FALTOU_PESO = dataRow.createCell(225);
                FD_FALTOU_PESO.setCellValue(obj.getFD_FALTOU_PESO());
                FD_FALTOU_PESO.setCellStyle(numCellStyle);

                Cell FD_TEMPERATURA = dataRow.createCell(226);
                FD_TEMPERATURA.setCellValue(obj.getFD_TEMPERATURA());
                FD_TEMPERATURA.setCellStyle(numCellStyle);

                Cell FD_OS_PERDIDA = dataRow.createCell(227);
                FD_OS_PERDIDA.setCellValue(obj.getFD_OS_PERDIDA());
                FD_OS_PERDIDA.setCellStyle(numCellStyle);

                Cell FD_PE_DE_PLACA = dataRow.createCell(228);
                FD_PE_DE_PLACA.setCellValue(obj.getFD_PE_DE_PLACA());
                FD_PE_DE_PLACA.setCellStyle(numCellStyle);

                Cell FD_REBARBA = dataRow.createCell(229);
                FD_REBARBA.setCellValue(obj.getFD_REBARBA());
                FD_REBARBA.setCellStyle(numCellStyle);

                Cell FD_MANCHA_D_AGUA = dataRow.createCell(230);
                FD_MANCHA_D_AGUA.setCellValue(obj.getFD_MANCHA_D_AGUA());
                FD_MANCHA_D_AGUA.setCellStyle(numCellStyle);

                Cell FD_OLHO = dataRow.createCell(231);
                FD_OLHO.setCellValue(obj.getFD_OLHO());
                FD_OLHO.setCellStyle(numCellStyle);

                Cell FD_ZINABRO = dataRow.createCell(232);
                FD_ZINABRO.setCellValue(obj.getFD_ZINABRO());
                FD_ZINABRO.setCellStyle(numCellStyle);

                Cell FD_DESALINHADA = dataRow.createCell(233);
                FD_DESALINHADA.setCellValue(obj.getFD_DESALINHADA());
                FD_DESALINHADA.setCellStyle(numCellStyle);

                Cell FD_SUMIU = dataRow.createCell(234);
                FD_SUMIU.setCellValue(obj.getFD_SUMIU());
                FD_SUMIU.setCellStyle(numCellStyle);

                Cell FD_TRINCADA = dataRow.createCell(235);
                FD_TRINCADA.setCellValue(obj.getFD_TRINCADA());
                FD_TRINCADA.setCellStyle(numCellStyle);

                Cell FD_BOB_ESTREITA = dataRow.createCell(236);
                FD_BOB_ESTREITA.setCellValue(obj.getFD_BOB_ESTREITA());
                FD_BOB_ESTREITA.setCellStyle(numCellStyle);

                Cell FD_MARCA_CILINDRO = dataRow.createCell(237);
                FD_MARCA_CILINDRO.setCellValue(obj.getFD_MARCA_CILINDRO());
                FD_MARCA_CILINDRO.setCellStyle(numCellStyle);

                Cell FD_SEM_OS = dataRow.createCell(238);
                FD_SEM_OS.setCellValue(obj.getFD_SEM_OS());
                FD_SEM_OS.setCellStyle(numCellStyle);

                Cell FD_LATERAL_AMASSADA = dataRow.createCell(239);
                FD_LATERAL_AMASSADA.setCellValue(obj.getFD_LATERAL_AMASSADA());
                FD_LATERAL_AMASSADA.setCellStyle(numCellStyle);

                Cell FD_QUEBRA_PROCED = dataRow.createCell(240);
                FD_QUEBRA_PROCED.setCellValue(obj.getFD_QUEBRA_PROCED());
                FD_QUEBRA_PROCED.setCellStyle(numCellStyle);

                Cell FD_RISCO_LQ = dataRow.createCell(241);
                FD_RISCO_LQ.setCellValue(obj.getFD_RISCO_LQ());
                FD_RISCO_LQ.setCellStyle(numCellStyle);

                Cell FD_RISCO_SG = dataRow.createCell(242);
                FD_RISCO_SG.setCellValue(obj.getFD_RISCO_SG());
                FD_RISCO_SG.setCellStyle(numCellStyle);

                Cell FD_RISCO_LF = dataRow.createCell(243);
                FD_RISCO_LF.setCellValue(obj.getFD_RISCO_LF());
                FD_RISCO_LF.setCellStyle(numCellStyle);

                Cell FD_RISCO_SF = dataRow.createCell(244);
                FD_RISCO_SF.setCellValue(obj.getFD_RISCO_SF());
                FD_RISCO_SF.setCellStyle(numCellStyle);

                Cell FD_PEDIDO_CANCEL = dataRow.createCell(245);
                FD_PEDIDO_CANCEL.setCellValue(obj.getFD_PEDIDO_CANCEL());
                FD_PEDIDO_CANCEL.setCellStyle(numCellStyle);

                Cell FD_BOB_DESVIADA = dataRow.createCell(246);
                FD_BOB_DESVIADA.setCellValue(obj.getFD_BOB_DESVIADA());
                FD_BOB_DESVIADA.setCellStyle(numCellStyle);

                Cell FD_MEIA_LUA = dataRow.createCell(247);
                FD_MEIA_LUA.setCellValue(obj.getFD_MEIA_LUA());
                FD_MEIA_LUA.setCellStyle(numCellStyle);

                Cell FD_REFILE_MAIOR = dataRow.createCell(248);
                FD_REFILE_MAIOR.setCellValue(obj.getFD_REFILE_MAIOR());
                FD_REFILE_MAIOR.setCellStyle(numCellStyle);

                Cell FD_ERRO_DE_MEDIA = dataRow.createCell(249);
                FD_ERRO_DE_MEDIA.setCellValue(obj.getFD_ERRO_DE_MEDIA());
                FD_ERRO_DE_MEDIA.setCellStyle(numCellStyle);

                Cell totalRefugo = dataRow.createCell(250);
                totalRefugo.setCellStyle(alignDataStyle);
                totalRefugo.setCellFormula("(SUM(GR" + rowNum + ":IP" + rowNum + "))");

            }   //PRIMEIRA LINHA
            //TEXTOS LINHA 0
            int Row0 = 0;
            Row linha0 = sheet.createRow(Row0);
            Cell celula026 = linha0.createCell(26);
            celula026.setCellValue("Plano de Fabricação");
            sheet.addMergedRegion(CellRangeAddress.valueOf("AA1:AD1"));
            Cell celula030 = linha0.createCell(30);
            celula030.setCellValue("F065-03");
            Cell celula032 = linha0.createCell(32);
            celula032.setCellValue("Dureza HB");
            // incrementamos a linha
            Row0++;
            //TEXTO LINHA 1
            int Row1 = 1;
            Row linha1 = sheet.createRow(Row1);
            Cell celula110 = linha1.createCell(10);
            celula110.setCellValue("NºPF:");
            Cell celula131 = linha1.createCell(31);
            celula131.setCellValue("F0-SUPER");
            sheet.autoSizeColumn(31);
            // incrementamos a linha
            Row1++;
            //TEXTO LINHA 2
            int Row2 = 2;
            Row linha2 = sheet.createRow(Row2);

            Cell celula201 = linha2.createCell(1);
            celula201.setCellValue("ATENDENTE");
            celula201.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("B3:C3"));

            Cell celula203 = linha2.createCell(3);
            celula203.setCellValue("FK");
            celula203.setCellStyle(linha3BrightGreen);

            Cell celula204 = linha2.createCell(4);
            celula204.setCellValue("RAF");
            celula204.setCellStyle(linha3BrightGreen);

            Cell celula205 = linha2.createCell(5);
            celula205.setCellValue("RAF/PCP");
            celula205.setCellStyle(linha3BrightGreen);

            Cell celula206 = linha2.createCell(6);
            celula206.setCellValue("PCP");
            celula206.setCellStyle(linha3BrightGreen);

            Cell celula207 = linha2.createCell(7);
            celula207.setCellValue("RAF");
            celula207.setCellStyle(linha3BrightGreen);

            Cell celula208 = linha2.createCell(8);
            celula208.setCellValue("PCP");
            celula208.setCellStyle(linha3BrightGreen);

            Cell celula209 = linha2.createCell(9);
            celula209.setCellValue("");
            celula209.setCellStyle(linha3Pink);

            Cell celula210 = linha2.createCell(10);
            celula210.setCellValue("F019");
            celula210.setCellStyle(linha3Lime);

            Cell celula211 = linha2.createCell(11);
            celula211.setCellValue("F094");
            celula211.setCellStyle(linha3Pink);

            sheet.addMergedRegion(CellRangeAddress.valueOf("L3:AD3"));
            Cell celula230 = linha2.createCell(30);
            celula230.setCellValue("F019");
            celula230.setCellStyle(linha3Lime);

            Cell celula231 = linha2.createCell(31);
            celula231.setCellValue("F094");
            celula231.setCellStyle(linha3Pink);

            Cell celula232 = linha2.createCell(32);
            celula232.setCellValue("F109");
            celula232.setCellStyle(linha3Pink);

            Cell celula233 = linha2.createCell(33);
            celula233.setCellValue("");
            celula233.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AH3:AI3"));

            Cell celula235 = linha2.createCell(35);
            celula235.setCellValue("F094");
            celula235.setCellStyle(linha3Pink);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AJ3:AK3"));

            Cell celula237 = linha2.createCell(37);
            celula237.setCellValue("F109");
            celula237.setCellStyle(linha3Lime);

            Cell celula238 = linha2.createCell(38);
            celula238.setCellValue("F109");
            celula238.setCellStyle(linha3Lime);

            Cell celula239 = linha2.createCell(39);
            celula239.setCellValue("");
            celula239.setCellStyle(linha3Pink);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AN3:AP3"));

            Cell celula242 = linha2.createCell(42);
            celula242.setCellValue("");
            celula242.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AQ3:AS3"));

            Cell celula245 = linha2.createCell(45);
            celula245.setCellValue("");
            celula245.setCellStyle(linha3Pink);

            Cell celula246 = linha2.createCell(46);
            celula246.setCellValue("F021");
            celula246.setCellStyle(linha3BrightGreen);

            Cell celula247 = linha2.createCell(47);
            celula247.setCellValue("");
            celula247.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AV3:AX3"));

            Cell celula250 = linha2.createCell(50);
            celula250.setCellValue("");
            celula250.setCellStyle(linha3Pink);

            Cell celula251 = linha2.createCell(51);
            celula251.setCellValue("F020");
            celula251.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("AZ3:BC3"));

            Cell celula255 = linha2.createCell(55);
            celula255.setCellValue("");
            celula255.setCellStyle(linha3Pink);

            Cell celula256 = linha2.createCell(56);
            celula256.setCellValue("F020");
            celula256.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("BE3:BJ3"));

            Cell celula262 = linha2.createCell(62);
            celula262.setCellValue("");
            celula262.setCellStyle(linha3Pink);

            Cell celula263 = linha2.createCell(63);
            celula263.setCellValue("F020");
            celula263.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("BL3:BO3"));

            Cell celula267 = linha2.createCell(67);
            celula267.setCellValue("");
            celula267.setCellStyle(linha3Pink);

            Cell celula268 = linha2.createCell(68);
            celula268.setCellValue("F020");
            celula268.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("BQ3:BT3"));

            Cell celula272 = linha2.createCell(72);
            celula272.setCellValue("");
            celula272.setCellStyle(linha3Pink);

            Cell celula273 = linha2.createCell(73);
            celula273.setCellValue("F118");
            celula273.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("BV3:BY3"));

            Cell celula277 = linha2.createCell(77);
            celula277.setCellValue("");
            celula277.setCellStyle(linha3Pink);

            Cell celula278 = linha2.createCell(78);
            celula278.setCellValue("F020");
            celula278.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("CA3:CG3"));

            Cell celula285 = linha2.createCell(85);
            celula285.setCellValue("");
            celula285.setCellStyle(linha3Pink);

            Cell celula286 = linha2.createCell(86);
            celula286.setCellValue("F020");
            celula286.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("CI3:CL3"));

            Cell celula290 = linha2.createCell(90);
            celula290.setCellValue("");
            celula290.setCellStyle(linha3Pink);

            Cell celula291 = linha2.createCell(91);
            celula291.setCellValue("F020");
            celula291.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("CN3:CQ3"));

            Cell celula295 = linha2.createCell(95);
            celula295.setCellValue("");
            celula295.setCellStyle(linha3Pink);

            Cell celula296 = linha2.createCell(96);
            celula296.setCellValue("F020");
            celula296.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("CS3:CV3"));

            Cell celula2100 = linha2.createCell(100);
            celula2100.setCellValue("");
            celula2100.setCellStyle(linha3Pink);

            Cell celula2101 = linha2.createCell(101);
            celula2101.setCellValue("F020");
            celula2101.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("CX3:DA3"));

            Cell celula2105 = linha2.createCell(105);
            celula2105.setCellValue("");
            celula2105.setCellStyle(linha3Pink);

            Cell celula2106 = linha2.createCell(106);
            celula2106.setCellValue("");
            celula2106.setCellStyle(linha3Lime);

            Cell celula2107 = linha2.createCell(107);
            celula2107.setCellValue("F118");
            celula2107.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("DD3:DF3"));

            Cell celula2110 = linha2.createCell(110);
            celula2110.setCellValue("");
            celula2110.setCellStyle(linha3Pink);

            Cell celula2111 = linha2.createCell(111);
            celula2111.setCellValue("");
            celula2111.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("DH3:DK3"));

            Cell celula2115 = linha2.createCell(115);
            celula2115.setCellValue("");
            celula2115.setCellStyle(linha3Pink);

            Cell celula2116 = linha2.createCell(116);
            celula2116.setCellValue("");
            celula2116.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("DM3:DP3"));

            Cell celula2120 = linha2.createCell(120);
            celula2120.setCellValue("");
            celula2120.setCellStyle(linha3Pink);

            Cell celula2121 = linha2.createCell(121);
            celula2121.setCellValue("F118");
            celula2121.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("DR3:DU3"));

            Cell celula2125 = linha2.createCell(125);
            celula2125.setCellValue("");
            celula2125.setCellStyle(linha3Pink);

            Cell celula2126 = linha2.createCell(126);
            celula2126.setCellValue("F020");
            celula2126.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("DW3:DZ3"));

            Cell celula2130 = linha2.createCell(130);
            celula2130.setCellValue("");
            celula2130.setCellStyle(linha3Pink);

            Cell celula2131 = linha2.createCell(131);
            celula2131.setCellValue("");
            celula2131.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("EB3:EE3"));

            Cell celula2135 = linha2.createCell(135);
            celula2135.setCellValue("");
            celula2135.setCellStyle(linha3Pink);

            Cell celula2136 = linha2.createCell(136);
            celula2136.setCellValue("");
            celula2136.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("EG3:EK3"));

            Cell celula2141 = linha2.createCell(141);
            celula2141.setCellValue("");
            celula2141.setCellStyle(linha3Pink);

            Cell celula2142 = linha2.createCell(142);
            celula2142.setCellValue("");
            celula2142.setCellStyle(linha3Yellow);
            sheet.addMergedRegion(CellRangeAddress.valueOf("EM3:EP3"));

            Cell celula2146 = linha2.createCell(146);
            celula2146.setCellValue("CAMINHÃO");
            celula2146.setCellStyle(linha3Yellow);
            sheet.addMergedRegion(CellRangeAddress.valueOf("EQ3:ES3"));

            Cell celula2149 = linha2.createCell(149);
            celula2149.setCellValue("F094");
            celula2149.setCellStyle(linha3Pink);

            Cell celula2150 = linha2.createCell(150);
            celula2150.setCellValue("F069");
            celula2150.setCellStyle(linha3Pink);

            Cell celula2151 = linha2.createCell(151);
            celula2151.setCellValue("");
            celula2151.setCellStyle(linha3Pink);

            Cell celula2152 = linha2.createCell(152);
            celula2152.setCellValue("");
            celula2152.setCellStyle(linha3Rose);

            Cell celula2153 = linha2.createCell(153);
            celula2153.setCellValue("F118");
            celula2153.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("EX3:GL3"));

            Cell celula2194 = linha2.createCell(194);
            celula2194.setCellValue("");
            celula2194.setCellStyle(linha3Pink);
            sheet.addMergedRegion(CellRangeAddress.valueOf("GM3:GO3"));

            Cell celula2199 = linha2.createCell(199);
            celula2199.setCellValue("F064");
            celula2199.setCellStyle(linha3BrightGreen);
            sheet.addMergedRegion(CellRangeAddress.valueOf("GR3:IP3"));

            Cell celula2249 = linha2.createCell(250);
            celula2249.setCellValue("");
            celula2249.setCellStyle(linha3Pink);
            // incrementamos a linha
            Row2++;
            //TEXTO LINHA 3
            int Row3 = 3;
            Row linha3 = sheet.createRow(Row3);

            Cell celula301 = linha3.createCell(1);
            celula301.setCellValue("NºPedido");
            celula301.setCellStyle(linha4White);
            Cell celula302 = linha3.createCell(2);
            celula302.setCellValue("Solicitação Cliente");
            celula302.setCellStyle(linha4White90);
            Cell celula303 = linha3.createCell(3);
            celula303.setCellValue("Aprovação do FK");
            celula303.setCellStyle(linha4White90);
            Cell celula304 = linha3.createCell(4);
            celula304.setCellValue("Análise Crítica ISO 9001");
            celula304.setCellStyle(linha4White90);
            Cell celula305 = linha3.createCell(5);
            celula305.setCellValue("Emitir pré-pedido");
            celula305.setCellStyle(linha4White90);
            Cell celula306 = linha3.createCell(6);
            celula306.setCellValue("Emitir O.S");
            celula306.setCellStyle(linha4White90);
            Cell celula307 = linha3.createCell(7);
            celula307.setCellValue("Prazo Fundir agendado");
            celula307.setCellStyle(linha4White90);
            Cell celula308 = linha3.createCell(8);
            celula308.setCellValue("PCP se o N°de placas for \n maior que 90 placas \n só de F1");
            celula308.setCellStyle(linha4White90);
            Cell celula309 = linha3.createCell(9);
            celula309.setCellValue("Envio P/ Fundição");
            celula309.setCellStyle(linha4White90);
            Cell celula310 = linha3.createCell(10);
            celula310.setCellValue("Lote");
            celula310.setCellStyle(linha3Yellow);
            Cell celula311 = linha3.createCell(11);
            celula311.setCellValue("Medida");
            celula311.setCellStyle(linha4White);
            sheet.addMergedRegion(CellRangeAddress.valueOf("L4:P4"));
            Cell celula316 = linha3.createCell(16);
            celula316.setCellValue("Dureza");
            celula316.setCellStyle(linha4White);
            sheet.addMergedRegion(CellRangeAddress.valueOf("Q4:R4"));
            Cell celula318 = linha3.createCell(18);
            celula318.setCellValue("OS");
            celula318.setCellStyle(linha4White);
            Cell celula319 = linha3.createCell(19);
            celula319.setCellValue("Cliente");
            celula319.setCellStyle(linha4White);
            Cell celula320 = linha3.createCell(20);
            celula320.setCellValue("F0");
            celula320.setCellStyle(linha4White);
            Cell celula321 = linha3.createCell(21);
            celula321.setCellValue("F1");
            celula321.setCellStyle(linha4White);
            Cell celula322 = linha3.createCell(22);
            celula322.setCellValue("F2");
            celula322.setCellStyle(linha4White);
            Cell celula323 = linha3.createCell(23);
            celula323.setCellValue("F3");
            celula323.setCellStyle(linha4White);
            Cell celula324 = linha3.createCell(24);
            celula324.setCellValue("S");
            celula324.setCellStyle(linha4White);
            Cell celula325 = linha3.createCell(25);
            celula325.setCellValue("F1 27");
            celula325.setCellStyle(linha4White);
            Cell celula326 = linha3.createCell(26);
            celula326.setCellValue("Outras Ligas");
            celula326.setCellStyle(linha4White90);
            Cell celula327 = linha3.createCell(27);
            celula327.setCellValue("N° de Placas");
            celula327.setCellStyle(linha4White90);
            Cell celula328 = linha3.createCell(28);
            celula328.setCellValue("Largura");
            celula328.setCellStyle(linha4White90);
            Cell celula329 = linha3.createCell(29);
            celula329.setCellValue("Kg");
            celula329.setCellStyle(linha4White);
            Cell celula330 = linha3.createCell(30);
            celula330.setCellValue("Largura Real");
            celula330.setCellStyle(linha4Lime90);
            Cell celula331 = linha3.createCell(31);
            celula331.setCellValue("N° Bobina");
            celula331.setCellStyle(linha4White);
            Cell celula332 = linha3.createCell(32);
            celula332.setCellValue("Solicitada");
            celula332.setCellStyle(linha4White90);
            Cell celula333 = linha3.createCell(33);
            celula333.setCellValue("Grão");
            celula333.setCellStyle(linha4White);
            Cell celula334 = linha3.createCell(34);
            celula334.setCellValue("Dureza");
            celula334.setCellStyle(linha4White);
            Cell celula335 = linha3.createCell(35);
            celula335.setCellValue("Setor");
            celula335.setCellStyle(linha4White);
            Cell celula336 = linha3.createCell(36);
            celula336.setCellValue("Atendente");
            celula336.setCellStyle(linha4White90);
            Cell celula337 = linha3.createCell(37);
            celula337.setCellValue("Forneiro");
            celula337.setCellStyle(linha4Lime90);

            Cell celula338 = linha3.createCell(38);
            celula338.setCellValue("F");
            celula338.setCellStyle(linha4Lime90);

            Cell celula339 = linha3.createCell(39);
            celula339.setCellValue("Data Levantamento");
            celula339.setCellStyle(linha4PaleBlue90);

            Cell celula340 = linha3.createCell(40);
            celula340.setCellValue("N° Discrepância");
            celula340.setCellStyle(linha4PaleBlue90);

            Cell celula341 = linha3.createCell(41);
            celula341.setCellValue("Responsável");
            celula341.setCellStyle(linha4PaleBlue90);

            Cell celula342 = linha3.createCell(42);
            celula342.setCellValue("Defeito");
            celula342.setCellStyle(linha4Red90);

            Cell celula343 = linha3.createCell(43);
            celula343.setCellValue("Quantidade");
            celula343.setCellStyle(linha4Red90);

            Cell celula344 = linha3.createCell(44);
            celula344.setCellValue("Operador");
            celula344.setCellStyle(linha4Red90);

            Cell celula345 = linha3.createCell(45);
            celula345.setCellValue("N°Escolhinha F");
            celula345.setCellStyle(linha4PaleBlue90);

            Cell celula346 = linha3.createCell(46);
            celula346.setCellValue("LQ");
            celula346.setCellStyle(linha4White90);

            Cell celula347 = linha3.createCell(47);
            celula347.setCellValue("Defeito");
            celula347.setCellStyle(linha4Red90);

            Cell celula348 = linha3.createCell(48);
            celula348.setCellValue("Quantidade");
            celula348.setCellStyle(linha4Red90);

            Cell celula349 = linha3.createCell(49);
            celula349.setCellValue("Operador");
            celula349.setCellStyle(linha4Red90);

            Cell celula350 = linha3.createCell(50);
            celula350.setCellValue("N°Escolhinha LQ");
            celula350.setCellStyle(linha4PaleBlue90);

            Cell celula351 = linha3.createCell(51);
            celula351.setCellValue("S");
            celula351.setCellStyle(linha4White90);

            Cell celula352 = linha3.createCell(52);
            celula352.setCellValue("Defeito");
            celula352.setCellStyle(linha4Red90);

            Cell celula353 = linha3.createCell(53);
            celula353.setCellValue("Quantidade");
            celula353.setCellStyle(linha4Red90);

            Cell celula354 = linha3.createCell(54);
            celula354.setCellValue("Operador");
            celula354.setCellStyle(linha4Red90);

            Cell celula355 = linha3.createCell(55);
            celula355.setCellValue("N°Escolhinha Slit");
            celula355.setCellStyle(linha4PaleBlue90);

            Cell celula356 = linha3.createCell(56);
            celula356.setCellValue("EG");
            celula356.setCellStyle(linha4Tan);

            Cell celula357 = linha3.createCell(57);
            celula357.setCellValue("Esp.(mm)");
            celula357.setCellStyle(linha4White90);

            Cell celula358 = linha3.createCell(58);
            celula358.setCellValue("LF");
            celula358.setCellStyle(linha4White90);

            Cell celula359 = linha3.createCell(59);
            celula359.setCellValue("Defeito");
            celula359.setCellStyle(linha4Red90);

            Cell celula360 = linha3.createCell(60);
            celula360.setCellValue("Quantidade");
            celula360.setCellStyle(linha4Red90);

            Cell celula361 = linha3.createCell(61);
            celula361.setCellValue("Operador");
            celula361.setCellStyle(linha4Red90);

            Cell celula362 = linha3.createCell(62);
            celula362.setCellValue("N°Escolhinha LF");
            celula362.setCellStyle(linha4PaleBlue90);

            Cell celula363 = linha3.createCell(63);
            celula363.setCellValue("LFF2");
            celula363.setCellStyle(linha4White90);

            Cell celula364 = linha3.createCell(64);
            celula364.setCellValue("Defeito");
            celula364.setCellStyle(linha4Red90);

            Cell celula365 = linha3.createCell(65);
            celula365.setCellValue("Quantidade");
            celula365.setCellStyle(linha4Red90);

            Cell celula366 = linha3.createCell(66);
            celula366.setCellValue("Operador");
            celula366.setCellStyle(linha4Red90);

            Cell celula367 = linha3.createCell(67);
            celula367.setCellValue("N°Escolhinha LFF2");
            celula367.setCellStyle(linha4PaleBlue90);

            Cell celula368 = linha3.createCell(68);
            celula368.setCellValue("LFF1");
            celula368.setCellStyle(linha4White90);

            Cell celula369 = linha3.createCell(69);
            celula369.setCellValue("Defeito");
            celula369.setCellStyle(linha4Red90);

            Cell celula370 = linha3.createCell(70);
            celula370.setCellValue("Quantidade");
            celula370.setCellStyle(linha4Red90);

            Cell celula371 = linha3.createCell(71);
            celula371.setCellValue("Operador");
            celula371.setCellStyle(linha4Red90);

            Cell celula372 = linha3.createCell(72);
            celula372.setCellValue("N°Escolhinha LFF1");
            celula372.setCellStyle(linha4PaleBlue90);

            Cell celula373 = linha3.createCell(73);
            celula373.setCellValue("E");
            celula373.setCellStyle(linha4White90);

            Cell celula374 = linha3.createCell(74);
            celula374.setCellValue("Defeito");
            celula374.setCellStyle(linha4Red90);

            Cell celula375 = linha3.createCell(75);
            celula375.setCellValue("Quantidade");
            celula375.setCellStyle(linha4Red90);

            Cell celula376 = linha3.createCell(76);
            celula376.setCellValue("Operador");
            celula376.setCellStyle(linha4Red90);

            Cell celula377 = linha3.createCell(77);
            celula377.setCellValue("N°Escolhinha E");
            celula377.setCellStyle(linha4PaleBlue90);

            Cell celula378 = linha3.createCell(78);
            celula378.setCellValue("RL BUHLER");
            celula378.setCellStyle(linha4White90);

            Cell celula379 = linha3.createCell(79);
            celula379.setCellValue("RL LFF2");
            celula379.setCellStyle(linha4Tan);

            Cell celula380 = linha3.createCell(80);
            celula380.setCellValue("RL LFF1");
            celula380.setCellStyle(linha4Tan);

            Cell celula381 = linha3.createCell(81);
            celula381.setCellValue("RL LA9");
            celula381.setCellStyle(linha4White90);

            Cell celula382 = linha3.createCell(82);
            celula382.setCellValue("Defeito");
            celula382.setCellStyle(linha4Red90);

            Cell celula383 = linha3.createCell(83);
            celula383.setCellValue("Quantidade");
            celula383.setCellStyle(linha4Red90);

            Cell celula384 = linha3.createCell(84);
            celula384.setCellValue("Operador");
            celula384.setCellStyle(linha4Red90);

            Cell celula385 = linha3.createCell(85);
            celula385.setCellValue("N°Escolhinha RL");
            celula385.setCellStyle(linha4PaleBlue90);

            Cell celula386 = linha3.createCell(86);
            celula386.setCellValue("LA9");
            celula386.setCellStyle(linha4White90);

            Cell celula387 = linha3.createCell(87);
            celula387.setCellValue("Defeito");
            celula387.setCellStyle(linha4Red90);

            Cell celula388 = linha3.createCell(88);
            celula388.setCellValue("Quantidade");
            celula388.setCellStyle(linha4Red90);

            Cell celula389 = linha3.createCell(89);
            celula389.setCellValue("Operador");
            celula389.setCellStyle(linha4Red90);

            Cell celula390 = linha3.createCell(90);
            celula390.setCellValue("N°Escolinha LA9");
            celula390.setCellStyle(linha4PaleBlue90);

            Cell celula391 = linha3.createCell(91);
            celula391.setCellValue("ROT");
            celula391.setCellStyle(linha4White90);

            Cell celula392 = linha3.createCell(92);
            celula392.setCellValue("Defeito");
            celula392.setCellStyle(linha4Red90);

            Cell celula393 = linha3.createCell(93);
            celula393.setCellValue("Quantidade");
            celula393.setCellStyle(linha4Red90);

            Cell celula394 = linha3.createCell(94);
            celula394.setCellValue("Operador");
            celula394.setCellStyle(linha4Red90);

            Cell celula395 = linha3.createCell(95);
            celula395.setCellValue("N°Escolhinha ROT");
            celula395.setCellStyle(linha4PaleBlue90);

            Cell celula396 = linha3.createCell(96);
            celula396.setCellValue("SF(F)");
            celula396.setCellStyle(linha4White90);

            Cell celula397 = linha3.createCell(97);
            celula397.setCellValue("Defeito");
            celula397.setCellStyle(linha4Red90);

            Cell celula398 = linha3.createCell(98);
            celula398.setCellValue("Quantidade");
            celula398.setCellStyle(linha4Red90);

            Cell celula399 = linha3.createCell(99);
            celula399.setCellValue("Operador");
            celula399.setCellStyle(linha4Red90);

            Cell celula3100 = linha3.createCell(100);
            celula3100.setCellValue("N°Escolhinha SF(F)");
            celula3100.setCellStyle(linha4PaleBlue90);

            Cell celula3101 = linha3.createCell(101);
            celula3101.setCellValue("SF(R)");
            celula3101.setCellStyle(linha4White90);

            Cell celula3102 = linha3.createCell(102);
            celula3102.setCellValue("Defeito");
            celula3102.setCellStyle(linha4Red90);

            Cell celula3103 = linha3.createCell(103);
            celula3103.setCellValue("Quantidade");
            celula3103.setCellStyle(linha4Red90);

            Cell celula3104 = linha3.createCell(104);
            celula3104.setCellValue("Operador");
            celula3104.setCellStyle(linha4Red90);

            Cell celula3105 = linha3.createCell(105);
            celula3105.setCellValue("N°Escolhinha SF(R)");
            celula3105.setCellStyle(linha4PaleBlue90);

            Cell celula3106 = linha3.createCell(106);
            celula3106.setCellValue("P");
            celula3106.setCellStyle(linha4Lime90);

            Cell celula3107 = linha3.createCell(107);
            celula3107.setCellValue("Defeito");
            celula3107.setCellStyle(linha4Red90);

            Cell celula3108 = linha3.createCell(108);
            celula3108.setCellValue("Quantidade");
            celula3108.setCellStyle(linha4Red90);

            Cell celula3109 = linha3.createCell(109);
            celula3109.setCellValue("Operador");
            celula3109.setCellStyle(linha4Red90);

            Cell celula3110 = linha3.createCell(110);
            celula3110.setCellValue("N°Escolhinha P");
            celula3110.setCellStyle(linha4PaleBlue90);

            Cell celula3111 = linha3.createCell(111);
            celula3111.setCellValue("GUI(I)");
            celula3111.setCellStyle(linha4White90);

            Cell celula3112 = linha3.createCell(112);
            celula3112.setCellValue("Defeito");
            celula3112.setCellStyle(linha4Red90);

            Cell celula3113 = linha3.createCell(113);
            celula3113.setCellValue("Quantidade");
            celula3113.setCellStyle(linha4Red90);

            Cell celula3114 = linha3.createCell(114);
            celula3114.setCellValue("Operador");
            celula3114.setCellStyle(linha4Red90);

            Cell celula3115 = linha3.createCell(115);
            celula3115.setCellValue("N°Escolhinha GUI(I)");
            celula3115.setCellStyle(linha4PaleBlue90);

            Cell celula3116 = linha3.createCell(116);
            celula3116.setCellValue("GUI(F)");
            celula3116.setCellStyle(linha4White90);

            Cell celula3117 = linha3.createCell(117);
            celula3117.setCellValue("Defeito");
            celula3117.setCellStyle(linha4Red90);

            Cell celula3118 = linha3.createCell(118);
            celula3118.setCellValue("Quantidade");
            celula3118.setCellStyle(linha4Red90);

            Cell celula3119 = linha3.createCell(119);
            celula3119.setCellValue("Operador");
            celula3119.setCellStyle(linha4Red90);

            Cell celula3120 = linha3.createCell(120);
            celula3120.setCellValue("N°Escolhinha GUI(F)");
            celula3120.setCellStyle(linha4PaleBlue90);

            Cell celula3121 = linha3.createCell(121);
            celula3121.setCellValue("R");
            celula3121.setCellStyle(linha4White90);

            Cell celula3122 = linha3.createCell(122);
            celula3122.setCellValue("Defeito");
            celula3122.setCellStyle(linha4Red90);

            Cell celula3123 = linha3.createCell(123);
            celula3123.setCellValue("Quantidade");
            celula3123.setCellStyle(linha4Red90);

            Cell celula31124 = linha3.createCell(124);
            celula31124.setCellValue("Operador");
            celula31124.setCellStyle(linha4Red90);

            Cell celula3125 = linha3.createCell(125);
            celula3125.setCellValue("N°Escolhinha R");
            celula3125.setCellStyle(linha4PaleBlue90);

            Cell celula3126 = linha3.createCell(126);
            celula3126.setCellValue("PE");
            celula3126.setCellStyle(linha4White90);

            Cell celula3127 = linha3.createCell(127);
            celula3127.setCellValue("Defeito");
            celula3127.setCellStyle(linha4Red90);

            Cell celula3128 = linha3.createCell(128);
            celula3128.setCellValue("Quantidade");
            celula3128.setCellStyle(linha4Red90);

            Cell celula3129 = linha3.createCell(129);
            celula3129.setCellValue("Operador");
            celula3129.setCellStyle(linha4Red90);

            Cell celula3130 = linha3.createCell(130);
            celula3130.setCellValue("N°Escolhinha PE");
            celula3130.setCellStyle(linha4PaleBlue90);

            Cell celula3131 = linha3.createCell(131);
            celula3131.setCellValue("C");
            celula3131.setCellStyle(linha4White90);

            Cell celula3132 = linha3.createCell(132);
            celula3132.setCellValue("Defeito");
            celula3132.setCellStyle(linha4Red90);

            Cell celula3133 = linha3.createCell(133);
            celula3133.setCellValue("Quantidade");
            celula3133.setCellStyle(linha4Red90);

            Cell celula3134 = linha3.createCell(134);
            celula3134.setCellValue("Operador");
            celula3134.setCellStyle(linha4Red90);

            Cell celula3135 = linha3.createCell(135);
            celula3135.setCellValue("N°Escolhinha C");
            celula3135.setCellStyle(linha4PaleBlue90);

            Cell celula3136 = linha3.createCell(136);
            celula3136.setCellValue("ESCOLHA");
            celula3136.setCellStyle(linha4White90);

            Cell celula3137 = linha3.createCell(137);
            celula3137.setCellValue("Realizado");
            celula3137.setCellStyle(linha4White90);

            Cell celula3138 = linha3.createCell(138);
            celula3138.setCellValue("Defeito");
            celula3138.setCellStyle(linha4Red90);

            Cell celula139 = linha3.createCell(139);
            celula139.setCellValue("Quantidade");
            celula139.setCellStyle(linha4Red90);

            Cell celula3140 = linha3.createCell(140);
            celula3140.setCellValue("Operador");
            celula3140.setCellStyle(linha4Red90);

            Cell celula3141 = linha3.createCell(141);
            celula3141.setCellValue("N°Escolhinha ESC");
            celula3141.setCellStyle(linha4PaleBlue90);

            Cell celula3142 = linha3.createCell(142);
            celula3142.setCellValue("Embarque");
            celula3142.setCellStyle(linha4Yellow90);

            Cell celula3143 = linha3.createCell(143);
            celula3143.setCellValue("Manhã(M)/Tarde(T)");
            celula3143.setCellStyle(linha4White90);

            Cell celula3144 = linha3.createCell(144);
            celula3144.setCellValue("Realizado");
            celula3144.setCellStyle(linha4White90);

            Cell celula3145 = linha3.createCell(145);
            celula3145.setCellValue("Região");
            celula3145.setCellStyle(linha4White90);

            Cell celula3146 = linha3.createCell(146);
            celula3146.setCellValue("MÉDIO(BEBÊ)");
            celula3146.setCellStyle(linha4Yellow90);

            Cell celula3147 = linha3.createCell(147);
            celula3147.setCellValue("GRD(VOVÔ)");
            celula3147.setCellStyle(linha4Yellow90);

            Cell celula3148 = linha3.createCell(148);
            celula3148.setCellValue("BEBÊ/VOVÔ");
            celula3148.setCellStyle(linha4Yellow90);

            Cell celula3149 = linha3.createCell(149);
            celula3149.setCellValue("Entrada Solicitada");
            celula3149.setCellStyle(linha4White90);

            Cell celula3150 = linha3.createCell(150);
            celula3150.setCellValue("Kg Liquido(teor)");
            celula3150.setCellStyle(linha4White90);

            Cell celula3151 = linha3.createCell(151);
            celula3151.setCellValue("Kg Pedido");
            celula3151.setCellStyle(linha4White90);

            Cell celula3152 = linha3.createCell(152);
            celula3152.setCellValue("Estoque");
            celula3152.setCellStyle(linha3Rose);

            Cell celula3153 = linha3.createCell(153);
            celula3153.setCellValue("Saída");
            celula3153.setCellStyle(linha4White90);

            Cell celula3154 = linha3.createCell(154);
            celula3154.setCellValue("Peso");
            celula3154.setCellStyle(linha4White90);

            Cell celula3155 = linha3.createCell(155);
            celula3155.setCellValue("Saída");
            celula3155.setCellStyle(linha4White90);

            Cell celula3156 = linha3.createCell(156);
            celula3156.setCellValue("Peso");
            celula3156.setCellStyle(linha4White90);

            Cell celula3157 = linha3.createCell(157);
            celula3157.setCellValue("Saída");
            celula3157.setCellStyle(linha4White90);

            Cell celula3158 = linha3.createCell(158);
            celula3158.setCellValue("Peso");
            celula3158.setCellStyle(linha4White90);

            Cell celula3159 = linha3.createCell(159);
            celula3159.setCellValue("Saída");
            celula3159.setCellStyle(linha4White90);

            Cell celula3160 = linha3.createCell(160);
            celula3160.setCellValue("Peso");
            celula3160.setCellStyle(linha4White90);

            Cell celula3161 = linha3.createCell(161);
            celula3161.setCellValue("Saída");
            celula3161.setCellStyle(linha4White90);

            Cell celula3162 = linha3.createCell(162);
            celula3162.setCellValue("Peso");
            celula3162.setCellStyle(linha4White90);

            Cell celula3163 = linha3.createCell(163);
            celula3163.setCellValue("Saída");
            celula3163.setCellStyle(linha4White90);

            Cell celula3164 = linha3.createCell(164);
            celula3164.setCellValue("Peso");
            celula3164.setCellStyle(linha4White90);

            Cell celula3165 = linha3.createCell(165);
            celula3165.setCellValue("Saída");
            celula3165.setCellStyle(linha4White90);

            Cell celula3166 = linha3.createCell(166);
            celula3166.setCellValue("Peso");
            celula3166.setCellStyle(linha4White90);

            Cell celula3167 = linha3.createCell(167);
            celula3167.setCellValue("Saída");
            celula3167.setCellStyle(linha4White90);

            Cell celula3168 = linha3.createCell(168);
            celula3168.setCellValue("Peso");
            celula3168.setCellStyle(linha4White90);

            Cell celula3169 = linha3.createCell(169);
            celula3169.setCellValue("Saída");
            celula3169.setCellStyle(linha4White90);

            Cell celula3170 = linha3.createCell(170);
            celula3170.setCellValue("Peso");
            celula3170.setCellStyle(linha4White90);

            Cell celula3171 = linha3.createCell(171);
            celula3171.setCellValue("Saída");
            celula3171.setCellStyle(linha4White90);

            Cell celula3172 = linha3.createCell(172);
            celula3172.setCellValue("Peso");
            celula3172.setCellStyle(linha4White90);

            Cell celula3173 = linha3.createCell(173);
            celula3173.setCellValue("Saída");
            celula3173.setCellStyle(linha4White90);

            Cell celula3174 = linha3.createCell(174);
            celula3174.setCellValue("Peso");
            celula3174.setCellStyle(linha4White90);

            Cell celula3175 = linha3.createCell(175);
            celula3175.setCellValue("Saída");
            celula3175.setCellStyle(linha4White90);

            Cell celula3176 = linha3.createCell(176);
            celula3176.setCellValue("Peso");
            celula3176.setCellStyle(linha4White90);

            Cell celula3177 = linha3.createCell(177);
            celula3177.setCellValue("Saída");
            celula3177.setCellStyle(linha4White90);

            Cell celula3178 = linha3.createCell(178);
            celula3178.setCellValue("Peso");
            celula3178.setCellStyle(linha4White90);

            Cell celula3179 = linha3.createCell(179);
            celula3179.setCellValue("Saída");
            celula3179.setCellStyle(linha4White90);

            Cell celula3180 = linha3.createCell(180);
            celula3180.setCellValue("Peso");
            celula3180.setCellStyle(linha4White90);

            Cell celula3181 = linha3.createCell(181);
            celula3181.setCellValue("Saída");
            celula3181.setCellStyle(linha4White90);

            Cell celula3182 = linha3.createCell(182);
            celula3182.setCellValue("Peso");
            celula3182.setCellStyle(linha4White90);

            Cell celula3183 = linha3.createCell(183);
            celula3183.setCellValue("Saída");
            celula3183.setCellStyle(linha4White90);

            Cell celula3184 = linha3.createCell(184);
            celula3184.setCellValue("Peso");
            celula3184.setCellStyle(linha4White90);

            Cell celula3185 = linha3.createCell(185);
            celula3185.setCellValue("Saída");
            celula3185.setCellStyle(linha4White90);

            Cell celula3186 = linha3.createCell(186);
            celula3186.setCellValue("Peso");
            celula3186.setCellStyle(linha4White90);

            Cell celula3187 = linha3.createCell(187);
            celula3187.setCellValue("Saída");
            celula3187.setCellStyle(linha4White90);

            Cell celula3188 = linha3.createCell(188);
            celula3188.setCellValue("Peso");
            celula3188.setCellStyle(linha4White90);

            Cell celula3189 = linha3.createCell(189);
            celula3189.setCellValue("Saída");
            celula3189.setCellStyle(linha4White90);

            Cell celula3190 = linha3.createCell(190);
            celula3190.setCellValue("Peso");
            celula3190.setCellStyle(linha4White90);

            Cell celula3191 = linha3.createCell(191);
            celula3191.setCellValue("Saída");
            celula3191.setCellStyle(linha4White90);

            Cell celula3192 = linha3.createCell(192);
            celula3192.setCellValue("Peso");
            celula3192.setCellStyle(linha4White90);

            Cell celula3193 = linha3.createCell(193);
            celula3193.setCellValue("Total");
            celula3193.setCellStyle(linha4White90);

            Cell celula3194 = linha3.createCell(194);
            celula3194.setCellValue("%");
            celula3194.setCellStyle(linha4White90);

            Cell celula3195 = linha3.createCell(195);
            celula3195.setCellValue("Observação");
            celula3195.setCellStyle(linha4White90);

            Cell celula3196 = linha3.createCell(196);
            celula3196.setCellValue("Kg Refugo");
            celula3196.setCellStyle(linha4White90);

            Cell celula3197 = linha3.createCell(197);
            celula3197.setCellValue("");
            celula3197.setCellStyle(linha3Rose);

            Cell celula3198 = linha3.createCell(198);
            celula3198.setCellValue("");
            celula3198.setCellStyle(linha3Rose);

            Cell celula3199 = linha3.createCell(199);
            celula3199.setCellValue("Espessura");
            celula3199.setCellStyle(linha4Yellow90);

            Cell celula3200 = linha3.createCell(200);
            celula3200.setCellValue("Oxidação");
            celula3200.setCellStyle(linha4Lime90);

            Cell celula3201 = linha3.createCell(201);
            celula3201.setCellValue("Ondulação");
            celula3201.setCellStyle(linha4Lime90);

            Cell celula3202 = linha3.createCell(202);
            celula3202.setCellValue("Casquinha");
            celula3202.setCellStyle(linha4Lime90);

            Cell celula3203 = linha3.createCell(203);
            celula3203.setCellValue("Rastro Trator");
            celula3203.setCellStyle(linha4Lime90);

            Cell celula3204 = linha3.createCell(204);
            celula3204.setCellValue("Mancha Óleo");
            celula3204.setCellStyle(linha4Lime90);

            Cell celula3205 = linha3.createCell(205);
            celula3205.setCellValue("Linha do LQ");
            celula3205.setCellStyle(linha4Lime90);

            Cell celula3206 = linha3.createCell(206);
            celula3206.setCellValue("Risco da Prensa");
            celula3206.setCellStyle(linha4Lime90);

            Cell celula3207 = linha3.createCell(207);
            celula3207.setCellValue("Sujeira");
            celula3207.setCellStyle(linha4Lime90);

            Cell celula3208 = linha3.createCell(208);
            celula3208.setCellValue("Nata");
            celula3208.setCellStyle(linha4Lime90);

            Cell celula3209 = linha3.createCell(209);
            celula3209.setCellValue("Cavaco");
            celula3209.setCellStyle(linha4Yellow90);

            Cell celula3210 = linha3.createCell(210);
            celula3210.setCellValue("Linha Preta");
            celula3210.setCellStyle(linha4Lime90);

            Cell celula3211 = linha3.createCell(211);
            celula3211.setCellValue("Bolha Cab");
            celula3211.setCellStyle(linha4Lime90);

            Cell celula3212 = linha3.createCell(212);
            celula3212.setCellValue("Bolha Tub");
            celula3212.setCellStyle(linha4Lime90);

            Cell celula3213 = linha3.createCell(213);
            celula3213.setCellValue("Bolha Des");
            celula3213.setCellStyle(linha4Lime90);

            Cell celula3214 = linha3.createCell(214);
            celula3214.setCellValue("Risco");
            celula3214.setCellStyle(linha4Lime90);

            Cell celula3215 = linha3.createCell(215);
            celula3215.setCellValue("Quebra de Bobina");
            celula3215.setCellStyle(linha4Lime90);

            Cell celula3216 = linha3.createCell(216);
            celula3216.setCellValue("Buraco");
            celula3216.setCellStyle(linha4Lime90);

            Cell celula3217 = linha3.createCell(217);
            celula3217.setCellValue("Mancha Branca");
            celula3217.setCellStyle(linha4Lime90);

            Cell celula3218 = linha3.createCell(218);
            celula3218.setCellValue("Mancha Marrom");
            celula3218.setCellStyle(linha4Lime90);

            Cell celula3219 = linha3.createCell(219);
            celula3219.setCellValue("Batimento Lateral");
            celula3219.setCellStyle(linha4Yellow90);

            Cell celula3220 = linha3.createCell(220);
            celula3220.setCellValue("Bolha Aquecimento");
            celula3220.setCellStyle(linha4Lime90);

            Cell celula3221 = linha3.createCell(221);
            celula3221.setCellValue("Grão");
            celula3221.setCellStyle(linha4Lime90);

            Cell celula3222 = linha3.createCell(222);
            celula3222.setCellValue("Dureza");
            celula3222.setCellStyle(linha4Lime90);

            Cell celula3223 = linha3.createCell(223);
            celula3223.setCellValue("Sobra");
            celula3223.setCellStyle(linha4Lime90);

            Cell celula3224 = linha3.createCell(224);
            celula3224.setCellValue("Flecha");
            celula3224.setCellStyle(linha4Lime90);

            Cell celula3225 = linha3.createCell(225);
            celula3225.setCellValue("Faltou Peso");
            celula3225.setCellStyle(linha4Lime90);

            Cell celula3226 = linha3.createCell(226);
            celula3226.setCellValue("Temperatura");
            celula3226.setCellStyle(linha4Lime90);

            Cell celula3227 = linha3.createCell(227);
            celula3227.setCellValue("OS Perdida");
            celula3227.setCellStyle(linha4Lime90);

            Cell celula3228 = linha3.createCell(228);
            celula3228.setCellValue("Pe de Placa");
            celula3228.setCellStyle(linha4Lime90);

            Cell celula3229 = linha3.createCell(229);
            celula3229.setCellValue("Rebarba");
            celula3229.setCellStyle(linha4Yellow90);

            Cell celula3230 = linha3.createCell(230);
            celula3230.setCellValue("Mancha D'Agua");
            celula3230.setCellStyle(linha4Lime90);

            Cell celula3231 = linha3.createCell(231);
            celula3231.setCellValue("Olho");
            celula3231.setCellStyle(linha4Lime90);

            Cell celula3232 = linha3.createCell(232);
            celula3232.setCellValue("Zinabro");
            celula3232.setCellStyle(linha4Lime90);

            Cell celula3233 = linha3.createCell(233);
            celula3233.setCellValue("Desalinhada");
            celula3233.setCellStyle(linha4Lime90);

            Cell celula3234 = linha3.createCell(234);
            celula3234.setCellValue("Sumiu");
            celula3234.setCellStyle(linha4Lime90);

            Cell celula3235 = linha3.createCell(235);
            celula3235.setCellValue("Trincada");
            celula3235.setCellStyle(linha4Lime90);

            Cell celula3236 = linha3.createCell(236);
            celula3236.setCellValue("Bobina Estreita");
            celula3236.setCellStyle(linha4Lime90);

            Cell celula3237 = linha3.createCell(237);
            celula3237.setCellValue("Marca do Cilindro");
            celula3237.setCellStyle(linha4Lime90);

            Cell celula3238 = linha3.createCell(238);
            celula3238.setCellValue("Sem OS");
            celula3238.setCellStyle(linha4Yellow90);

            Cell celula3239 = linha3.createCell(239);
            celula3239.setCellValue("Lateral Amassada");
            celula3239.setCellStyle(linha4Lime90);

            Cell celula3240 = linha3.createCell(240);
            celula3240.setCellValue("Quebra de procedimento");
            celula3240.setCellStyle(linha4Lime90);

            Cell celula3241 = linha3.createCell(241);
            celula3241.setCellValue("Risco LQ");
            celula3241.setCellStyle(linha4Lime90);

            Cell celula3242 = linha3.createCell(242);
            celula3242.setCellValue("Risco SG");
            celula3242.setCellStyle(linha4Lime90);

            Cell celula3243 = linha3.createCell(243);
            celula3243.setCellValue("Risco LF");
            celula3243.setCellStyle(linha4Lime90);

            Cell celula3244 = linha3.createCell(244);
            celula3244.setCellValue("Risco SF");
            celula3244.setCellStyle(linha4Lime90);

            Cell celula3245 = linha3.createCell(245);
            celula3245.setCellValue("Pedido Cancelado");
            celula3245.setCellStyle(linha4Lime90);

            Cell celula3246 = linha3.createCell(246);
            celula3246.setCellValue("Bobina Desviada");
            celula3246.setCellStyle(linha4Lime90);

            Cell celula3247 = linha3.createCell(247);
            celula3247.setCellValue("Meia Lua");
            celula3247.setCellStyle(linha4Lime90);

            Cell celula3248 = linha3.createCell(248);
            celula3248.setCellValue("Refile Maior");
            celula3248.setCellStyle(linha4Lime90);

            Cell celula3249 = linha3.createCell(249);
            celula3249.setCellValue("Erro de Medida");
            celula3249.setCellStyle(linha4Lime90);

            Cell celula3250 = linha3.createCell(250);
            celula3250.setCellValue("Total Refugo");
            celula3250.setCellStyle(linha4White90);

            Cell celula3251 = linha3.createCell(251);
            celula3251.setCellValue("");
            celula3251.setCellStyle(linha3Rose);

            // LINHA PÓS PLANILHA CONTENDO OS VALORES TOTAIS DA COLUNA
            Row3++;
            Font headerFont5 = workbook.createFont();
            headerFont5.setFontHeightInPoints((short) 10);
            headerFont5.setBold(true);

            CellStyle linha5WhiteIn = workbook.createCellStyle();
            linha5WhiteIn.setFont(headerFont5);
            linha5WhiteIn.setAlignment(HorizontalAlignment.CENTER);
            linha5WhiteIn.setBorderBottom(BorderStyle.THIN);
            linha5WhiteIn.setBorderTop(BorderStyle.THIN);
            linha5WhiteIn.setBorderLeft(BorderStyle.DOUBLE);

            CellStyle linha5WhiteOut = workbook.createCellStyle();
            linha5WhiteOut.setFont(headerFont5);
            linha5WhiteOut.setAlignment(HorizontalAlignment.LEFT);
            linha5WhiteOut.setBorderBottom(BorderStyle.THIN);
            linha5WhiteOut.setBorderTop(BorderStyle.THIN);
            linha5WhiteOut.setBorderRight(BorderStyle.DOUBLE);
            linha5WhiteOut.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
            
            CellStyle linha5WhiteAll = workbook.createCellStyle();
            linha5WhiteAll.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00"));
            linha5WhiteAll.setFont(headerFont5);
            linha5WhiteAll.setAlignment(HorizontalAlignment.LEFT);
            linha5WhiteAll.setBorderBottom(BorderStyle.THIN);
            linha5WhiteAll.setBorderTop(BorderStyle.THIN);
            linha5WhiteAll.setBorderRight(BorderStyle.DOUBLE);
            linha5WhiteAll.setBorderLeft(BorderStyle.DOUBLE);

            int Row5 = rowNum + 1;
            Row linha5 = sheet.createRow(Row5);

            Cell celula542 = linha5.createCell(42);
            celula542.setCellValue("Total:");
            celula542.setCellStyle(linha5WhiteIn);

            Cell celula543 = linha5.createCell(43);
            celula543.setCellFormula("SUM(AR5:AR" + rowNum + ")");
            celula543.setCellStyle(linha5WhiteOut);

            Cell celula547 = linha5.createCell(47);
            celula547.setCellValue("Total:");
            celula547.setCellStyle(linha5WhiteIn);

            Cell celula548 = linha5.createCell(48);
            celula548.setCellFormula("SUM(AW5:AW" + rowNum + ")");
            celula548.setCellStyle(linha5WhiteOut);

            Cell celula552 = linha5.createCell(52);
            celula552.setCellValue("Total:");
            celula552.setCellStyle(linha5WhiteIn);

            Cell celula553 = linha5.createCell(53);
            celula553.setCellFormula("SUM(BB5:BB" + rowNum + ")");
            celula553.setCellStyle(linha5WhiteOut);

            Cell celula559 = linha5.createCell(59);
            celula559.setCellValue("Total:");
            celula559.setCellStyle(linha5WhiteIn);

            Cell celula560 = linha5.createCell(60);
            celula560.setCellFormula("SUM(BI5:BI" + rowNum + ")");
            celula560.setCellStyle(linha5WhiteOut);

            Cell celula564 = linha5.createCell(64);
            celula564.setCellValue("Total:");
            celula564.setCellStyle(linha5WhiteIn);

            Cell celula565 = linha5.createCell(65);
            celula565.setCellFormula("SUM(BN5:BN" + rowNum + ")");
            celula565.setCellStyle(linha5WhiteOut);

            Cell celula569 = linha5.createCell(69);
            celula569.setCellValue("Total:");
            celula569.setCellStyle(linha5WhiteIn);

            Cell celula570 = linha5.createCell(70);
            celula570.setCellFormula("SUM(BS5:BS" + rowNum + ")");
            celula570.setCellStyle(linha5WhiteOut);

            Cell celula574 = linha5.createCell(74);
            celula574.setCellValue("Total:");
            celula574.setCellStyle(linha5WhiteIn);

            Cell celula575 = linha5.createCell(75);
            celula575.setCellFormula("SUM(BX5:BX" + rowNum + ")");
            celula575.setCellStyle(linha5WhiteOut);

            Cell celula582 = linha5.createCell(82);
            celula582.setCellValue("Total:");
            celula582.setCellStyle(linha5WhiteIn);

            Cell celula583 = linha5.createCell(83);
            celula583.setCellFormula("SUM(CF5:CF" + rowNum + ")");
            celula583.setCellStyle(linha5WhiteOut);

            Cell celula587 = linha5.createCell(87);
            celula587.setCellValue("Total:");
            celula587.setCellStyle(linha5WhiteIn);

            Cell celula588 = linha5.createCell(88);
            celula588.setCellFormula("SUM(CK5:CK" + rowNum + ")");
            celula588.setCellStyle(linha5WhiteOut);

            Cell celula592 = linha5.createCell(92);
            celula592.setCellValue("Total:");
            celula592.setCellStyle(linha5WhiteIn);

            Cell celula593 = linha5.createCell(93);
            celula593.setCellFormula("SUM(CP5:CP" + rowNum + ")");
            celula593.setCellStyle(linha5WhiteOut);

            Cell celula597 = linha5.createCell(97);
            celula597.setCellValue("Total:");
            celula597.setCellStyle(linha5WhiteIn);

            Cell celula598 = linha5.createCell(98);
            celula598.setCellFormula("SUM(CU5:CU" + rowNum + ")");
            celula598.setCellStyle(linha5WhiteOut);

            Cell celula5102 = linha5.createCell(102);
            celula5102.setCellValue("Total:");
            celula5102.setCellStyle(linha5WhiteIn);

            Cell celula5103 = linha5.createCell(103);
            celula5103.setCellFormula("SUM(CZ5:CZ" + rowNum + ")");
            celula5103.setCellStyle(linha5WhiteOut);

            Cell celula5107 = linha5.createCell(107);
            celula5107.setCellValue("Total:");
            celula5107.setCellStyle(linha5WhiteIn);

            Cell celula5108 = linha5.createCell(108);
            celula5108.setCellFormula("SUM(DE5:DE" + rowNum + ")");
            celula5108.setCellStyle(linha5WhiteOut);

            Cell celula5112 = linha5.createCell(112);
            celula5112.setCellValue("Total:");
            celula5112.setCellStyle(linha5WhiteIn);

            Cell celula5113 = linha5.createCell(113);

            celula5113.setCellFormula("SUM(DJ5:DJ" + rowNum + ")");
            celula5113.setCellStyle(linha5WhiteOut);

            Cell celula5117 = linha5.createCell(117);
            celula5117.setCellValue("Total:");
            celula5117.setCellStyle(linha5WhiteIn);

            Cell celula5118 = linha5.createCell(118);
            celula5118.setCellFormula("SUM(DO5:DO" + rowNum + ")");
            celula5118.setCellStyle(linha5WhiteOut);

            Cell celula5122 = linha5.createCell(122);
            celula5122.setCellValue("Total:");
            celula5122.setCellStyle(linha5WhiteIn);

            Cell celula5123 = linha5.createCell(123);
            celula5123.setCellFormula("SUM(DT5:DT" + rowNum + ")");
            celula5123.setCellStyle(linha5WhiteOut);

            Cell celula5127 = linha5.createCell(127);
            celula5127.setCellValue("Total:");
            celula5127.setCellStyle(linha5WhiteIn);

            Cell celula5128 = linha5.createCell(128);
            celula5128.setCellFormula("SUM(DY5:DY" + rowNum + ")");
            celula5128.setCellStyle(linha5WhiteOut);

            Cell celula5132 = linha5.createCell(132);
            celula5132.setCellValue("Total:");
            celula5132.setCellStyle(linha5WhiteIn);

            Cell celula5133 = linha5.createCell(133);
            celula5133.setCellFormula("SUM(ED5:ED" + rowNum + ")");
            celula5133.setCellStyle(linha5WhiteOut);

            Cell celula5138 = linha5.createCell(138);
            celula5138.setCellValue("Total:");
            celula5138.setCellStyle(linha5WhiteIn);

            Cell celula5139 = linha5.createCell(139);
          celula5139.setCellFormula("SUM(EJ5:EJ" + rowNum + ")");
            celula5139.setCellStyle(linha5WhiteOut);

            Cell celula5150 = linha5.createCell(150);
            celula5150.setCellFormula("SUM(EU5:EU" + rowNum + ")");
            celula5150.setCellStyle(linha5WhiteAll);

            Cell celula5151 = linha5.createCell(151);
            celula5151.setCellFormula("SUM(EV5:EV" + rowNum + ")");
            celula5151.setCellStyle(linha5WhiteAll);

            Cell celula5152 = linha5.createCell(152);
            celula5152.setCellFormula("SUM(EW5:EW" + rowNum + ")");
            celula5152.setCellStyle(linha5WhiteAll);

            Cell celula5153 = linha5.createCell(153);
            celula5153.setCellValue("Total:");
            celula5153.setCellStyle(linha5WhiteIn);

            Cell celula5154 = linha5.createCell(154);
            celula5154.setCellFormula("SUM(EY5:EY" + rowNum + ")");
            celula5154.setCellStyle(linha5WhiteOut);

            Cell celula5155 = linha5.createCell(155);
            celula5155.setCellValue("Total:");
            celula5155.setCellStyle(linha5WhiteIn);

            Cell celula5156 = linha5.createCell(156);
           celula5156.setCellFormula("SUM(FA5:FA" + rowNum + ")");
            celula5156.setCellStyle(linha5WhiteOut);

            Cell celula5157 = linha5.createCell(157);
            celula5157.setCellValue("Total:");
            celula5157.setCellStyle(linha5WhiteIn);

            Cell celula5158 = linha5.createCell(158);
           celula5158.setCellFormula("SUM(FC5:FC" + rowNum + ")");
            celula5158.setCellStyle(linha5WhiteOut);

            Cell celula5159 = linha5.createCell(159);
            celula5159.setCellValue("Total:");
            celula5159.setCellStyle(linha5WhiteIn);

            Cell celula5160 = linha5.createCell(160);
            celula5160.setCellFormula("SUM(FE5:FE" + rowNum + ")");
            celula5160.setCellStyle(linha5WhiteOut);

            Cell celula5161 = linha5.createCell(161);
            celula5161.setCellValue("Total:");
            celula5161.setCellStyle(linha5WhiteIn);

            Cell celula5162 = linha5.createCell(162);
           celula5162.setCellFormula("SUM(FG5:FG" + rowNum + ")");
            celula5162.setCellStyle(linha5WhiteOut);

            Cell celula5163 = linha5.createCell(163);
            celula5163.setCellValue("Total:");
            celula5163.setCellStyle(linha5WhiteIn);

            Cell celula5164 = linha5.createCell(164);
            celula5164.setCellFormula("SUM(FI5:FI" + rowNum + ")");
            celula5164.setCellStyle(linha5WhiteOut);

            Cell celula5165 = linha5.createCell(165);
            celula5165.setCellValue("Total:");
            celula5165.setCellStyle(linha5WhiteIn);

            Cell celula5166 = linha5.createCell(166);
            celula5166.setCellFormula("SUM(FK5:FK" + rowNum + ")");
            celula5166.setCellStyle(linha5WhiteOut);

            Cell celula5167 = linha5.createCell(167);
            celula5167.setCellValue("Total:");
            celula5167.setCellStyle(linha5WhiteIn);

            Cell celula5168 = linha5.createCell(168);
            celula5168.setCellFormula("SUM(FM5:FM" + rowNum + ")");
            celula5168.setCellStyle(linha5WhiteOut);

            Cell celula169 = linha5.createCell(169);
            celula169.setCellValue("Total:");
            celula169.setCellStyle(linha5WhiteIn);

            Cell celula5170 = linha5.createCell(170);
            celula5170.setCellFormula("SUM(FO5:FO" + rowNum + ")");
            celula5170.setCellStyle(linha5WhiteOut);

            Cell celula5171 = linha5.createCell(171);
            celula5171.setCellValue("Total:");
            celula5171.setCellStyle(linha5WhiteIn);

            Cell celula5172 = linha5.createCell(172);
            celula5172.setCellFormula("SUM(FQ5:FQ" + rowNum + ")");
            celula5172.setCellStyle(linha5WhiteOut);

            Cell celula5173 = linha5.createCell(173);
            celula5173.setCellValue("Total:");
            celula5173.setCellStyle(linha5WhiteIn);

            Cell celula5174 = linha5.createCell(174);
           celula5174.setCellFormula("SUM(FS5:FS" + rowNum + ")");
            celula5174.setCellStyle(linha5WhiteOut);

            Cell celula5175 = linha5.createCell(175);
            celula5175.setCellValue("Total:");
            celula5175.setCellStyle(linha5WhiteIn);

            Cell celula5176 = linha5.createCell(176);
            celula5176.setCellFormula("SUM(FU5:FU" + rowNum + ")");
            celula5176.setCellStyle(linha5WhiteOut);

            Cell celula5177 = linha5.createCell(177);
            celula5177.setCellValue("Total:");
            celula5177.setCellStyle(linha5WhiteIn);

            Cell celula5178 = linha5.createCell(178);
            celula5178.setCellFormula("SUM(FW5:FW" + rowNum + ")");
            celula5178.setCellStyle(linha5WhiteOut);

            Cell celula5179 = linha5.createCell(179);
            celula5179.setCellValue("Total:");
            celula5179.setCellStyle(linha5WhiteIn);

            Cell celula5180 = linha5.createCell(180);
            celula5180.setCellFormula("SUM(FY5:FY" + rowNum + ")");
            celula5180.setCellStyle(linha5WhiteOut);

            Cell celula5181 = linha5.createCell(181);
            celula5181.setCellValue("Total:");
            celula5181.setCellStyle(linha5WhiteIn);

            Cell celula5182 = linha5.createCell(182);
            celula5182.setCellFormula("SUM(GA5:GA" + rowNum + ")");
            celula5182.setCellStyle(linha5WhiteOut);

            Cell celula5183 = linha5.createCell(183);
            celula5183.setCellValue("Total:");
            celula5183.setCellStyle(linha5WhiteIn);

            Cell celula5184 = linha5.createCell(184);
            celula5184.setCellFormula("SUM(GC5:GC" + rowNum + ")");
            celula5184.setCellStyle(linha5WhiteOut);

            Cell celula5185 = linha5.createCell(185);
            celula5185.setCellValue("Total:");
            celula5185.setCellStyle(linha5WhiteIn);

            Cell celula5186 = linha5.createCell(186);
            celula5186.setCellFormula("SUM(GE5:GE" + rowNum + ")");
            celula5186.setCellStyle(linha5WhiteOut);

            Cell celula5187 = linha5.createCell(187);
            celula5187.setCellValue("Total:");
            celula5187.setCellStyle(linha5WhiteIn);

            Cell celula5188 = linha5.createCell(188);
            celula5188.setCellFormula("SUM(GG5:GG" + rowNum + ")");
            celula5188.setCellStyle(linha5WhiteOut);

            Cell celula5189 = linha5.createCell(189);
            celula5189.setCellValue("Total:");
            celula5189.setCellStyle(linha5WhiteIn);

            Cell celula5190 = linha5.createCell(190);
           celula5190.setCellFormula("SUM(FI5:FI" + rowNum + ")");
            celula5190.setCellStyle(linha5WhiteOut);

            Cell celula5191 = linha5.createCell(191);
            celula5191.setCellValue("Total:");
            celula5191.setCellStyle(linha5WhiteIn);

            Cell celula5192 = linha5.createCell(192);
            celula5192.setCellFormula("SUM(GK5:GK" + rowNum + ")");
            celula5192.setCellStyle(linha5WhiteOut);

            Cell celula5193 = linha5.createCell(193);
           celula5193.setCellFormula("SUM(GL5:GL" + rowNum + ")");
            celula5193.setCellStyle(linha5WhiteOut);

            Cell celula5196 = linha5.createCell(196);
            celula5196.setCellFormula("SUM(GO5:GO" + rowNum + ")");
            celula5196.setCellStyle(linha5WhiteAll);

            Cell celula5199 = linha5.createCell(199);
            celula5199.setCellFormula("SUM(GR5:GR" + rowNum + ")");
            celula5199.setCellStyle(linha5WhiteAll);

            Cell celula5200 = linha5.createCell(200);
            celula5200.setCellFormula("SUM(GS5:GS" + rowNum + ")");
            celula5200.setCellStyle(linha5WhiteOut);

            Cell celula5201 = linha5.createCell(201);
            celula5201.setCellFormula("SUM(GT5:GT" + rowNum + ")");
            celula5201.setCellStyle(linha5WhiteOut);

            Cell celula5202 = linha5.createCell(202);
            celula5202.setCellFormula("SUM(GU5:GU" + rowNum + ")");
            celula5202.setCellStyle(linha5WhiteOut);

            Cell celula5203 = linha5.createCell(203);
            celula5203.setCellFormula("SUM(GV5:GV" + rowNum + ")");
            celula5203.setCellStyle(linha5WhiteOut);

            Cell celula5204 = linha5.createCell(204);
            celula5204.setCellFormula("SUM(GW5:GW" + rowNum + ")");
            celula5204.setCellStyle(linha5WhiteOut);

            Cell celula5205 = linha5.createCell(205);
            celula5205.setCellFormula("SUM(GX5:GX" + rowNum + ")");
            celula5205.setCellStyle(linha5WhiteOut);

            Cell celula5206 = linha5.createCell(206);
            celula5206.setCellFormula("SUM(GY5:GY" + rowNum + ")");
            celula5206.setCellStyle(linha5WhiteOut);

            Cell celula5207 = linha5.createCell(207);
            celula5207.setCellFormula("SUM(GZ5:GZ" + rowNum + ")");
            celula5207.setCellStyle(linha5WhiteOut);

            Cell celula5208 = linha5.createCell(208);
            celula5208.setCellFormula("SUM(HA5:HA" + rowNum + ")");
            celula5208.setCellStyle(linha5WhiteOut);

            Cell celula5209 = linha5.createCell(209);
            celula5209.setCellFormula("SUM(HB5:HB" + rowNum + ")");
            celula5209.setCellStyle(linha5WhiteOut);

            Cell celula5210 = linha5.createCell(210);
            celula5210.setCellFormula("SUM(HC5:HC" + rowNum + ")");
            celula5210.setCellStyle(linha5WhiteOut);

            Cell celula5211 = linha5.createCell(211);
            celula5211.setCellFormula("SUM(HD5:HD" + rowNum + ")");
            celula5211.setCellStyle(linha5WhiteOut);

            Cell celula5212 = linha5.createCell(212);
            celula5212.setCellFormula("SUM(HE5:HE" + rowNum + ")");
            celula5212.setCellStyle(linha5WhiteOut);

            Cell celula5213 = linha5.createCell(213);
            celula5213.setCellFormula("SUM(HF5:HF" + rowNum + ")");
            celula5213.setCellStyle(linha5WhiteOut);

            Cell celula5214 = linha5.createCell(214);
             celula5214.setCellFormula("SUM(HG5:HG" + rowNum + ")");
            celula5214.setCellStyle(linha5WhiteOut);

            Cell celula5215 = linha5.createCell(215);
             celula5215.setCellFormula("SUM(HH5:HH" + rowNum + ")");
            celula5215.setCellStyle(linha5WhiteOut);

            Cell celula5216 = linha5.createCell(216);
            celula5216.setCellFormula("SUM(HI5:HI" + rowNum + ")");
            celula5216.setCellStyle(linha5WhiteOut);

            Cell celula5217 = linha5.createCell(217);
            celula5217.setCellFormula("SUM(HJ5:HJ" + rowNum + ")");
            celula5217.setCellStyle(linha5WhiteOut);

            Cell celula5218 = linha5.createCell(218);
            celula5218.setCellFormula("SUM(HK5:HK" + rowNum + ")");
            celula5218.setCellStyle(linha5WhiteOut);

            Cell celula5219 = linha5.createCell(219);
            celula5219.setCellFormula("SUM(HL5:HL" + rowNum + ")");
            celula5219.setCellStyle(linha5WhiteOut);

            Cell celula5220 = linha5.createCell(220);
            celula5220.setCellFormula("SUM(HM5:HM" + rowNum + ")");
            celula5220.setCellStyle(linha5WhiteOut);

            Cell celula5221 = linha5.createCell(221);
            celula5221.setCellFormula("SUM(HN5:HN" + rowNum + ")");
            celula5221.setCellStyle(linha5WhiteOut);

            Cell celula5222 = linha5.createCell(222);
            celula5222.setCellFormula("SUM(HO5:HO" + rowNum + ")");
            celula5222.setCellStyle(linha5WhiteOut);

            Cell celula5223 = linha5.createCell(223);
            celula5223.setCellFormula("SUM(HP5:HP" + rowNum + ")");
            celula5223.setCellStyle(linha5WhiteOut);

            Cell celula5224 = linha5.createCell(224);
            celula5224.setCellFormula("SUM(HQ5:HQ" + rowNum + ")");
            celula5224.setCellStyle(linha5WhiteOut);

            Cell celula5225 = linha5.createCell(225);
            celula5225.setCellFormula("SUM(HR5:HR" + rowNum + ")");
            celula5225.setCellStyle(linha5WhiteOut);

            Cell celula5226 = linha5.createCell(226);
            celula5226.setCellFormula("SUM(HS5:HS" + rowNum + ")");
            celula5226.setCellStyle(linha5WhiteOut);

            Cell celula5227 = linha5.createCell(227);
            celula5227.setCellFormula("SUM(HT5:HT" + rowNum + ")");
            celula5227.setCellStyle(linha5WhiteOut);

            Cell celula5228 = linha5.createCell(228);
            celula5228.setCellFormula("SUM(HU5:HU" + rowNum + ")");
            celula5228.setCellStyle(linha5WhiteOut);

            Cell celula5229 = linha5.createCell(229);
            celula5229.setCellFormula("SUM(HV5:HV" + rowNum + ")");
            celula5229.setCellStyle(linha5WhiteOut);

            Cell celula5230 = linha5.createCell(230);
            celula5230.setCellFormula("SUM(HW5:HW" + rowNum + ")");
            celula5230.setCellStyle(linha5WhiteOut);

            Cell celula5231 = linha5.createCell(231);
            celula5231.setCellFormula("SUM(HX5:HX" + rowNum + ")");
            celula5231.setCellStyle(linha5WhiteOut);

            Cell celula5232 = linha5.createCell(232);
            celula5232.setCellFormula("SUM(HY5:HY" + rowNum + ")");
            celula5232.setCellStyle(linha5WhiteOut);

            Cell celula5233 = linha5.createCell(233);
            celula5233.setCellFormula("SUM(HZ5:HZ" + rowNum + ")");
            celula5233.setCellStyle(linha5WhiteOut);

            Cell celula5234 = linha5.createCell(234);
             celula5234.setCellFormula("SUM(IA5:IA" + rowNum + ")");
            celula5234.setCellStyle(linha5WhiteOut);

            Cell celula5235 = linha5.createCell(235);
            celula5235.setCellFormula("SUM(IB5:IB" + rowNum + ")");
            celula5235.setCellStyle(linha5WhiteOut);

            Cell celula5236 = linha5.createCell(236);
            celula5236.setCellFormula("SUM(IC5:IC" + rowNum + ")");
            celula5236.setCellStyle(linha5WhiteOut);

            Cell celula5237 = linha5.createCell(237);
            celula5237.setCellFormula("SUM(ID5:ID" + rowNum + ")");
            celula5237.setCellStyle(linha5WhiteOut);

            Cell celula5238 = linha5.createCell(238);
            celula5238.setCellFormula("SUM(IE5:IE" + rowNum + ")");
            celula5238.setCellStyle(linha5WhiteOut);

            Cell celula5239 = linha5.createCell(239);
             celula5239.setCellFormula("SUM(IF5:IF" + rowNum + ")");
            celula5239.setCellStyle(linha5WhiteOut);

            Cell celula5240 = linha5.createCell(240);
            celula5240.setCellFormula("SUM(IG5:IG" + rowNum + ")");
            celula5240.setCellStyle(linha5WhiteOut);

            Cell celula5241 = linha5.createCell(241);
            celula5241.setCellFormula("SUM(IH5:IH" + rowNum + ")");
            celula5241.setCellStyle(linha5WhiteOut);

            Cell celula5242 = linha5.createCell(242);
            celula5242.setCellFormula("SUM(IG5:II" + rowNum + ")");
            celula5242.setCellStyle(linha5WhiteOut);

            Cell celula5243 = linha5.createCell(243);
            celula5243.setCellFormula("SUM(IJ5:IJ" + rowNum + ")");
            celula5243.setCellStyle(linha5WhiteOut);

            Cell celula5244 = linha5.createCell(244);
            celula5244.setCellFormula("SUM(IK5:IK" + rowNum + ")");
            celula5244.setCellStyle(linha5WhiteOut);

            Cell celula5245 = linha5.createCell(245);
            celula5245.setCellFormula("SUM(IL5:IL" + rowNum + ")");
            celula5245.setCellStyle(linha5WhiteOut);

            Cell celula5246 = linha5.createCell(246);
            celula5246.setCellFormula("SUM(IM5:IM" + rowNum + ")");
            celula5246.setCellStyle(linha5WhiteOut);

            Cell celula5247 = linha5.createCell(247);
            celula5247.setCellFormula("SUM(IN5:IN" + rowNum + ")");
            celula5247.setCellStyle(linha5WhiteOut);

            Cell celula5248 = linha5.createCell(248);
             celula5248.setCellFormula("SUM(IO5:IO" + rowNum + ")");
            celula5248.setCellStyle(linha5WhiteOut);

            Cell celula5249 = linha5.createCell(249);
            celula5249.setCellFormula("SUM(IN5:IP" + rowNum + ")");
            celula5249.setCellStyle(linha5WhiteOut);

            Cell celula5250 = linha5.createCell(250);
            celula5250.setCellFormula("SUM(IQ5:IQ" + rowNum + ")");
            celula5250.setCellStyle(linha5WhiteOut);

            // incrementamos a linha
            Row5++;
            
            // Adiciona filtragem de dados...
            sheet.createFreezePane(0, 4, 0, 1);

            // Dimensionar todas as colunas e algumas específicas;
            for (int i = 0; i < arrayList.size(); i++) {

                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(1, 3000);
                sheet.setColumnWidth(2, 3000);
                sheet.setColumnWidth(3, 3000);
                sheet.setColumnWidth(4, 3000);
                sheet.setColumnWidth(5, 3000);
                sheet.setColumnWidth(6, 3000);
                sheet.setColumnWidth(8, 3000);
                sheet.setColumnWidth(9, 3000);
                sheet.setColumnWidth(33, 4000);
                sheet.setColumnWidth(34, 4000);
                sheet.setColumnWidth(38, 3000);
                sheet.setColumnWidth(46, 3000);
                sheet.setColumnWidth(51, 3000);
                sheet.setColumnWidth(56, 3000);
                sheet.setColumnWidth(58, 3000);
                sheet.setColumnWidth(63, 3000);
                sheet.setColumnWidth(68, 3000);
                sheet.setColumnWidth(73, 3000);
                sheet.setColumnWidth(78, 3000);
                sheet.setColumnWidth(81, 3000);
                sheet.setColumnWidth(86, 3000);
                sheet.setColumnWidth(91, 3000);
                sheet.setColumnWidth(96, 3000);
                sheet.setColumnWidth(101, 3000);
                sheet.setColumnWidth(106, 3000);
                sheet.setColumnWidth(111, 3000);
                sheet.setColumnWidth(116, 3000);
                sheet.setColumnWidth(121, 3000);
                sheet.setColumnWidth(126, 3000);
                sheet.setColumnWidth(131, 3000);
                sheet.setColumnWidth(136, 3000);
                sheet.setColumnWidth(137, 3000);
                sheet.setColumnWidth(149, 3000);
                sheet.setColumnWidth(150, 3000);
                sheet.setColumnWidth(151, 3000);
                sheet.setColumnWidth(152, 3000);
                sheet.setColumnWidth(153, 3000);
                sheet.setColumnWidth(155, 3000);
                sheet.setColumnWidth(157, 3000);
                sheet.setColumnWidth(159, 3000);
                sheet.setColumnWidth(161, 3000);
                sheet.setColumnWidth(163, 3000);
                sheet.setColumnWidth(165, 3000);
                sheet.setColumnWidth(167, 3000);
                sheet.setColumnWidth(169, 3000);
                sheet.setColumnWidth(171, 3000);
                sheet.setColumnWidth(173, 3000);
                sheet.setColumnWidth(175, 3000);
                sheet.setColumnWidth(177, 3000);
                sheet.setColumnWidth(179, 3000);
                sheet.setColumnWidth(181, 3000);
                sheet.setColumnWidth(183, 3000);
                sheet.setColumnWidth(185, 3000);
                sheet.setColumnWidth(187, 3000);
                sheet.setColumnWidth(189, 3000);
                sheet.setColumnWidth(191, 3000);
                sheet.setColumnWidth(193, 3000);
                sheet.setColumnWidth(194, 3000);
                sheet.setColumnWidth(195, 3000);
                sheet.setColumnWidth(196, 3000);
                sheet.setColumnWidth(197, 3000);
                sheet.setColumnWidth(198, 3000);
                sheet.setColumnWidth(150, 3000);

            }
            
            //Adiciona filtro na linha 4, filtrando os dados  coletados do banco de dados
            sheet.setAutoFilter(new CellRangeAddress(3, rowNum, 1, 250));

            // ESCREVER SAíDA em um ARQUIVO
            SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yy");
            String data1 = dateFormat.format(dataVe1);
            String data2 = dateFormat.format(dataVe2);
           
            long hora = new Date().getHours();
            long minuto = new Date().getMinutes();
            
            try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\user\\Desktop\\PF(" + data1 + " - " + data2 + ").xlsx")) {
                System.out.println("Excel de Fundicao gerado com sucesso");
                System.out.println("Processo Finalizado... " + hora + ":" + minuto + ".");
                workbook.write(fileOut);
                // FECHAR MÉTODO
            }
        }
    }
}
