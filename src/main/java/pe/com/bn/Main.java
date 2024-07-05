package pe.com.bn;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;

public class Main {

    public static void main(String[] args) {


        String excelFilePath = "D:\\repo.xlsx";
        String jdbcUrl = "jdbc:oracle:thin:@//10.7.12.177:1521/orades";
        String jdbcUser = "bn_msds";
        String jdbcPassword = "bn_msds";

        try (FileInputStream inputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream);
             Connection connection = DriverManager.getConnection(jdbcUrl, jdbcUser, jdbcPassword)) {

            Sheet sheet = workbook.getSheetAt(0);
            String sql = "INSERT INTO BNMSDSF01_PROGRAMAS (" +
                    "ID, F01_CODSISTEMA, F01_NOMSISTEMA, F01_DESSISTEMA, F01_NIVEL, " +
                    "F01_AREAUSUARIA, F01_AREARESPONSABLE, F01_AREAPROCESO, F01_AREATIPO, " +
                    "F01_PLATAFORMA, F01_RTO, F01_MTDP, F01_RPO, F01_MTDL, F01_MBCO, " +
                    "F01_VERSION, F01_VERSIONDESA, F01_URLDESARROLLO, F01_URLCERTIFICACION, " +
                    "F01_URLPRODUCCION, F01_URLGIT, F01_LENGUAJE, F01_BASEDATOS, F01_PERSONARESP, " +
                    "F01_DESAREALIZADO, F01_PLATAFORMACOM, F01_REQUISITOS, F01_FECHA, F01_COMENTARIOS" +
                    ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            try (PreparedStatement preparedStatement = connection.prepareStatement(sql)) {
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue; // Saltar la primera fila si es el encabezado
                    }

                    // ID
                    if (isCellEmpty(row.getCell(0))) {
                        preparedStatement.setNull(1, java.sql.Types.INTEGER);
                    } else {
                        preparedStatement.setInt(1, (int) row.getCell(0).getNumericCellValue());
                    }

                    // F01_CODSISTEMA
                    if (isCellEmpty(row.getCell(1))) {
                        preparedStatement.setNull(2, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(2, row.getCell(1).getStringCellValue());
                    }

                    // F01_NOMSISTEMA
                    if (isCellEmpty(row.getCell(2))) {
                        preparedStatement.setNull(3, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(3, row.getCell(2).getStringCellValue());
                    }

                    // F01_DESSISTEMA
                    if (isCellEmpty(row.getCell(3))) {
                        preparedStatement.setNull(4, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(4, row.getCell(3).getStringCellValue());
                    }

                    // F01_NIVEL
                    if (isCellEmpty(row.getCell(4))) {
                        preparedStatement.setNull(5, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(5, row.getCell(4).getStringCellValue());
                    }

                    // F01_AREAUSUARIA
                    if (isCellEmpty(row.getCell(5))) {
                        preparedStatement.setNull(6, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(6, row.getCell(5).getStringCellValue());
                    }

                    // F01_AREARESPONSABLE
                    if (isCellEmpty(row.getCell(6))) {
                        preparedStatement.setNull(7, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(7, row.getCell(6).getStringCellValue());
                    }

                    // F01_AREAPROCESO
                    if (isCellEmpty(row.getCell(7))) {
                        preparedStatement.setNull(8, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(8, row.getCell(7).getStringCellValue());
                    }

                    // F01_AREATIPO
                    if (isCellEmpty(row.getCell(8))) {
                        preparedStatement.setNull(9, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(9, row.getCell(8).getStringCellValue());
                    }

                    // F01_PLATAFORMA
                    if (isCellEmpty(row.getCell(9))) {
                        preparedStatement.setNull(10, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(10, row.getCell(9).getStringCellValue());
                    }

                    // F01_RTO
                    if (isCellEmpty(row.getCell(10))) {
                        preparedStatement.setNull(11, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(11, row.getCell(10).getStringCellValue());
                    }

                    // F01_MTDP
                    if (isCellEmpty(row.getCell(11))) {
                        preparedStatement.setNull(12, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(12, row.getCell(11).getStringCellValue());
                    }

                    // F01_RPO
                    if (isCellEmpty(row.getCell(12))) {
                        preparedStatement.setNull(13, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(13, row.getCell(12).getStringCellValue());
                    }

                    // F01_MTDL
                    if (isCellEmpty(row.getCell(13))) {
                        preparedStatement.setNull(14, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(14, row.getCell(13).getStringCellValue());
                    }

                    // F01_MBCO
                    if (isCellEmpty(row.getCell(14))) {
                        preparedStatement.setNull(15, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(15, row.getCell(14).getStringCellValue());
                    }

                    // F01_VERSION
                    if (isCellEmpty(row.getCell(15))) {
                        preparedStatement.setNull(16, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(16, row.getCell(15).getStringCellValue());
                    }

                    // F01_VERSIONDESA
                    if (isCellEmpty(row.getCell(16))) {
                        preparedStatement.setNull(17, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(17, row.getCell(16).getStringCellValue());
                    }

                    // F01_URLDESARROLLO
                    if (isCellEmpty(row.getCell(17))) {
                        preparedStatement.setNull(18, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(18, row.getCell(17).getStringCellValue());
                    }

                    // F01_URLCERTIFICACION
                    if (isCellEmpty(row.getCell(18))) {
                        preparedStatement.setNull(19, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(19, row.getCell(18).getStringCellValue());
                    }

                    // F01_URLPRODUCCION
                    if (isCellEmpty(row.getCell(19))) {
                        preparedStatement.setNull(20, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(20, row.getCell(19).getStringCellValue());
                    }

                    // F01_URLGIT
                    if (isCellEmpty(row.getCell(20))) {
                        preparedStatement.setNull(21, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(21, row.getCell(20).getStringCellValue());
                    }

                    // F01_LENGUAJE
                    if (isCellEmpty(row.getCell(21))) {
                        preparedStatement.setNull(22, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(22, row.getCell(21).getStringCellValue());
                    }

                    // F01_BASEDATOS
                    if (isCellEmpty(row.getCell(22))) {
                        preparedStatement.setNull(23, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(23, row.getCell(22).getStringCellValue());
                    }

                    // F01_PERSONARESP
                    if (isCellEmpty(row.getCell(23))) {
                        preparedStatement.setNull(24, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(24, row.getCell(23).getStringCellValue());
                    }

                    // F01_DESAREALIZADO
                    if (isCellEmpty(row.getCell(24))) {
                        preparedStatement.setNull(25, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(25, row.getCell(24).getStringCellValue());
                    }

                    // F01_ARQUITECTURA
                    if (isCellEmpty(row.getCell(25))) {
                        preparedStatement.setNull(26, Types.BLOB);
                    } else {
                        preparedStatement.setString(26, row.getCell(25).getStringCellValue());
                    }

                    // F01_PLATAFORMACOM
                    if (isCellEmpty(row.getCell(26))) {
                        preparedStatement.setNull(27, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(27, row.getCell(26).getStringCellValue());
                    }

                    // F01_REQUISITOS
                    if (isCellEmpty(row.getCell(27))) {
                        preparedStatement.setNull(28, java.sql.Types.VARCHAR);
                    } else {
                        preparedStatement.setString(28, row.getCell(27).getStringCellValue());
                    }

                    // F01_FECHA
                    if (isCellEmpty(row.getCell(28))) {
                        preparedStatement.setNull(29, java.sql.Types.DATE);
                    } else {
                        preparedStatement.setString(29, row.getCell(28).getStringCellValue());
                    }





                    preparedStatement.addBatch();
                }

                preparedStatement.executeBatch();
                System.out.println("Datos insertados correctamente en la base de datos.");
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }


    }
    private static boolean isCellEmpty(Cell cell) {
        return (cell == null || cell.getCellType() == CellType.BLANK);
    }
}