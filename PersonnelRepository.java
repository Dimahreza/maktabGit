import org.springframework.beans.factory.annotation.Autowired;

import javax.sql.DataSource;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

public class PersonnelRepository extends IPersonnelRepository{
    @Autowired
    private DataSource dataSource;
    @Override
    public List<RetiredGeneralReportExcelDto> getRetiredPersonnel(String statusCode, String year, String monh) {
       try(Connection connection = dataSource.getConnection()) {
           String query = "select * from rts.excel_general_report_detail";
           CallableStatement statement =  connection.prepareCall(query);
           statement.setString(1,statusCode);
           statement.setString(2,"04");
           statement.setString(3,"1400");
           ResultSet rs = statement.executeQuery();

           List<RetiredGeneralReportExcelDto> list = new ArrayList<>();
           while (rs.next()){
               RetiredGeneralReportExcelDto dto =  new RetiredGeneralReportExcelDto();

           }
           return List;
       }
    }
}
