import java.util.List;

public interface IPersonnelRepository {

    List<RetiredGeneralReportExcelDto> getRetiredPersonnel(String statusCode, String year, String monh);
}
