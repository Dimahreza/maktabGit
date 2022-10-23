import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RetiredGeneralReportExcelDto {

    private String personCode;
    private String nationalCode;
    private String fullName;
    private Double membershipFee;
    private Integer StatusCode;
    private String statusDescription;
    private Integer totalCount;
    private Double sumMembershipFee;
    private Integer payCount;
    private Integer notPayCount;
    
}
