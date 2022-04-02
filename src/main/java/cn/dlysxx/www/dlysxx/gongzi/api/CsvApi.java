package cn.dlysxx.www.dlysxx.gongzi.api;

import cn.dlysxx.www.dlysxx.gongzi.service.CsvService;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class CsvApi {

    private final CsvService csvService;

    public CsvApi(CsvService csvService) {
        this.csvService = csvService;
    }

    @RequestMapping(
        method = RequestMethod.POST,
        value = "/v1/csv/gongzi",
        produces = {"application/json"}
    )
    public void test(@RequestParam(value = "excelFile") MultipartFile multipartFile) {
        csvService.readExcelData(multipartFile);
    }
}
