package cn.dlysxx.www.dlysxx.gongzi.service;

import org.springframework.web.multipart.MultipartFile;

public interface CsvService {

    void readExcelData(MultipartFile multipartFile);
}
