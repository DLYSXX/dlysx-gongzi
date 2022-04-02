package cn.dlysxx.www.dlysxx.gongzi.service.impl;

import cn.dlysxx.www.common.date.DateUtil;
import cn.dlysxx.www.common.string.StringUtil;
import cn.dlysxx.www.dlysxx.gongzi.api.CsvApiDelegate;
import cn.dlysxx.www.dlysxx.gongzi.model.CsvApiCommand;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

/**
 * Csv Service.
 *
 * @author shuai
 */
@Service
public class CsvServiceImpl implements CsvApiDelegate {

    @Value("${spring.mail.username}")
    private String from;

    private final JavaMailSender javaMailSender;

    private static final String HEADER_RANGE = "A1:S4";

    private static final int START_ROW = 5;

    private static final String USERNAME_INDEX = "B";

    private static final String EMAIL_INDEX = "T";

    private static final String HTML_MAIL_TEMPLATE = ".html";

    private static final Logger LOGGER = LoggerFactory.getLogger(CsvServiceImpl.class);

    public CsvServiceImpl(JavaMailSender javaMailSender) {
        this.javaMailSender = javaMailSender;
    }

    @Override
    public ResponseEntity<Void> postV1CsvGongzi(CsvApiCommand csvApiCommand) {
        try (InputStream inputStream = new FileInputStream(csvApiCommand.getCsvFilePath())) {
            Workbook workbook = new Workbook();
            workbook.loadFromStream(inputStream);
            Worksheet sheet = workbook.getWorksheets().get(0);
            int lastColumn = sheet.getLastColumn();

            final String ym = DateUtil.toString(LocalDateTime.now(), DateUtil.UUUUMM);
            String backupPath = csvApiCommand.getOutputPath() + ym;
            if (!Files.isDirectory(Paths.get(backupPath))) {
                Files.createDirectories(Paths.get(backupPath));
            }

            for (int i = START_ROW; i < lastColumn - 1; i++) {
                // get current record's user info
                final String emailAddress = sheet.getRange().get(EMAIL_INDEX + i).getValue();
                final String userName = sheet.getRange().get(USERNAME_INDEX + i).getValue();
                final String emailTemplate = csvApiCommand.getOutputPath() + userName + HTML_MAIL_TEMPLATE;
                final String emailTemplateBackup = backupPath + "/" + userName + HTML_MAIL_TEMPLATE;
                if (StringUtils.hasLength(userName) || StringUtils.hasLength(emailAddress)) {
                    continue;
                }
                LOGGER.info("Current user info -----> username is {}, email is {}", userName, emailAddress);
                // create temp excel for copy details
                Workbook gongziCsvFile = new Workbook();
                Worksheet gongziCsvSheet = gongziCsvFile.getWorksheets().get(0);
                // copy header
                sheet.copy(sheet.getCellRange(HEADER_RANGE),
                    gongziCsvSheet.getCellRange(HEADER_RANGE), true);
                // copy gongzi details
                sheet.copy(sheet.getCellRange("A" + i + ":S" + i),
                    gongziCsvSheet.getCellRange("A" + START_ROW + ":S" + START_ROW), true);
                // save to html template
                gongziCsvSheet.saveToHtml(emailTemplate);

                // send mail
                this.sendMail(userName, emailAddress, emailTemplate, emailTemplateBackup);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return new ResponseEntity(HttpStatus.OK);
    }

    /**
     * Send Mail.
     *
     * @param userName user name
     * @param emailAddress email address
     * @param emailTemplate email template
     * @param emailTemplateBackup email template backup
     */
    private void sendMail(String userName, String emailAddress, String emailTemplate, String emailTemplateBackup) {
        try {
            this.sender(emailAddress, emailTemplate);
            // backup email template
            Files.move(Paths.get(emailTemplate), Paths.get(emailTemplateBackup),
                StandardCopyOption.REPLACE_EXISTING);
            Thread.sleep(500);
        } catch (Exception ex) {
            LOGGER.error("Send email failed!!! username is {}, email is {}", userName, emailAddress, ex);
        }
    }

    /**
     * Mail sender.
     *
     * @param email email address
     * @param filePath mail template path
     * @throws Exception exception
     */
    private void sender(String email, String filePath) throws Exception {
        javaMailSender.send(mimeMessage -> {
            try (FileInputStream fileInputStream = new FileInputStream(filePath)) {
                MimeMessageHelper helper = new MimeMessageHelper(mimeMessage, StandardCharsets.UTF_8.name());
                helper.setFrom(from);
                helper.setTo(email);
                helper.setSubject(DateUtil.toString(LocalDateTime.now(), DateUtil.UUUUMM) + "工资单");
                helper.setText(StringUtil.conversion(fileInputStream), true);
            }
        });
        LOGGER.info("Send email succeed -----> email is {}", email);
    }
}
