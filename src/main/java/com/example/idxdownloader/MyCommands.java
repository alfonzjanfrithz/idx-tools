package com.example.idxdownloader;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.shell.command.annotation.Command;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;


@Command
public class MyCommands {
    @Autowired
    private FinancialStatementService financialStatementService;

    @Autowired
    private FileDownloadService fileDownloadService;

    //example --year 2023 --periode tw2 --kodeEmiten ANJT
    @Command(command = "example", description = "some example of the command will be shown in help")
    public void example(int year, String periode, String kodeEmiten) {
        ApiResponse apiResponse = financialStatementService.fetchData(year, periode, kodeEmiten);
        List<Attachment> attachmentExcel = filterAttachmentsByFileType(apiResponse, "xlsx");
        Optional<String> link = attachmentExcel.stream().map(attachment -> "https://idx.co.id" + attachment.getFilePath()).findAny();
        System.out.println(link);

        fileDownloadService.downloadFile(link.get());
    }

    public List<Attachment> filterAttachmentsByFileType(ApiResponse apiResponse, String fileType) {
        return apiResponse.getResults().stream()
                .flatMap(result -> result.getAttachments().stream())
                .filter(attachment -> (attachment.getFileType().contains(fileType)))
                .collect(Collectors.toList());
    }
}