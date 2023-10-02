package com.example.idxdownloader;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.shell.command.annotation.EnableCommand;

@SpringBootApplication
@EnableCommand(MyCommands.class)
public class IdxDownloaderApplication {

    public static void main(String[] args) {
        SpringApplication.run(IdxDownloaderApplication.class, args);
    }

}
