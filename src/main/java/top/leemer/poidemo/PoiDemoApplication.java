package top.leemer.poidemo;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import top.leemer.poidemo.excel.PoiReadExcel;
import top.leemer.poidemo.excel.PoiWriteExcel;

@SpringBootApplication
public class PoiDemoApplication implements CommandLineRunner {

    public static void main(String[] args) {
        SpringApplication.run(PoiDemoApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        PoiWriteExcel.run();
    }
}
