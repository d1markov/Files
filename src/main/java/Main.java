import java.io.IOException;

public class Main {



    public static void main(String[] args) throws IOException {




        String filePath = "C:\\Users\\d.markov\\Desktop\\testData\\excel\\2.xls";
        FilesExcel files = new FilesExcel();
        files.writeIntoExcel(filePath);



    }
}
