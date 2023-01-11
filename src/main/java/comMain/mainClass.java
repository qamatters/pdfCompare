package comMain;

import helper.commonHelper;

import java.io.IOException;

public class mainClass {
    public static void main(String[] args) throws IOException {
            commonHelper helper = new commonHelper();
            helper.getFileNameAndCompare();
    }
}
