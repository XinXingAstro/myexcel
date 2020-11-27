package com.xinxing.rename;

import java.io.File;

// 修改文件夹名、文件名及路径
public class EasyRename {
    public static void main(String[] args) {
        String path = "D:\\Repositories\\Coding-Pool\\Leetcode\\src";
        File root = new File(path);
        File[] listFiles = root.listFiles();
        for (File f : listFiles) {
            System.out.println(f.getName());
//            f.renameTo(new File(path + "\\src\\"+ f.getName()));
        }
    }
}
