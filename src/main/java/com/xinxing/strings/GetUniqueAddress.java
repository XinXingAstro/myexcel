package com.xinxing.strings;

import java.util.HashSet;
import java.util.Scanner;
import java.util.Set;

public class GetUniqueAddress {
    private Set<String> set;
    public GetUniqueAddress() {
    }

    public void getUniqueAddrss() {
        set = new HashSet<String>();
        System.out.println("Input:...");
        Scanner sc = new Scanner(System.in);
        while (!sc.hasNext("0")) {
            String str = sc.nextLine();
//            System.out.println(str);
//            System.out.println("Format: " + str.trim());
            set.add(str.trim());
        }
        int len = set.size();
        System.out.println("Number of Address: " + len);
        for (String s : set) {
            System.out.println(s);
        }
    }

    public static void main(String[] args) {
        GetUniqueAddress gua = new GetUniqueAddress();
        gua.getUniqueAddrss();
    }
}
