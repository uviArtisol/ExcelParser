package com.artificialsolutions.excelparser

class Main {

    static void main(String[] args) {
        def filename = args[0]
        println(ExcelParser.parse(filename, true, true, true));

    }
}
