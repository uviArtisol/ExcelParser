package com.artificialsolutions.excelparser

class FileGetter {
    static def getFile(String sResourceFileName) {
        return sResourceFileName
//        URL url = FileGetter.class.getClassLoader().getResource(sResourceFileName);
//        if (url == null) {
//            throw new RuntimeException('Failure to obtain URL for file [' + sResourceFileName + '] in Resources');
//        }
//        URI uri;
//        try {
//            uri = url.toURI();
//        } catch (exc) {
//            throw new RuntimeException('Failure to create URI for URL [' + url + '] for file [' + sResourceFileName + '] in Resources', exc);
//        }
//        File file;
//        try {
//            return new File(uri);
//        } catch (exc) {
//            throw new RuntimeException('Failure to create File object for URL [' + url + '] and URI [' + uri + '] for file [' + sResourceFileName + '] in Resources', exc);
//        }
    }
}
