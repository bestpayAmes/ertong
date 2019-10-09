package com.ames;

import java.io.File;
import java.util.function.Predicate;

public enum TypeAndCheckEnum {
    Routineurine("尿常规", file -> {
        return true;
    }),
    Vision("视力", file -> {
        return true;
    }),
    Hearing("听力", file -> true),
    Physical("体检", file -> {
        return true;
    }),
    Blood("血常规", file -> {
        return true;
    });

    private String key;

    private Predicate<File> checkOption;

    TypeAndCheckEnum(String key, Predicate<File> checkOption) {
        this.key = key;
        this.checkOption = checkOption;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public Predicate<File> getCheckOption() {
        return checkOption;
    }

    public void setCheckOption(Predicate<File> checkOption) {
        this.checkOption = checkOption;
    }

    public static String getTypeByFileName(String name){
        for(TypeAndCheckEnum key:TypeAndCheckEnum.values()){
            if(name.contains(key.getKey())){
                return key.getKey();
            }
        }
        throw new RuntimeException("错误的文件"+name);
    }
}
